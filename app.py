# MAT-TELETRABALHO-001 · app.py
# Painel de Teletrabalho TJ/MA — visualização + entrada de dados
# Fonte única: planilha "Teletrabalho TJ/MA" (Google Sheets)
# COGEX-MA/TJMA

from __future__ import annotations

import io
import datetime as dt

import pandas as pd
import plotly.express as px
import streamlit as st

import sheets_io as sio
from parser_produtividade import parse_produtividade

# ───────────────────────────── Config geral ──────────────────────────────
st.set_page_config(
    page_title="Painel de Teletrabalho — TJ/MA",
    page_icon="🏛️",
    layout="wide",
)

AZUL = "#1B3A6B"
AZUL_CLARO = "#3D6DB5"
CINZA = "#5A6472"
CINZA_CLARO = "#F2F3F5"
VERDE = "#1E8F5F"
BRANCO = "#FFFFFF"

CUSTOM_CSS = f"""
<style>
.stApp {{ background-color: {BRANCO}; }}
section[data-testid="stSidebar"] {{ background-color: {AZUL}; }}
section[data-testid="stSidebar"] * {{ color: {BRANCO} !important; }}
div[data-testid="stMetric"] {{
    background-color: {CINZA_CLARO};
    border: 1px solid {AZUL_CLARO};
    border-radius: 10px;
    padding: 12px 16px;
}}
div[data-testid="stMetric"] label {{ color: {CINZA}; }}
.stButton>button {{
    background-color: {AZUL};
    color: {BRANCO};
    border-radius: 8px;
    border: 1px solid {AZUL};
}}
.stButton>button:hover {{
    background-color: {BRANCO};
    color: {AZUL};
    border: 1px solid {AZUL};
}}
.stForm {{
    border: 1px solid {VERDE};
    border-radius: 10px;
    padding: 16px;
}}
h1, h2, h3 {{ color: {AZUL}; }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

PLOTLY_COLORWAY = [AZUL, VERDE, AZUL_CLARO, CINZA, "#8FBF7F", "#A9B6C8"]
px.defaults.color_discrete_sequence = PLOTLY_COLORWAY

# ───────────────────────────── Login ──────────────────────────────────────
LOGIN_USUARIO = "TELERH"
LOGIN_SENHA = "RH@TELE"


def tela_login() -> bool:
    """Gate simples de usuário/senha. Retorna True se já autenticado."""
    if st.session_state.get("autenticado"):
        return True

    st.markdown(
        f"<h2 style='color:{AZUL};text-align:center;'>🏛️ Painel de Teletrabalho — TJ/MA</h2>",
        unsafe_allow_html=True,
    )
    col_a, col_b, col_c = st.columns([1, 1.2, 1])
    with col_b:
        with st.form("form_login"):
            st.subheader("Acesso restrito")
            usuario = st.text_input("Usuário")
            senha = st.text_input("Senha", type="password")
            entrar = st.form_submit_button("Entrar", use_container_width=True)
        if entrar:
            if usuario.strip().upper() == LOGIN_USUARIO and senha == LOGIN_SENHA:
                st.session_state["autenticado"] = True
                st.session_state["usuario_logado"] = usuario.strip().upper()
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos.")
    return False


if not tela_login():
    st.stop()


# ───────────────────────────── Utilidades ────────────────────────────────
def kpi_row(items: list[tuple[str, object]]) -> None:
    cols = st.columns(len(items))
    for c, (label, value) in zip(cols, items):
        c.metric(label, value)


def df_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    return buf.getvalue()


def filter_multiselect(df: pd.DataFrame, col: str, label: str, key: str) -> pd.DataFrame:
    if col not in df.columns:
        return df
    opts = sorted([v for v in df[col].dropna().unique().tolist() if str(v).strip()])
    chosen = st.multiselect(label, opts, key=key)
    if chosen:
        return df[df[col].isin(chosen)]
    return df


@st.cache_data(ttl=120, show_spinner=False)
def load_produtividade_historico() -> pd.DataFrame:
    ws = sio.get_spreadsheet().worksheet(sio.SHEET_PRODUTIVIDADE)
    values = ws.get_all_values()
    return parse_produtividade(values)


@st.cache_data(ttl=120, show_spinner=False)
def load_produtividade_lancamentos() -> pd.DataFrame:
    try:
        return sio.load_sheet_df(sio.SHEET_PRODUTIVIDADE_LANCAMENTOS, header_row=1)
    except Exception:
        return pd.DataFrame(columns=["id"] + sio.HEADERS[sio.SHEET_PRODUTIVIDADE_LANCAMENTOS])


# ───────────────────────────── Sidebar / navegação ───────────────────────
st.sidebar.title("🏛️ Teletrabalho TJ/MA")
st.sidebar.caption("COGEX-MA/TJMA")

PAGINAS = [
    "📊 Visão Geral",
    "📈 Produtividade",
    "⚖️ Magistrados",
    "🧑‍💼 Servidores Ativos",
    "📄 CNJ - Informações",
    "🚪 Servidores Desligados",
    "📝 Lançar Produtividade Mensal",
    "🖨️ Relatórios",
]
pagina = st.sidebar.radio("Navegação", PAGINAS, label_visibility="collapsed")

st.sidebar.divider()
if st.sidebar.button("🔄 Recarregar dados da planilha"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.caption(
    "Fonte: planilha Google Sheets **Teletrabalho TJ/MA** — 100% dos registros "
    "carregados, sem amostragem. Lacuna na fonte aparece em branco, nunca "
    "preenchida por inferência."
)

# ───────────────────────────── 📊 Visão Geral ────────────────────────────
if pagina == "📊 Visão Geral":
    st.title("📊 Visão Geral — Teletrabalho TJ/MA")

    df_mag = sio.load_sheet_df(sio.SHEET_MAGISTRADOS, sio.HEADER_ROW[sio.SHEET_MAGISTRADOS])
    df_sa = sio.load_sheet_df(sio.SHEET_SERVIDORES_ATIVOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_ATIVOS])
    df_cnj = sio.load_sheet_df(sio.SHEET_CNJ, sio.HEADER_ROW[sio.SHEET_CNJ])
    df_desl = sio.load_sheet_df(sio.SHEET_SERVIDORES_DESLIGADOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_DESLIGADOS])

    kpi_row(
        [
            ("Magistrados em condição especial", len(df_mag)),
            ("Servidores ativos", len(df_sa)),
            ("Pedidos CNJ registrados", len(df_cnj)),
            ("Servidores desligados do regime", len(df_desl)),
        ]
    )

    c1, c2 = st.columns(2)
    with c1:
        if not df_sa.empty and "TIPO" in df_sa.columns:
            fig = px.pie(df_sa, names="TIPO", title="Servidores ativos — por tipo de pedido", hole=0.45)
            st.plotly_chart(fig, use_container_width=True)
    with c2:
        if not df_sa.empty and "LOTAÇÃO" in df_sa.columns:
            top_lot = df_sa["LOTAÇÃO"].value_counts().head(15).reset_index()
            top_lot.columns = ["LOTAÇÃO", "QTD"]
            fig = px.bar(top_lot, x="QTD", y="LOTAÇÃO", orientation="h", title="Top 15 lotações — servidores ativos")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

    if not df_sa.empty and "GÊNERO" in df_sa.columns and "REGIME" in df_sa.columns:
        c3, c4 = st.columns(2)
        with c3:
            fig = px.pie(df_sa, names="GÊNERO", title="Servidores ativos — por gênero", hole=0.45)
            st.plotly_chart(fig, use_container_width=True)
        with c4:
            fig = px.pie(df_sa, names="REGIME", title="Servidores ativos — por regime", hole=0.45)
            st.plotly_chart(fig, use_container_width=True)

# ───────────────────────────── 📈 Produtividade ──────────────────────────
elif pagina == "📈 Produtividade":
    st.title("📈 Produtividade — série histórica")
    st.caption(
        "Parser aplicado sobre a aba original **produtividade** (blocos por "
        "servidor/período) + lançamentos novos feitos por este app. Mês sem "
        "meta/produção lançada aparece em branco, não como zero."
    )

    hist = load_produtividade_historico()
    lanc = load_produtividade_lancamentos()

    tab_v, tab_d = st.tabs(["🔎 Visualização", "🗂️ Dados brutos"])

    with tab_v:
        if hist.empty:
            st.warning("Nenhum registro parseado a partir da aba 'produtividade'.")
        else:
            f1, f2, f3 = st.columns(3)
            with f1:
                servidores = sorted(hist["servidor"].dropna().unique().tolist())
                sel_serv = st.multiselect("Servidor", servidores, key="prod_serv")
            with f2:
                meses = sorted(hist["mes_ano"].dropna().unique().tolist())
                sel_mes = st.multiselect("Mês/Ano", meses, key="prod_mes")
            with f3:
                min_pct = st.slider("Superávit mínimo (produção - meta)", -5000, 5000, -5000, key="prod_sup")

            dff = hist.copy()
            if sel_serv:
                dff = dff[dff["servidor"].isin(sel_serv)]
            if sel_mes:
                dff = dff[dff["mes_ano"].isin(sel_mes)]
            dff = dff[(dff["superavit"].isna()) | (dff["superavit"] >= min_pct)]

            kpi_row(
                [
                    ("Lançamentos (histórico)", int(dff["meta"].notna().sum())),
                    ("Meta total", int(dff["meta"].sum(skipna=True)) if dff["meta"].notna().any() else 0),
                    ("Produção total", int(dff["producao"].sum(skipna=True)) if dff["producao"].notna().any() else 0),
                    ("Servidores no recorte", dff["servidor"].nunique()),
                ]
            )

            plot_df = dff.dropna(subset=["meta", "producao"])
            if not plot_df.empty:
                agg = plot_df.groupby("mes_ano", as_index=False)[["meta", "producao"]].sum()
                fig = px.bar(agg, x="mes_ano", y=["meta", "producao"], barmode="group", title="Meta x Produção por mês")
                st.plotly_chart(fig, use_container_width=True)

            st.dataframe(dff, use_container_width=True, hide_index=True)

        if not lanc.empty:
            st.subheader("Lançamentos manuais (aba produtividade_lancamentos)")
            st.dataframe(lanc, use_container_width=True, hide_index=True)

    with tab_d:
        st.dataframe(hist, use_container_width=True, hide_index=True)

# ───────────────────────────── Abas tabulares padrão ─────────────────────
elif pagina in (
    "⚖️ Magistrados",
    "🧑‍💼 Servidores Ativos",
    "📄 CNJ - Informações",
    "🚪 Servidores Desligados",
):
    MAP = {
        "⚖️ Magistrados": sio.SHEET_MAGISTRADOS,
        "🧑‍💼 Servidores Ativos": sio.SHEET_SERVIDORES_ATIVOS,
        "📄 CNJ - Informações": sio.SHEET_CNJ,
        "🚪 Servidores Desligados": sio.SHEET_SERVIDORES_DESLIGADOS,
    }
    sheet_name = MAP[pagina]
    headers = sio.HEADERS[sheet_name]
    header_row = sio.HEADER_ROW[sheet_name]

    st.title(f"{pagina}")
    df = sio.load_sheet_df(sheet_name, header_row)

    tab_v, tab_e = st.tabs(["🔎 Visualização", "➕ Inserir dado"])

    with tab_v:
        cols_f = st.columns(3)
        dff = df.copy()
        filtraveis = [c for c in headers if c.strip() in (
            "LOTAÇÃO", "LOTAÇÃO - EXERCÍCIO", "LOTAÇÃO DE EXERCÍCIO", "TIPO",
            "REGIME", "GÊNERO", "NATUREZA", "CARGO", "CARGO EFETIVO",
        )]
        for i, col in enumerate(filtraveis[:3]):
            with cols_f[i]:
                dff = filter_multiselect(dff, col, col, key=f"{sheet_name}_{col}")

        busca = st.text_input("Buscar (qualquer coluna)", key=f"{sheet_name}_busca")
        if busca:
            mask = dff.apply(lambda r: r.astype(str).str.contains(busca, case=False, na=False).any(), axis=1)
            dff = dff[mask]

        kpi_row([("Total de registros (filtro)", len(dff)), ("Total de registros (aba)", len(df))])

        nome_col = "NOME" if "NOME" in headers else headers[0]
        if not dff.empty and nome_col in dff.columns:
            for cand in ("LOTAÇÃO", "LOTAÇÃO - EXERCÍCIO", "LOTAÇÃO DE EXERCÍCIO"):
                if cand in dff.columns:
                    top = dff[cand].value_counts().head(15).reset_index()
                    top.columns = [cand, "QTD"]
                    fig = px.bar(top, x="QTD", y=cand, orientation="h", title=f"Top 15 — {cand}")
                    fig.update_layout(yaxis={"categoryorder": "total ascending"})
                    st.plotly_chart(fig, use_container_width=True)
                    break

        st.dataframe(dff, use_container_width=True, hide_index=True)

    with tab_e:
        st.caption(
            "Formulário reflete exatamente as colunas da aba na planilha. "
            "O registro é **acrescentado** ao final — nada existente é alterado."
        )
        with st.form(f"form_{sheet_name}", clear_on_submit=True):
            valores = {}
            for h in headers:
                label = h.strip() if h.strip() else "NOME"
                if label.upper() in ("INÍCIO", "TÉRMINO", "FINAL"):
                    v = st.date_input(label, value=None, key=f"{sheet_name}_{h}")
                    valores[h] = v.strftime("%d/%m/%Y") if v else ""
                elif label.upper() in ("MOTIVOS DO PEDIDO", "FUNDAMENTOS DECISÃO", "FUNDAMENTOS PARA O DEFERIMENTO"):
                    valores[h] = st.text_area(label, key=f"{sheet_name}_{h}")
                else:
                    valores[h] = st.text_input(label, key=f"{sheet_name}_{h}")
            enviado = st.form_submit_button("Salvar na planilha")

        if enviado:
            preenchido = any(str(v).strip() for v in valores.values())
            if not preenchido:
                st.error("Preencha ao menos um campo antes de salvar.")
            else:
                try:
                    novo_id = sio.append_row(sheet_name, ["id"] + headers, valores)
                    st.success(f"Registro salvo com id {novo_id} na aba '{sheet_name}'.")
                    st.cache_data.clear()
                except Exception as exc:
                    st.error(f"Falha ao gravar na planilha: {exc}")

# ───────────────────────────── 📝 Lançar Produtividade Mensal ────────────
elif pagina == "📝 Lançar Produtividade Mensal":
    st.title("📝 Lançar Produtividade Mensal")
    st.caption(
        "A aba original 'produtividade' tem formatação manual em blocos e "
        "**não é editada diretamente** por este app, para não corromper o "
        "histórico. Os lançamentos novos vão para a aba aditiva "
        "'produtividade_lancamentos' e entram automaticamente no painel de "
        "Produtividade."
    )

    hist = load_produtividade_historico()
    servidores_conhecidos = sorted(hist["servidor"].dropna().unique().tolist()) if not hist.empty else []

    with st.form("form_lancamento_mensal", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            servidor = st.selectbox("Servidor (existente)", [""] + servidores_conhecidos)
            servidor_novo = st.text_input("Ou digite um servidor novo")
            matricula = st.text_input("Matrícula")
            lotacao = st.text_input("Lotação")
        with c2:
            mes_ano = st.text_input("Mês/Ano (ex.: jul26)")
            meta = st.number_input("Meta", min_value=0, step=1)
            producao = st.number_input("Produção", min_value=0, step=1)
            processo = st.text_input("Processo (se houver)")
        observacao = st.text_area("Observação")
        enviado = st.form_submit_button("Salvar lançamento")

    if enviado:
        nome_final = (servidor_novo or servidor).strip()
        if not nome_final or not mes_ano.strip():
            st.error("Informe ao menos o servidor e o mês/ano.")
        else:
            row = {
                "SERVIDOR": nome_final,
                "MATRÍCULA": matricula,
                "LOTAÇÃO": lotacao,
                "MÊS/ANO": mes_ano.strip(),
                "META": meta,
                "PRODUÇÃO": producao,
                "PROCESSO": processo,
                "OBSERVAÇÃO": observacao,
                "DATA_LANÇAMENTO": sio.timestamp(),
                "USUÁRIO_LANÇAMENTO": st.session_state.get("usuario_logado", LOGIN_USUARIO),
            }
            try:
                sio.get_or_create_worksheet(
                    sio.SHEET_PRODUTIVIDADE_LANCAMENTOS,
                    sio.HEADERS[sio.SHEET_PRODUTIVIDADE_LANCAMENTOS],
                )
                novo_id = sio.append_row(
                    sio.SHEET_PRODUTIVIDADE_LANCAMENTOS,
                    ["id"] + sio.HEADERS[sio.SHEET_PRODUTIVIDADE_LANCAMENTOS],
                    row,
                )
                st.success(f"Lançamento salvo com id {novo_id}.")
                st.cache_data.clear()
            except Exception as exc:
                st.error(f"Falha ao gravar lançamento: {exc}")

# ───────────────────────────── 🖨️ Relatórios ─────────────────────────────
elif pagina == "🖨️ Relatórios":
    st.title("🖨️ Emissão de Relatórios")

    fonte = st.selectbox(
        "Base para o relatório",
        [
            "Produtividade (histórico parseado)",
            "Magistrados",
            "Servidores Ativos",
            "CNJ - Informações",
            "Servidores Desligados",
        ],
    )

    if fonte == "Produtividade (histórico parseado)":
        df = load_produtividade_historico()
    elif fonte == "Magistrados":
        df = sio.load_sheet_df(sio.SHEET_MAGISTRADOS, sio.HEADER_ROW[sio.SHEET_MAGISTRADOS])
    elif fonte == "Servidores Ativos":
        df = sio.load_sheet_df(sio.SHEET_SERVIDORES_ATIVOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_ATIVOS])
    elif fonte == "CNJ - Informações":
        df = sio.load_sheet_df(sio.SHEET_CNJ, sio.HEADER_ROW[sio.SHEET_CNJ])
    else:
        df = sio.load_sheet_df(sio.SHEET_SERVIDORES_DESLIGADOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_DESLIGADOS])

    st.dataframe(df, use_container_width=True, hide_index=True)
    st.caption(f"{len(df)} registro(s) no relatório — 100% da base filtrada, sem amostragem.")

    col1, col2 = st.columns(2)
    with col1:
        xlsx_bytes = df_to_excel_bytes({fonte[:31]: df})
        st.download_button(
            "⬇️ Baixar Excel (.xlsx)",
            data=xlsx_bytes,
            file_name=f"relatorio_{fonte.split()[0].lower()}_{dt.date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col2:
        csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Baixar CSV",
            data=csv_bytes,
            file_name=f"relatorio_{fonte.split()[0].lower()}_{dt.date.today().isoformat()}.csv",
            mime="text/csv",
        )

# ───────────────────────────── Rodapé ─────────────────────────────────────
st.markdown(
    "<div style='text-align:right; font-size:9px; color:#B9C0CA; "
    "margin-top:40px;'>ADMJESUSIA 107805</div>",
    unsafe_allow_html=True,
)
