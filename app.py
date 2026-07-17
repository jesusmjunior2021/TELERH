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
import relatorios
from parser_produtividade import parse_produtividade

# ───────────────────────────── Config geral ──────────────────────────────
st.set_page_config(
    page_title="Painel de Teletrabalho — TJ/MA",
    page_icon="🏛️",
    layout="wide",
)

# Paleta — tema escuro com azul institucional.
AZUL_FUNDO = "#0B1220"
AZUL_PAINEL = "#131C2E"
AZUL_PAINEL_2 = "#1A2740"
AZUL_PRIMARIO = "#2F6FED"
AZUL_CLARO = "#5B8DEF"
AZUL_ESCURO = "#1B3A6B"
BORDA = "#24314A"
TEXTO_CLARO = "#E7ECF5"
TEXTO_SECUNDARIO = "#9AA7BD"
VERDE_OK = "#22C55E"
AMARELO_ALERTA = "#F5B942"
VERMELHO_ALERTA = "#EF4444"

CUSTOM_CSS = f"""
<style>
.stApp {{ background-color: {AZUL_FUNDO}; }}

section[data-testid="stSidebar"] {{
    background-color: {AZUL_PAINEL};
    border-right: 1px solid {BORDA};
}}
section[data-testid="stSidebar"] * {{ color: {TEXTO_CLARO} !important; }}

div[data-testid="stMetric"] {{
    background: linear-gradient(160deg, {AZUL_PAINEL_2}, {AZUL_PAINEL});
    border: 1px solid {BORDA};
    border-radius: 12px;
    padding: 14px 18px;
    box-shadow: 0 0 0 1px rgba(47,111,237,0.06), 0 4px 14px rgba(0,0,0,0.25);
}}
div[data-testid="stMetric"] label {{ color: {TEXTO_SECUNDARIO} !important; }}
div[data-testid="stMetricValue"] {{ color: {TEXTO_CLARO} !important; }}

.stButton>button, .stDownloadButton>button {{
    background-color: {AZUL_PRIMARIO};
    color: #FFFFFF;
    border-radius: 8px;
    border: 1px solid {AZUL_PRIMARIO};
    font-weight: 600;
}}
.stButton>button:hover, .stDownloadButton>button:hover {{
    background-color: {AZUL_CLARO};
    border: 1px solid {AZUL_CLARO};
    color: #FFFFFF;
}}

.stForm {{
    border: 1px solid {BORDA};
    background-color: {AZUL_PAINEL};
    border-radius: 12px;
    padding: 18px;
}}

h1, h2, h3 {{ color: {TEXTO_CLARO}; }}
h1 {{ border-bottom: 2px solid {AZUL_PRIMARIO}; padding-bottom: 8px; }}
p, span, label, .stMarkdown {{ color: {TEXTO_CLARO}; }}
.stCaption, [data-testid="stCaptionContainer"] {{ color: {TEXTO_SECUNDARIO} !important; }}

div[data-testid="stDataFrame"] {{
    border: 1px solid {BORDA};
    border-radius: 10px;
    overflow: hidden;
}}

.stTabs [data-baseweb="tab-list"] {{ gap: 4px; }}
.stTabs [data-baseweb="tab"] {{
    background-color: {AZUL_PAINEL};
    border-radius: 8px 8px 0 0;
    color: {TEXTO_SECUNDARIO};
}}
.stTabs [aria-selected="true"] {{
    background-color: {AZUL_PAINEL_2} !important;
    color: {TEXTO_CLARO} !important;
    border-bottom: 2px solid {AZUL_PRIMARIO};
}}

/* Feedback visual de salvamento */
.salvo-ok {{
    background: linear-gradient(90deg, {AZUL_ESCURO}, {VERDE_OK});
    color: #FFFFFF;
    padding: 12px 18px;
    border-radius: 10px;
    font-weight: 600;
    margin: 10px 0 6px 0;
    animation: slideFade 0.35s ease-out;
}}
@keyframes slideFade {{
    from {{ opacity: 0; transform: translateY(-8px); }}
    to {{ opacity: 1; transform: translateY(0); }}
}}

.rodape-app {{
    text-align: right;
    font-size: 9px;
    color: #3A4459;
    margin-top: 40px;
}}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

PLOTLY_COLORWAY = [AZUL_PRIMARIO, VERDE_OK, AZUL_CLARO, AMARELO_ALERTA, "#8FBF7F", TEXTO_SECUNDARIO]
px.defaults.color_discrete_sequence = PLOTLY_COLORWAY
px.defaults.template = "plotly_dark"

PLOTLY_LAYOUT_ESCURO = dict(
    paper_bgcolor=AZUL_FUNDO,
    plot_bgcolor=AZUL_FUNDO,
    font_color=TEXTO_CLARO,
    title_font_color=TEXTO_CLARO,
    legend_font_color=TEXTO_CLARO,
)


def estilizar(fig):
    fig.update_layout(**PLOTLY_LAYOUT_ESCURO)
    return fig


# ───────────────────────────── Login ──────────────────────────────────────
LOGIN_USUARIO = "TELERH"
LOGIN_SENHA = "RH@TELE"


def tela_login() -> bool:
    """Gate simples de usuário/senha. Retorna True se já autenticado."""
    if st.session_state.get("autenticado"):
        return True

    st.markdown(
        f"<h2 style='color:{TEXTO_CLARO};text-align:center;'>🏛️ Painel de Teletrabalho — TJ/MA</h2>",
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


def feedback_salvo(msg: str) -> None:
    """Feedback visual de salvamento — toast + barra animada, sem depender
    de o usuário notar um st.success discreto."""
    st.toast(msg, icon="✅")
    st.markdown(f"<div class='salvo-ok'>✅ {msg}</div>", unsafe_allow_html=True)


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


# Colunas que, quando presentes num formulário de inserção, viram caixa de
# seleção (com opção de digitar valor novo) em vez de campo de texto livre —
# acelera o lançamento e padroniza a grafia usada na planilha.
CAMPOS_SELECAO = {
    "TIPO", "REGIME", "GÊNERO", "GRAU", "NATUREZA", "CARGO", "CARGO EFETIVO",
    "FUNÇÃO", "ESTÁGIO PROBATÓRIO", "COND ESPEC", "CONDIÇÕES ESPECIAIS",
    "LOTAÇÃO", "LOTAÇÃO - EXERCÍCIO", "LOTAÇÃO ORIGEM", "LOTAÇÃO DE EXERCÍCIO",
    "LOTAÇÃO DE ORIGEM", "MOTIVO",
}
OPCAO_NOVO_VALOR = "✏️ Digitar novo valor"
OPCAO_VAZIA = "— selecionar —"


def campo_formulario(df_ref: pd.DataFrame, sheet_name: str, h: str):
    """Escolhe o widget certo para o campo `h`: data, área de texto longa,
    seleção rápida (com opção de digitar valor novo) ou texto simples."""
    label = h.strip() if h.strip() else "NOME"
    chave = label.upper()

    if chave in ("INÍCIO", "TÉRMINO", "FINAL"):
        v = st.date_input(label, value=None, key=f"{sheet_name}_{h}")
        return v.strftime("%d/%m/%Y") if v else ""

    if chave in ("MOTIVOS DO PEDIDO", "FUNDAMENTOS DECISÃO", "FUNDAMENTOS PARA O DEFERIMENTO"):
        return st.text_area(label, key=f"{sheet_name}_{h}")

    if chave in CAMPOS_SELECAO and h in df_ref.columns:
        opcoes = sorted([v for v in df_ref[h].dropna().unique().tolist() if str(v).strip()])
        escolha = st.selectbox(
            label, [OPCAO_VAZIA] + opcoes + [OPCAO_NOVO_VALOR], key=f"{sheet_name}_{h}_sel"
        )
        if escolha == OPCAO_NOVO_VALOR:
            return st.text_input(f"Novo valor — {label}", key=f"{sheet_name}_{h}_novo")
        if escolha == OPCAO_VAZIA:
            return ""
        return escolha

    return st.text_input(label, key=f"{sheet_name}_{h}")


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


@st.cache_data(ttl=120, show_spinner=False)
def load_ponto() -> pd.DataFrame:
    try:
        return sio.load_sheet_df(sio.SHEET_PONTO_TELETRABALHO, header_row=1)
    except Exception:
        return pd.DataFrame(columns=["id"] + sio.HEADERS[sio.SHEET_PONTO_TELETRABALHO])


# ───────────────────────────── Sidebar / navegação ───────────────────────
st.sidebar.title("🏛️ Teletrabalho TJ/MA")
st.sidebar.caption("COGEX-MA/TJMA")

PAGINAS = [
    "📊 Dashboard",
    "🕒 Check-in / Check-out",
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
    "preenchida por inferência. Toda inserção é feita por aqui — o uso "
    "direto da planilha deve ser evitado, para preservar formatação e "
    "reduzir erro humano."
)

# ───────────────────────────── 📊 Dashboard ───────────────────────────────
if pagina == "📊 Dashboard":
    st.title("📊 Dashboard — Teletrabalho TJ/MA")

    df_mag = sio.load_sheet_df(sio.SHEET_MAGISTRADOS, sio.HEADER_ROW[sio.SHEET_MAGISTRADOS])
    df_sa = sio.load_sheet_df(sio.SHEET_SERVIDORES_ATIVOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_ATIVOS])
    df_cnj = sio.load_sheet_df(sio.SHEET_CNJ, sio.HEADER_ROW[sio.SHEET_CNJ])
    df_desl = sio.load_sheet_df(sio.SHEET_SERVIDORES_DESLIGADOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_DESLIGADOS])
    df_ponto = load_ponto()

    hoje_str = dt.date.today().strftime("%d/%m/%Y")
    checkins_hoje = 0
    checkouts_hoje = 0
    if not df_ponto.empty and "DATA" in df_ponto.columns and "TIPO" in df_ponto.columns:
        hoje_df = df_ponto[df_ponto["DATA"] == hoje_str]
        checkins_hoje = int((hoje_df["TIPO"] == "CHECK-IN").sum())
        checkouts_hoje = int((hoje_df["TIPO"] == "CHECK-OUT").sum())

    kpi_row(
        [
            ("Servidores ativos", len(df_sa)),
            ("Magistrados em condição especial", len(df_mag)),
            ("Check-ins hoje", checkins_hoje),
            ("Check-outs hoje", checkouts_hoje),
        ]
    )
    kpi_row(
        [
            ("Pedidos CNJ registrados", len(df_cnj)),
            ("Servidores desligados do regime", len(df_desl)),
            ("Registros de ponto (total)", len(df_ponto)),
            ("Última atualização", dt.datetime.now().strftime("%H:%M")),
        ]
    )

    c1, c2 = st.columns(2)
    with c1:
        if not df_sa.empty and "TIPO" in df_sa.columns:
            fig = px.pie(df_sa, names="TIPO", title="Servidores ativos — por tipo de pedido", hole=0.45)
            st.plotly_chart(estilizar(fig), use_container_width=True)
    with c2:
        if not df_sa.empty and "LOTAÇÃO" in df_sa.columns:
            top_lot = df_sa["LOTAÇÃO"].value_counts().head(15).reset_index()
            top_lot.columns = ["LOTAÇÃO", "QTD"]
            fig = px.bar(top_lot, x="QTD", y="LOTAÇÃO", orientation="h", title="Top 15 lotações — servidores ativos")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(estilizar(fig), use_container_width=True)

    if not df_sa.empty and "GÊNERO" in df_sa.columns and "REGIME" in df_sa.columns:
        c3, c4 = st.columns(2)
        with c3:
            fig = px.pie(df_sa, names="GÊNERO", title="Servidores ativos — por gênero", hole=0.45)
            st.plotly_chart(estilizar(fig), use_container_width=True)
        with c4:
            fig = px.pie(df_sa, names="REGIME", title="Servidores ativos — por regime", hole=0.45)
            st.plotly_chart(estilizar(fig), use_container_width=True)

    if not df_ponto.empty and "DATA" in df_ponto.columns:
        st.subheader("Movimentação de ponto — últimos registros")
        por_dia = df_ponto.groupby(["DATA", "TIPO"]).size().reset_index(name="QTD")
        if not por_dia.empty:
            fig = px.bar(por_dia, x="DATA", y="QTD", color="TIPO", barmode="group", title="Check-in x Check-out por dia")
            st.plotly_chart(estilizar(fig), use_container_width=True)

# ───────────────────────────── 🕒 Check-in / Check-out ───────────────────
elif pagina == "🕒 Check-in / Check-out":
    st.title("🕒 Registro de Ponto — Check-in / Check-out")
    st.caption(
        "Registro rápido de entrada e saída do teletrabalho do dia. Grava "
        "direto na aba 'ponto_teletrabalho' da planilha — não é preciso "
        "abrir o Google Sheets."
    )

    df_sa = sio.load_sheet_df(sio.SHEET_SERVIDORES_ATIVOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_ATIVOS])
    nome_col_sa = "NOME" if "NOME" in df_sa.columns else (df_sa.columns[0] if not df_sa.empty else None)
    matricula_col_sa = "MATRÍCULA" if "MATRÍCULA" in df_sa.columns else None
    servidores_ativos = sorted(df_sa[nome_col_sa].dropna().unique().tolist()) if nome_col_sa else []

    with st.form("form_ponto", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            servidor_sel = st.selectbox(
                "Servidor", [OPCAO_VAZIA] + servidores_ativos + [OPCAO_NOVO_VALOR], key="ponto_servidor"
            )
            servidor_novo = (
                st.text_input("Nome (se não estiver na lista)", key="ponto_servidor_novo")
                if servidor_sel == OPCAO_NOVO_VALOR
                else ""
            )
            tipo_registro = st.radio(
                "Tipo de registro", ["Check-in (entrada)", "Check-out (saída)"], horizontal=True, key="ponto_tipo"
            )
        with c2:
            data_registro = st.date_input("Data", value=dt.date.today(), key="ponto_data")
            hora_registro = st.time_input("Horário", value=dt.datetime.now().time(), key="ponto_hora")
            observacao_ponto = st.text_input("Observação (opcional)", key="ponto_obs")
        enviado_ponto = st.form_submit_button("Registrar", use_container_width=True)

    if enviado_ponto:
        nome_final = servidor_novo.strip() if servidor_sel == OPCAO_NOVO_VALOR else (
            servidor_sel if servidor_sel != OPCAO_VAZIA else ""
        )
        if not nome_final:
            st.error("Selecione ou digite o nome do servidor.")
        else:
            matricula = ""
            if matricula_col_sa and not df_sa.empty:
                achado = df_sa[df_sa[nome_col_sa] == nome_final]
                if not achado.empty:
                    matricula = str(achado.iloc[0][matricula_col_sa])

            tipo_final = "CHECK-IN" if tipo_registro.startswith("Check-in") else "CHECK-OUT"
            row = {
                "SERVIDOR": nome_final,
                "MATRÍCULA": matricula,
                "DATA": data_registro.strftime("%d/%m/%Y"),
                "HORARIO": hora_registro.strftime("%H:%M"),
                "TIPO": tipo_final,
                "OBSERVAÇÃO": observacao_ponto,
                "DATA_LANÇAMENTO": sio.timestamp(),
                "USUÁRIO_LANÇAMENTO": st.session_state.get("usuario_logado", LOGIN_USUARIO),
            }
            try:
                sio.get_or_create_worksheet(sio.SHEET_PONTO_TELETRABALHO, sio.HEADERS[sio.SHEET_PONTO_TELETRABALHO])
                novo_id = sio.append_row(
                    sio.SHEET_PONTO_TELETRABALHO, ["id"] + sio.HEADERS[sio.SHEET_PONTO_TELETRABALHO], row
                )
                feedback_salvo(f"{tipo_final} registrado para {nome_final} às {row['HORARIO']} (id {novo_id}).")
                st.cache_data.clear()
            except Exception as exc:
                st.error(f"Falha ao registrar ponto: {exc}")

    st.divider()
    st.subheader("Registros recentes")
    pontos = load_ponto()
    if not pontos.empty:
        ordenar_por = "DATA_LANÇAMENTO" if "DATA_LANÇAMENTO" in pontos.columns else pontos.columns[-1]
        st.dataframe(
            pontos.sort_values(ordenar_por, ascending=False), use_container_width=True, hide_index=True
        )
    else:
        st.info("Nenhum registro de ponto ainda.")

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
            dff = dff[dff["superavit"].fillna(-999999) >= min_pct]

            if not dff.empty:
                fig = px.bar(
                    dff.sort_values("mes_ano"),
                    x="mes_ano",
                    y=["meta", "producao"],
                    barmode="group",
                    title="Meta x Produção por mês (filtro aplicado)",
                    facet_col="servidor" if dff["servidor"].nunique() <= 4 else None,
                )
                st.plotly_chart(estilizar(fig), use_container_width=True)

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
                    st.plotly_chart(estilizar(fig), use_container_width=True)
                    break

        st.dataframe(dff, use_container_width=True, hide_index=True)

    with tab_e:
        st.caption(
            "Formulário reflete exatamente as colunas da aba na planilha. "
            "Campos com valores conhecidos viram caixa de seleção — escolha "
            "um existente ou digite um novo. O registro é **acrescentado** "
            "ao final — nada existente é alterado."
        )
        with st.form(f"form_{sheet_name}", clear_on_submit=True):
            valores = {}
            for h in headers:
                valores[h] = campo_formulario(df, sheet_name, h)
            enviado = st.form_submit_button("Salvar na planilha", use_container_width=True)

        if enviado:
            preenchido = any(str(v).strip() for v in valores.values())
            if not preenchido:
                st.error("Preencha ao menos um campo antes de salvar.")
            else:
                try:
                    novo_id = sio.append_row(sheet_name, ["id"] + headers, valores)
                    feedback_salvo(f"Registro salvo com id {novo_id} na aba '{sheet_name}'.")
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
        enviado = st.form_submit_button("Salvar lançamento", use_container_width=True)

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
                feedback_salvo(f"Lançamento salvo com id {novo_id}.")
                st.cache_data.clear()
            except Exception as exc:
                st.error(f"Falha ao gravar lançamento: {exc}")

# ───────────────────────────── 🖨️ Relatórios ─────────────────────────────
elif pagina == "🖨️ Relatórios":
    st.title("🖨️ Emissão de Relatórios")
    st.caption(
        "Relatórios com timbre da Diretoria de RH do TJMA, filtráveis por "
        "aba de origem, prontos para impressão (HTML) ou distribuição (PDF)."
    )

    fonte = st.selectbox(
        "Base para o relatório",
        [
            "Produtividade (histórico parseado)",
            "Lançamentos de Produtividade",
            "Registro de Ponto (Check-in-Check-out)",
            "Magistrados",
            "Servidores Ativos",
            "CNJ - Informações",
            "Servidores Desligados",
        ],
    )

    if fonte == "Produtividade (histórico parseado)":
        df = load_produtividade_historico()
    elif fonte == "Lançamentos de Produtividade":
        df = load_produtividade_lancamentos()
    elif fonte == "Registro de Ponto (Check-in-Check-out)":
        df = load_ponto()
    elif fonte == "Magistrados":
        df = sio.load_sheet_df(sio.SHEET_MAGISTRADOS, sio.HEADER_ROW[sio.SHEET_MAGISTRADOS])
    elif fonte == "Servidores Ativos":
        df = sio.load_sheet_df(sio.SHEET_SERVIDORES_ATIVOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_ATIVOS])
    elif fonte == "CNJ - Informações":
        df = sio.load_sheet_df(sio.SHEET_CNJ, sio.HEADER_ROW[sio.SHEET_CNJ])
    else:
        df = sio.load_sheet_df(sio.SHEET_SERVIDORES_DESLIGADOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_DESLIGADOS])

    dff = df.copy()
    filtros_aplicados: dict[str, str] = {}

    if not df.empty:
        colunas_filtraveis = [
            c for c in df.columns if c.strip().upper() in (
                "LOTAÇÃO", "LOTAÇÃO - EXERCÍCIO", "LOTAÇÃO DE EXERCÍCIO", "TIPO",
                "REGIME", "GÊNERO", "NATUREZA", "CARGO", "CARGO EFETIVO",
                "SERVIDOR", "MÊS/ANO", "MATRÍCULA",
            )
        ]
        if colunas_filtraveis:
            cols_f = st.columns(min(3, len(colunas_filtraveis)))
            for i, col in enumerate(colunas_filtraveis[:3]):
                with cols_f[i]:
                    opts = sorted([v for v in dff[col].dropna().unique().tolist() if str(v).strip()])
                    escolhidos = st.multiselect(col, opts, key=f"rel_{fonte}_{col}")
                    if escolhidos:
                        dff = dff[dff[col].isin(escolhidos)]
                        filtros_aplicados[col] = ", ".join(escolhidos)

        busca_rel = st.text_input("Buscar (qualquer coluna)", key=f"rel_busca_{fonte}")
        if busca_rel:
            mask = dff.apply(lambda r: r.astype(str).str.contains(busca_rel, case=False, na=False).any(), axis=1)
            dff = dff[mask]
            filtros_aplicados["Busca"] = busca_rel

    st.dataframe(dff, use_container_width=True, hide_index=True)
    st.caption(f"{len(dff)} registro(s) no relatório — de {len(df)} no total da base, sem amostragem.")

    col1, col2, col3, col4 = st.columns(4)
    nome_arquivo_base = f"relatorio_{fonte.split()[0].lower()}_{dt.date.today().isoformat()}"
    with col1:
        xlsx_bytes = df_to_excel_bytes({fonte[:31]: dff})
        st.download_button(
            "⬇️ Excel (.xlsx)",
            data=xlsx_bytes,
            file_name=f"{nome_arquivo_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col2:
        csv_bytes = dff.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ CSV",
            data=csv_bytes,
            file_name=f"{nome_arquivo_base}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col3:
        html_str = relatorios.gerar_html(fonte, dff, filtros_aplicados)
        st.download_button(
            "⬇️ HTML (impressão)",
            data=html_str.encode("utf-8"),
            file_name=f"{nome_arquivo_base}.html",
            mime="text/html",
            use_container_width=True,
        )
    with col4:
        pdf_bytes = relatorios.gerar_pdf(fonte, dff, filtros_aplicados)
        st.download_button(
            "⬇️ PDF",
            data=pdf_bytes,
            file_name=f"{nome_arquivo_base}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

# ───────────────────────────── Rodapé ─────────────────────────────────────
st.markdown("<div class='rodape-app'>ADMJESUSIA 107805</div>", unsafe_allow_html=True)
