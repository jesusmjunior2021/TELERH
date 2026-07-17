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
TEAL_INSERCAO = "#0EA5A5"
TEAL_INSERCAO_CLARO = "#2DD4BF"
AMBAR_RELATORIO = "#F5B942"
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
    border-left: 4px solid {TEAL_INSERCAO};
    background-color: {AZUL_PAINEL};
    border-radius: 12px;
    padding: 18px;
}}

/* Selos de contexto — sinalizam se a tela é de inserção, leitura ou relatório */
.badge-insercao, .badge-leitura, .badge-relatorio {{
    display: inline-block;
    padding: 4px 12px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.3px;
    margin-bottom: 10px;
}}
.badge-insercao {{
    background: rgba(14,165,165,0.16);
    color: {TEAL_INSERCAO_CLARO};
    border: 1px solid {TEAL_INSERCAO_CLARO};
}}
.badge-leitura {{
    background: rgba(91,141,239,0.16);
    color: {AZUL_CLARO};
    border: 1px solid {AZUL_CLARO};
}}
.badge-relatorio {{
    background: rgba(245,185,66,0.16);
    color: {AMBAR_RELATORIO};
    border: 1px solid {AMBAR_RELATORIO};
}}

.aviso-app {{
    background: rgba(245,185,66,0.10);
    border: 1px solid {AMBAR_RELATORIO};
    border-radius: 10px;
    padding: 10px 14px;
    font-size: 13px;
    color: {TEXTO_CLARO};
    margin-bottom: 12px;
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


def combinar_produtividade(hist: pd.DataFrame, lanc: pd.DataFrame) -> pd.DataFrame:
    """Junta o histórico parseado da aba 'produtividade' com os lançamentos
    manuais feitos pelo app, num único formato, para que os gráficos e
    filtros reflitam os dois — sem alterar nenhum dos dois na origem."""
    hist2 = hist.copy()
    if not hist2.empty:
        hist2["origem"] = "histórico (planilha)"

    if lanc.empty:
        return hist2

    lanc2 = lanc.rename(
        columns={
            "SERVIDOR": "servidor",
            "MATRÍCULA": "matricula",
            "LOTAÇÃO": "lotacao",
            "MÊS/ANO": "mes_ano",
            "META": "meta",
            "PRODUÇÃO": "producao",
            "PROCESSO": "processo_mes",
            "OBSERVAÇÃO": "observacao",
        }
    ).copy()
    lanc2["meta"] = pd.to_numeric(lanc2.get("meta"), errors="coerce")
    lanc2["producao"] = pd.to_numeric(lanc2.get("producao"), errors="coerce")
    lanc2["superavit"] = lanc2["producao"] - lanc2["meta"]
    lanc2["origem"] = "lançamento manual (app)"

    return pd.concat([hist2, lanc2], ignore_index=True, sort=False)


# ───────────────────────────── Sidebar / navegação ───────────────────────
st.sidebar.title("🏛️ Teletrabalho TJ/MA")
st.sidebar.caption("COGEX-MA/TJMA")

PAGINAS = [
    "📊 Dashboard",
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
    st.markdown("<span class='badge-leitura'>🔎 MODO LEITURA</span>", unsafe_allow_html=True)

    df_mag = sio.load_sheet_df(sio.SHEET_MAGISTRADOS, sio.HEADER_ROW[sio.SHEET_MAGISTRADOS])
    df_sa = sio.load_sheet_df(sio.SHEET_SERVIDORES_ATIVOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_ATIVOS])
    df_cnj = sio.load_sheet_df(sio.SHEET_CNJ, sio.HEADER_ROW[sio.SHEET_CNJ])
    df_desl = sio.load_sheet_df(sio.SHEET_SERVIDORES_DESLIGADOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_DESLIGADOS])

    kpi_row(
        [
            ("Servidores ativos", len(df_sa)),
            ("Magistrados em condição especial", len(df_mag)),
            ("Pedidos CNJ registrados", len(df_cnj)),
            ("Servidores desligados do regime", len(df_desl)),
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
    combinado = combinar_produtividade(hist, lanc)

    st.markdown("<span class='badge-leitura'>🔎 MODO LEITURA</span>", unsafe_allow_html=True)

    tab_v, tab_d = st.tabs(["🔎 Visualização", "🗂️ Dados brutos"])

    with tab_v:
        if combinado.empty:
            st.warning("Nenhum registro parseado a partir da aba 'produtividade' nem lançado pelo app.")
        else:
            f1, f2, f3 = st.columns(3)
            with f1:
                servidores = sorted(combinado["servidor"].dropna().unique().tolist())
                sel_serv = st.multiselect("Servidor", servidores, key="prod_serv")
            with f2:
                meses = sorted(combinado["mes_ano"].dropna().unique().tolist())
                sel_mes = st.multiselect("Mês/Ano", meses, key="prod_mes")
            with f3:
                min_pct = st.slider("Superávit mínimo (produção - meta)", -5000, 5000, -5000, key="prod_sup")

            dff = combinado.copy()
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
                    title="Meta x Produção por mês — histórico + lançamentos do app (filtro aplicado)",
                    facet_col="servidor" if dff["servidor"].nunique() <= 4 else None,
                )
                st.plotly_chart(estilizar(fig), use_container_width=True)

            st.caption(
                "Coluna **origem** mostra se a linha vem do histórico original da "
                "planilha ou de um lançamento feito por este app."
            )
            st.dataframe(dff, use_container_width=True, hide_index=True)

    with tab_d:
        sub_hist, sub_lanc = st.tabs(["Histórico (planilha)", "Lançamentos manuais (app)"])
        with sub_hist:
            st.dataframe(hist, use_container_width=True, hide_index=True)
        with sub_lanc:
            if lanc.empty:
                st.info("Nenhum lançamento manual ainda.")
            else:
                st.dataframe(lanc, use_container_width=True, hide_index=True)

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
        st.markdown("<span class='badge-leitura'>🔎 MODO LEITURA</span>", unsafe_allow_html=True)
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
        st.markdown("<span class='badge-insercao'>✏️ MODO INSERÇÃO — grava na planilha</span>", unsafe_allow_html=True)
        st.markdown(
            "<div class='aviso-app'>⚠️ A partir de agora, servidores, magistrados e demais "
            "registros novos devem ser cadastrados <strong>somente por este formulário</strong>. "
            "Editar a planilha diretamente pode quebrar a formatação e o parser que alimenta "
            "os relatórios.</div>",
            unsafe_allow_html=True,
        )
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
    st.markdown("<span class='badge-insercao'>✏️ MODO INSERÇÃO — grava na planilha</span>", unsafe_allow_html=True)
    st.caption(
        "A aba original 'produtividade' tem formatação manual em blocos e "
        "**não é editada diretamente** por este app, para não corromper o "
        "histórico. Os lançamentos novos vão para a aba aditiva "
        "'produtividade_lancamentos' e entram automaticamente no painel de "
        "Produtividade — nenhum registro existente é apagado ou sobrescrito."
    )

    hist = load_produtividade_historico()
    df_sa = sio.load_sheet_df(sio.SHEET_SERVIDORES_ATIVOS, sio.HEADER_ROW[sio.SHEET_SERVIDORES_ATIVOS])
    nome_col_sa = "NOME" if "NOME" in df_sa.columns else (df_sa.columns[0] if not df_sa.empty else None)

    servidores_hist = set(hist["servidor"].dropna().unique().tolist()) if not hist.empty else set()
    servidores_ativos_nomes = set(df_sa[nome_col_sa].dropna().unique().tolist()) if nome_col_sa else set()
    servidores_conhecidos = sorted(servidores_hist | servidores_ativos_nomes)

    # Selectbox de servidor FORA do form: precisa disparar rerun na hora para
    # já buscar matrícula/lotação reais da aba "servidores ativos" (convergência
    # de dados) antes de o usuário preencher o resto do formulário.
    servidor_sel = st.selectbox(
        "Servidor", [OPCAO_VAZIA] + servidores_conhecidos + [OPCAO_NOVO_VALOR], key="prodm_servidor"
    )
    servidor_novo = (
        st.text_input("Nome do novo servidor", key="prodm_servidor_novo")
        if servidor_sel == OPCAO_NOVO_VALOR
        else ""
    )
    nome_final = servidor_novo.strip() if servidor_sel == OPCAO_NOVO_VALOR else (
        servidor_sel if servidor_sel != OPCAO_VAZIA else ""
    )

    matricula_padrao, lotacao_padrao = "", ""
    if nome_final and nome_col_sa and not df_sa.empty:
        achado = df_sa[df_sa[nome_col_sa] == nome_final]
        if not achado.empty:
            if "MATRÍCULA" in achado.columns:
                matricula_padrao = str(achado.iloc[0]["MATRÍCULA"]).strip()
            if "LOTAÇÃO" in achado.columns:
                lotacao_padrao = str(achado.iloc[0]["LOTAÇÃO"]).strip()
            if matricula_padrao or lotacao_padrao:
                st.caption(
                    "🔗 Matrícula e lotação preenchidas automaticamente a partir "
                    "da aba 'servidores ativos' — confira antes de salvar."
                )

    with st.form("form_lancamento_mensal", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            matricula = st.text_input("Matrícula", value=matricula_padrao, key="prodm_matricula")
            lotacao = st.text_input("Lotação", value=lotacao_padrao, key="prodm_lotacao")
            mes_ano = st.text_input("Mês/Ano (ex.: jul26)", key="prodm_mes")
        with c2:
            meta = st.number_input("Meta", min_value=0, step=1, key="prodm_meta")
            producao = st.number_input("Produção", min_value=0, step=1, key="prodm_producao")
            processo = st.text_input("Processo (se houver)", key="prodm_processo")
        observacao = st.text_area("Observação", key="prodm_obs")
        enviado = st.form_submit_button("Salvar lançamento", use_container_width=True)

    if enviado:
        if not nome_final or not mes_ano.strip():
            st.error("Informe ao menos o servidor e o mês/ano.")
        else:
            lanc_existente = load_produtividade_lancamentos()
            ja_existe = False
            if not lanc_existente.empty and {"SERVIDOR", "MÊS/ANO"}.issubset(lanc_existente.columns):
                ja_existe = not lanc_existente[
                    (lanc_existente["SERVIDOR"] == nome_final)
                    & (lanc_existente["MÊS/ANO"] == mes_ano.strip())
                ].empty
            if ja_existe:
                st.warning(
                    f"Já existe lançamento de **{nome_final}** para **{mes_ano.strip()}**. "
                    "O registro novo foi adicionado como uma linha à parte — nada do "
                    "lançamento anterior foi apagado ou alterado. Revise em "
                    "'📈 Produtividade → Dados brutos' se não era essa a intenção."
                )
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
    st.markdown("<span class='badge-relatorio'>🖨️ EMISSÃO DE RELATÓRIO</span>", unsafe_allow_html=True)
    st.caption(
        "Relatórios com timbre da Diretoria de RH do TJMA, filtráveis por "
        "aba de origem, prontos para impressão (HTML) ou distribuição (PDF)."
    )

    fonte = st.selectbox(
        "Base para o relatório",
        [
            "Produtividade (histórico parseado)",
            "Lançamentos de Produtividade",
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
