# MAT-TELETRABALHO-001 · sheets_io.py
# Camada de acesso ao Google Sheets (planilha "Teletrabalho TJ/MA").
# Regra dura: nunca sobrescrever registro existente. Toda escrita é aditiva
# (append) ou update pontual de uma linha específica pelo próprio ID.

from __future__ import annotations

import datetime as dt
from typing import Optional

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = "1EwAxGlt04CGhOWCMXEaJyzYMXpJPCVlQfXnl9i0X-gw"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Nome real das abas na planilha de origem.
SHEET_PRODUTIVIDADE = "produtividade"
SHEET_MAGISTRADOS = "Magistrados"
SHEET_SERVIDORES_ATIVOS = "servidores ativos"
SHEET_CNJ = "CNJ - Informações"
SHEET_SERVIDORES_DESLIGADOS = "servidores desligados"

# Aba nova, criada por este app, para lançamento mensal de produtividade
# (aditiva — não mexe no bloco irregular histórico da aba "produtividade").
SHEET_PRODUTIVIDADE_LANCAMENTOS = "produtividade_lancamentos"

# Aba nova, criada por este app, para check-in/check-out de teletrabalho.
SHEET_PONTO_TELETRABALHO = "ponto_teletrabalho"

# Cabeçalho REAL de cada aba tabular (na ordem em que aparece na planilha).
# O rótulo "NOME" entre colchetes em servidores ativos/desligados marca uma
# coluna que na planilha não tem cabeçalho de texto (célula vazia na linha 1),
# mas contém o nome do servidor — mantido assim para não inventar rótulo novo
# na aba original.
HEADERS = {
    SHEET_MAGISTRADOS: [
        "NOME", "MATRÍCULA", "PROCESSO", "CARGO EFETIVO", "FUNÇÃO",
        "LOTAÇÃO - EXERCÍCIO", "LOTAÇÃO ORIGEM", "NATUREZA", "INÍCIO",
        "TÉRMINO", "MOTIVOS DO PEDIDO", "ESTÁGIO PROBATÓRIO",
        "FUNDAMENTOS DECISÃO", "LOCAL DE RESIDÊNCIA",
    ],
    SHEET_SERVIDORES_ATIVOS: [
        "NOME", "MATRÍCULA", "CARGO", "FUNÇÃO", "LOTAÇÃO", "TIPO", "INÍCIO",
        "FINAL", "PROCESSO", "CONDIÇÕES ESPECIAIS", "REGIME", "META",
        "GÊNERO", "GRAU", "DOMICÍLIO", "CELULAR",
    ],
    SHEET_CNJ: [
        "PROCESSO ", "NOME", "MATRICULA", "CARGO EFETIVO", "CARGO FUNÇÃO",
        "LOTAÇÃO DE EXERCÍCIO", "LOTAÇÃO DE ORIGEM", "ESTÁGIO PROBATÓRIO ",
        "NATUREZA", "INÍCIO", "TÉRMINO", "MOTIVOS DO PEDIDO",
        "FUNDAMENTOS PARA O DEFERIMENTO", "LOTAÇÃO/ONDE VAI RESIDIR",
    ],
    SHEET_SERVIDORES_DESLIGADOS: [
        "NOME", "MATRÍCULA", "CARGO", "FUNÇÃO", "LOTAÇÃO", "TIPO", "INÍCIO",
        "FINAL", "PROCESSO", "COND ESPEC", "REGIME", "META", "GÊNERO",
        "GRAU", "DOMICÍLIO", "CELULAR", "MOTIVO",
    ],
    SHEET_PRODUTIVIDADE_LANCAMENTOS: [
        "SERVIDOR", "MATRÍCULA", "LOTAÇÃO", "MÊS/ANO", "META", "PRODUÇÃO",
        "PROCESSO", "OBSERVAÇÃO", "DATA_LANÇAMENTO", "USUÁRIO_LANÇAMENTO",
    ],
    SHEET_PONTO_TELETRABALHO: [
        "SERVIDOR", "MATRÍCULA", "DATA", "HORARIO", "TIPO", "OBSERVAÇÃO",
        "DATA_LANÇAMENTO", "USUÁRIO_LANÇAMENTO",
    ],
}

# Linha (1-indexed) onde começa o cabeçalho real de cada aba, na planilha
# de origem (algumas abas têm a linha 1 em branco antes do cabeçalho).
HEADER_ROW = {
    SHEET_MAGISTRADOS: 1,
    SHEET_SERVIDORES_ATIVOS: 1,
    SHEET_CNJ: 2,
    SHEET_SERVIDORES_DESLIGADOS: 1,
}


@st.cache_resource(show_spinner=False)
def get_client() -> gspread.Client:
    """Autentica via conta de serviço (credenciais em st.secrets)."""
    if "gcp_service_account" not in st.secrets:
        st.error(
            "⚠️ Credenciais do Google não configuradas neste ambiente.\n\n"
            "No Streamlit Cloud: abra o app → **⋮ (menu) → Settings → Secrets** "
            "e cole ali o conteúdo do arquivo `secrets.toml` (seção "
            "`[gcp_service_account]`), preenchido com os dados reais da conta "
            "de serviço do Google Cloud. O modelo está em `secrets_toml.example` "
            "no repositório."
        )
        st.stop()
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def get_spreadsheet() -> gspread.Spreadsheet:
    return get_client().open_by_key(SPREADSHEET_ID)


def ensure_id_column(ws: gspread.Worksheet, header_row: int) -> None:
    """Garante que a primeira coluna da aba se chame 'id'.
    Não reordena nem apaga nada existente — só insere a coluna quando ausente.
    """
    first_cell = ws.cell(header_row, 1).value
    if (first_cell or "").strip().lower() == "id":
        return
    ws.insert_cols([[]], col=1)
    ws.update_cell(header_row, 1, "id")
    # Numera as linhas de dado já existentes (aditivo, não altera outras colunas).
    n_rows = ws.row_count
    ids = [[i - header_row] for i in range(header_row + 1, n_rows + 1)]
    if ids:
        ws.update(
            f"A{header_row + 1}:A{n_rows}",
            ids,
            value_input_option="USER_ENTERED",
        )


def get_or_create_worksheet(name: str, headers: list[str]) -> gspread.Worksheet:
    sh = get_spreadsheet()
    try:
        ws = sh.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=name, rows=1000, cols=len(headers) + 2)
        ws.update(
            "A1",
            [["id"] + headers],
            value_input_option="USER_ENTERED",
        )
    return ws


def _cabecalho_unico(header: list[str]) -> list[str]:
    """Garante nomes de coluna únicos e não vazios, preservando ordem.
    Célula vazia na 1ª coluna vira 'NOME' (padrão observado nas abas de
    servidores). Demais vazias e nomes repetidos recebem sufixo numérico,
    sem alterar nada na planilha de origem — só na leitura em memória.
    """
    vistos: dict[str, int] = {}
    resultado = []
    for i, h in enumerate(header):
        nome = (h or "").strip()
        if not nome:
            nome = "NOME" if i == 0 else f"COLUNA_{i + 1}"
        base = nome
        if nome in vistos:
            vistos[nome] += 1
            nome = f"{base}_{vistos[nome]}"
        else:
            vistos[nome] = 0
        resultado.append(nome)
    return resultado


@st.cache_data(ttl=120, show_spinner=False)
def load_sheet_df(sheet_name: str, header_row: int = 1) -> pd.DataFrame:
    """Carrega 100% das linhas de uma aba tabular em DataFrame, sem amostrar."""
    ws = get_spreadsheet().worksheet(sheet_name)
    values = ws.get_all_values()
    if len(values) < header_row:
        return pd.DataFrame()
    header = _cabecalho_unico(values[header_row - 1])
    rows = values[header_row:]
    df = pd.DataFrame(rows, columns=header)
    # remove linhas 100% vazias
    df = df[~(df.apply(lambda r: all(str(v).strip() == "" for v in r), axis=1))]
    return df.reset_index(drop=True)


def next_id(df: pd.DataFrame, id_col: str = "id") -> int:
    if df.empty or id_col not in df.columns:
        return 1
    nums = pd.to_numeric(df[id_col], errors="coerce").dropna()
    return int(nums.max()) + 1 if not nums.empty else 1


def append_row(sheet_name: str, headers_with_id: list[str], row_dict: dict) -> int:
    """Acrescenta UMA linha ao final da aba, gerando o próximo id.
    Nunca sobrescreve linha existente — sempre append."""
    ws = get_spreadsheet().worksheet(sheet_name)
    values = ws.get_all_values()
    header_row_idx = 1
    for i, r in enumerate(values, start=1):
        if any(str(c).strip() for c in r):
            header_row_idx = i
            break
    header = values[header_row_idx - 1] if values else headers_with_id
    body = values[header_row_idx:]
    df_existing = pd.DataFrame(body, columns=header) if body else pd.DataFrame(columns=header)
    new_id = next_id(df_existing, id_col=header[0] if header else "id")
    row = [new_id] + [row_dict.get(h, "") for h in header[1:]]
    ws.append_row(row, value_input_option="USER_ENTERED")
    load_sheet_df.clear()
    return new_id


def timestamp() -> str:
    return dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
