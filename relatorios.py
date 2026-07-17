# MAT-TELETRABALHO-001 · relatorios.py
# Geração de relatórios com timbre institucional, em PDF (reportlab) e em
# HTML/CSS3 estruturado para impressão (Ctrl+P do navegador ou "Salvar como PDF").
# Nenhuma função aqui grava na planilha — é só formatação de saída, sempre
# a partir dos dados exatos já carregados (nenhum valor é inventado; campos
# vazios/NaN aparecem como "—", nunca em branco silencioso nem como "nan").

from __future__ import annotations

import base64
import datetime as dt
import html
import io
import re

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# ─────────────────────── Utilitários de dados (embutidos) ─────────────────
# Antes viviam num módulo separado (utils_dados.py); trazidos para cá para
# que este arquivo não dependa de mais nenhum outro arquivo do projeto além
# dos pacotes instalados via requirements.txt.

MESES_PT = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}
_VAZIOS_TEXTUAIS = {"nan", "none", "nat", "<na>", ""}


def valor_texto(v) -> str:
    """Converte qualquer valor de célula (float NaN, None, número, string)
    em texto de exibição seguro — nunca deixa passar 'nan'/'None' cru, que é
    como o pandas mais novo representa ausência de dado ao converter tipos
    mistos. Números inteiros perdem o '.0' à toa (ex.: 100.0 -> '100')."""
    if v is None:
        return ""
    if isinstance(v, float):
        if pd.isna(v):
            return ""
        if v.is_integer():
            return str(int(v))
        return str(v)
    s = str(v).strip()
    return "" if s.lower() in _VAZIOS_TEXTUAIS else s


def chave_mes_ano(valor) -> tuple[int, int, str]:
    """Converte um texto tipo 'jun26', 'jun/2026', '06/2026' em chave
    ordenável (ano, mês, texto original). Formatos não reconhecidos vão
    para o final da lista, mantidos na ordem em que apareceram."""
    texto = valor_texto(valor).lower()
    if not texto:
        return (9999, 99, texto)

    m = re.match(r"([a-zç]{3,})\D*(\d{2,4})", texto)
    if m:
        mes = MESES_PT.get(m.group(1)[:3], 99)
        ano = int(m.group(2))
        ano = ano if ano > 100 else 2000 + ano
        return (ano, mes, texto)

    m2 = re.match(r"(\d{1,2})\D+(\d{2,4})", texto)
    if m2:
        mes = int(m2.group(1))
        ano = int(m2.group(2))
        ano = ano if ano > 100 else 2000 + ano
        if 1 <= mes <= 12:
            return (ano, mes, texto)

    return (9999, 99, texto)


def ordenar_por_mes_ano(df: pd.DataFrame, coluna: str = "mes_ano") -> pd.DataFrame:
    """Retorna o DataFrame ordenado cronologicamente pela coluna de mês/ano
    informada, sem alterar o DataFrame original."""
    if df.empty or coluna not in df.columns:
        return df
    df2 = df.copy()
    df2["_ordem_tmp"] = df2[coluna].map(chave_mes_ano)
    df2 = df2.sort_values("_ordem_tmp").drop(columns="_ordem_tmp")
    return df2.reset_index(drop=True)


def meses_ordenados(valores) -> list[str]:
    """Ordena uma lista/coleção de textos de mês/ano em ordem cronológica
    (para popular multiselect, eixo de gráfico etc.)."""
    return sorted({valor_texto(v) for v in valores if valor_texto(v)}, key=chave_mes_ano)


# ─────────────────────────── Geração de relatórios ────────────────────────

ORGAO = "DIRETORIA DE RH DO TJMA"
SISTEMA = "TELETRABALHO"

AZUL_PDF = colors.HexColor("#1B3A6B")
VERDE_PDF = colors.HexColor("#1E8F5F")
CINZA_PDF = colors.HexColor("#F2F3F5")
CINZA_TEXTO_PDF = colors.HexColor("#5A6472")

# Quantas colunas cabem, com quebra de texto, numa página A4 paisagem sem
# ficar ilegível. Tabelas com mais colunas que isso são divididas em blocos
# — cada bloco repete a 1ª coluna (identificador) para não perder o contexto
# de qual linha é qual.
MAX_COLS_POR_BLOCO = 6

CELULA_STYLE = ParagraphStyle(
    "celula", fontName="Helvetica", fontSize=6.6, leading=7.8, textColor=colors.HexColor("#1B1F24"),
)
CABECALHO_CELULA_STYLE = ParagraphStyle(
    "cabecalho_celula", fontName="Helvetica-Bold", fontSize=7, leading=8.2, textColor=colors.white,
)


def _linha_filtros(filtros: dict[str, str] | None) -> str:
    if not filtros:
        return ""
    partes = [f"{k}: {v}" for k, v in filtros.items() if str(v).strip()]
    return " · ".join(partes)


def _preparar_dados(df: pd.DataFrame) -> pd.DataFrame:
    """Prepara o DataFrame para emissão: ordena cronologicamente por
    mês/ano quando a coluna existir (nunca ordem alfabética de texto tipo
    'jun26'/'jan26'), e não altera nem inventa nenhum valor — só reordena
    linhas já existentes."""
    for candidato in ("mes_ano", "MÊS/ANO", "Mês/Ano"):
        if candidato in df.columns:
            return ordenar_por_mes_ano(df, candidato)
    return df


def _quebrar_colunas(colunas: list[str], max_por_bloco: int = MAX_COLS_POR_BLOCO) -> list[list[str]]:
    """Divide uma lista longa de colunas em blocos legíveis. A 1ª coluna
    (tratada como identificador — nome/servidor/processo) se repete em
    todos os blocos."""
    if len(colunas) <= max_por_bloco:
        return [colunas]
    identificador = colunas[0]
    resto = colunas[1:]
    passo = max(max_por_bloco - 1, 1)
    blocos = []
    for i in range(0, len(resto), passo):
        blocos.append([identificador] + resto[i : i + passo])
    return blocos


def _tabela_bloco(df: pd.DataFrame, colunas: list[str], largura_disponivel: float) -> Table:
    """Monta uma Table com célula em Paragraph (quebra automática de linha)
    e largura de coluna calculada para caber na página — nada de coluna
    cortada ou texto estourando a borda. Usa valor_texto() para sanitizar
    cada célula (NaN/None viram '—', nunca 'nan' cru)."""
    n = len(colunas)
    pesos = [1.3] + [1.0] * (n - 1) if n > 1 else [1.0]
    soma_pesos = sum(pesos)
    larguras = [largura_disponivel * p / soma_pesos for p in pesos]

    linha_cabecalho = [Paragraph(html.escape(str(c)), CABECALHO_CELULA_STYLE) for c in colunas]
    linhas = [linha_cabecalho]
    for _, linha in df[colunas].iterrows():
        celulas = []
        for v in linha:
            texto = valor_texto(v)
            celulas.append(Paragraph(html.escape(texto) if texto else "—", CELULA_STYLE))
        linhas.append(celulas)

    tabela = Table(linhas, colWidths=larguras, repeatRows=1)
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), AZUL_PDF),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, CINZA_PDF]),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#C7CDD6")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return tabela


def _altura_proporcional(png_bytes: bytes, largura_alvo: float) -> float:
    """Calcula a altura correta para exibir a imagem em `largura_alvo`
    mantendo a proporção REAL do PNG (lida via Pillow) — evita esticar ou
    achatar o gráfico quando o bbox_inches='tight' do matplotlib recorta a
    figura para um formato diferente do original."""
    try:
        from PIL import Image as PILImage

        with PILImage.open(io.BytesIO(png_bytes)) as img:
            largura_px, altura_px = img.size
        if largura_px <= 0:
            return largura_alvo * 0.31
        return largura_alvo * (altura_px / largura_px)
    except Exception:
        return largura_alvo * 0.31


def _grafico_produtividade_png(df: pd.DataFrame) -> bytes | None:
    """Gera um PNG simples de Meta x Produção por mês/ano (ordem cronológica),
    só quando as colunas existirem — usado nos relatórios de Produtividade.
    Retorna None silenciosamente para qualquer outra aba/formato de dado."""
    colunas_necessarias = {"mes_ano", "meta", "producao"}
    if not colunas_necessarias.issubset(set(df.columns)):
        return None

    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception:
        return None

    base = df[["mes_ano", "meta", "producao"]].copy()
    base["meta"] = pd.to_numeric(base["meta"], errors="coerce")
    base["producao"] = pd.to_numeric(base["producao"], errors="coerce")
    agregado = base.groupby("mes_ano", as_index=False)[["meta", "producao"]].sum(min_count=1)
    if agregado.empty:
        return None
    agregado = agregado.sort_values(
        by="mes_ano", key=lambda col: col.map(chave_mes_ano)
    ).reset_index(drop=True)

    fig, ax = plt.subplots(figsize=(10, 3.1), dpi=150)
    posicoes = range(len(agregado))
    largura = 0.38
    ax.bar(
        [i - largura / 2 for i in posicoes], agregado["meta"].fillna(0),
        width=largura, label="Meta", color="#1B3A6B",
    )
    ax.bar(
        [i + largura / 2 for i in posicoes], agregado["producao"].fillna(0),
        width=largura, label="Produção", color="#1E8F5F",
    )
    ax.set_xticks(list(posicoes))
    ax.set_xticklabels(agregado["mes_ano"], rotation=45, ha="right", fontsize=7)
    ax.tick_params(axis="y", labelsize=7)
    ax.legend(fontsize=8, frameon=False)
    ax.set_title("Meta x Produção por mês/ano (ordem cronológica)", fontsize=10, color="#1B3A6B")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", pad_inches=0.15)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def gerar_pdf(nome_aba: str, df: pd.DataFrame, filtros: dict[str, str] | None = None) -> bytes:
    """Gera um PDF paisagem A4 com timbre institucional, gráfico (quando
    aplicável) e a tabela de dados — exatamente os dados recebidos, na
    ordem cronológica quando há mês/ano. Tabelas com muitas colunas são
    divididas em blocos legíveis (com quebra de página entre eles)."""
    df = _preparar_dados(df)

    buf = io.BytesIO()
    pagesize = landscape(A4)
    margem_lateral = 1.0 * cm
    doc = SimpleDocTemplate(
        buf,
        pagesize=pagesize,
        topMargin=1.3 * cm,
        bottomMargin=1.1 * cm,
        leftMargin=margem_lateral,
        rightMargin=margem_lateral,
        title=f"{SISTEMA} - {nome_aba}",
    )
    largura_disponivel = pagesize[0] - 2 * margem_lateral

    styles = getSampleStyleSheet()
    titulo_style = ParagraphStyle(
        "titulo", parent=styles["Heading1"], textColor=AZUL_PDF, fontSize=14, spaceAfter=1,
    )
    sub_style = ParagraphStyle(
        "sub", parent=styles["Heading2"], textColor=CINZA_TEXTO_PDF, fontSize=11, spaceAfter=8,
    )
    meta_style = ParagraphStyle(
        "meta", parent=styles["Normal"], fontSize=8, textColor=CINZA_TEXTO_PDF, spaceAfter=2,
    )
    bloco_titulo_style = ParagraphStyle(
        "bloco_titulo", parent=styles["Normal"], fontSize=9, textColor=AZUL_PDF,
        spaceBefore=2, spaceAfter=6, fontName="Helvetica-Bold",
    )

    elementos = [
        Paragraph(ORGAO, titulo_style),
        Paragraph(f"{SISTEMA} — {nome_aba}", sub_style),
    ]
    texto_filtros = _linha_filtros(filtros)
    if texto_filtros:
        elementos.append(Paragraph(f"Filtros aplicados: {html.escape(texto_filtros)}", meta_style))
    elementos.append(
        Paragraph(
            f"Emitido em {dt.datetime.now().strftime('%d/%m/%Y %H:%M')} — "
            f"{len(df)} registro(s) — 100% dos dados filtrados, sem amostragem, "
            f"exatamente como constam na planilha.",
            meta_style,
        )
    )
    elementos.append(Spacer(1, 8))

    grafico_png = _grafico_produtividade_png(df) if not df.empty else None
    if grafico_png:
        altura_grafico = _altura_proporcional(grafico_png, largura_disponivel)
        elementos.append(Image(io.BytesIO(grafico_png), width=largura_disponivel, height=altura_grafico))
        elementos.append(Spacer(1, 10))

    if df.empty:
        elementos.append(Paragraph("Nenhum registro encontrado para os filtros aplicados.", styles["Normal"]))
    else:
        colunas = [str(c) for c in df.columns]
        blocos = _quebrar_colunas(colunas)
        for i, bloco in enumerate(blocos):
            if len(blocos) > 1:
                elementos.append(
                    Paragraph(
                        f"Bloco {i + 1} de {len(blocos)} de colunas "
                        f"({bloco[0]} + {len(bloco) - 1} campo(s))",
                        bloco_titulo_style,
                    )
                )
            elementos.append(_tabela_bloco(df, bloco, largura_disponivel))
            if i < len(blocos) - 1:
                elementos.append(PageBreak())

    elementos.append(Spacer(1, 14))
    elementos.append(
        Paragraph(
            "ADMJESUSIA 107805",
            ParagraphStyle(
                "rodape", parent=styles["Normal"], fontSize=6,
                textColor=colors.HexColor("#B9C0CA"), alignment=2,
            ),
        )
    )

    doc.build(elementos)
    return buf.getvalue()


def gerar_html(nome_aba: str, df: pd.DataFrame, filtros: dict[str, str] | None = None) -> str:
    """Gera um HTML autônomo (CSS3 embutido, sem dependência externa) pronto
    para impressão direta do navegador ou 'Salvar como PDF'. Mesma ordenação
    cronológica e mesmo gráfico do PDF, dados exatos, sem inventar nada."""
    df = _preparar_dados(df)

    texto_filtros = _linha_filtros(filtros)
    bloco_filtros = (
        f"<div class='filtros'>Filtros aplicados: <strong>{html.escape(texto_filtros)}</strong></div>"
        if texto_filtros
        else ""
    )

    grafico_html = ""
    grafico_png = _grafico_produtividade_png(df) if not df.empty else None
    if grafico_png:
        b64 = base64.b64encode(grafico_png).decode("ascii")
        grafico_html = (
            f"<img class='grafico' src='data:image/png;base64,{b64}' "
            f"alt='Gráfico Meta x Produção por mês/ano'/>"
        )

    if df.empty:
        corpo = "<p class='vazio'>Nenhum registro encontrado para os filtros aplicados.</p>"
    else:
        thead = "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
        linhas_html = []
        for _, linha in df.iterrows():
            tds = []
            for v in linha:
                texto = valor_texto(v)
                tds.append(f"<td>{html.escape(texto) if texto else '—'}</td>")
            linhas_html.append(f"<tr>{''.join(tds)}</tr>")
        corpo = f"<table><thead><tr>{thead}</tr></thead><tbody>{''.join(linhas_html)}</tbody></table>"

    emitido = dt.datetime.now().strftime("%d/%m/%Y %H:%M")
    nome_aba_esc = html.escape(nome_aba)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>{html.escape(SISTEMA)} — {nome_aba_esc}</title>
<style>
  @page {{ size: A4 landscape; margin: 14mm; }}
  * {{ box-sizing: border-box; }}
  body {{
    font-family: "Segoe UI", Arial, sans-serif;
    color: #1B1F24;
    margin: 0;
    padding: 24px;
    background: #FFFFFF;
  }}
  header {{
    border-bottom: 3px solid #1B3A6B;
    padding-bottom: 10px;
    margin-bottom: 14px;
  }}
  header .orgao {{
    color: #1B3A6B;
    font-size: 19px;
    font-weight: 700;
    letter-spacing: 0.4px;
    margin: 0;
  }}
  header .sistema {{
    color: #5A6472;
    font-size: 13px;
    margin: 3px 0 0 0;
    text-transform: uppercase;
    letter-spacing: 1.2px;
  }}
  .filtros {{
    font-size: 11px;
    color: #3D6DB5;
    margin: 10px 0 2px 0;
  }}
  .meta {{
    font-size: 10px;
    color: #8A93A3;
    margin-bottom: 14px;
  }}
  .grafico {{
    display: block;
    width: 100%;
    max-width: 950px;
    margin: 4px 0 18px 0;
  }}
  .tabela-wrap {{
    width: 100%;
    overflow-x: auto;
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 10.5px;
  }}
  thead th {{
    background: #1B3A6B;
    color: #FFFFFF;
    text-align: left;
    padding: 7px 9px;
    white-space: nowrap;
  }}
  tbody td {{
    padding: 6px 9px;
    border-bottom: 1px solid #E4E7EB;
  }}
  tbody tr:nth-child(even) {{ background: #F2F3F5; }}
  .vazio {{ color: #5A6472; font-style: italic; }}
  footer {{
    margin-top: 20px;
    font-size: 9px;
    color: #B9C0CA;
    text-align: right;
  }}
  .barra-acoes {{
    margin-bottom: 14px;
  }}
  .barra-acoes button {{
    background: #1B3A6B;
    color: #FFFFFF;
    border: none;
    padding: 8px 16px;
    border-radius: 6px;
    font-size: 12px;
    cursor: pointer;
  }}
  @media print {{
    .barra-acoes {{ display: none; }}
    tbody tr {{ page-break-inside: avoid; }}
    body {{ padding: 0; }}
    .tabela-wrap {{ overflow-x: visible; }}
    table {{ font-size: 8px; }}
    thead th, tbody td {{ padding: 4px 5px; white-space: normal; }}
  }}
</style>
</head>
<body>
  <div class="barra-acoes">
    <button onclick="window.print()">🖨️ Imprimir / Salvar como PDF</button>
  </div>
  <header>
    <p class="orgao">{html.escape(ORGAO)}</p>
    <p class="sistema">{html.escape(SISTEMA)} — {nome_aba_esc}</p>
  </header>
  {bloco_filtros}
  <p class="meta">Emitido em {emitido} — {len(df)} registro(s) — 100% dos dados filtrados, sem amostragem, exatamente como constam na planilha.</p>
  {grafico_html}
  <div class="tabela-wrap">
  {corpo}
  </div>
  <footer>ADMJESUSIA 107805</footer>
</body>
</html>
"""
