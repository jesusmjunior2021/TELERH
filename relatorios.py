# MAT-TELETRABALHO-001 · relatorios.py
# Geração de relatórios com timbre institucional, em PDF (reportlab) e em
# HTML/CSS3 estruturado para impressão (Ctrl+P do navegador ou "Salvar como PDF").
# Nenhuma função aqui grava na planilha — é só formatação de saída.

from __future__ import annotations

import datetime as dt
import html
import io

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ORGAO = "DIRETORIA DE RH DO TJMA"
SISTEMA = "TELETRABALHO"

AZUL_PDF = colors.HexColor("#1B3A6B")
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
    cortada ou texto estourando a borda."""
    n = len(colunas)
    pesos = [1.3] + [1.0] * (n - 1) if n > 1 else [1.0]
    soma_pesos = sum(pesos)
    larguras = [largura_disponivel * p / soma_pesos for p in pesos]

    linha_cabecalho = [Paragraph(html.escape(str(c)), CABECALHO_CELULA_STYLE) for c in colunas]
    linhas = [linha_cabecalho]
    for _, linha in df[colunas].astype(str).iterrows():
        linhas.append(
            [Paragraph(html.escape(v) if v.strip() else "—", CELULA_STYLE) for v in linha]
        )

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


def gerar_pdf(nome_aba: str, df: pd.DataFrame, filtros: dict[str, str] | None = None) -> bytes:
    """Gera um PDF paisagem A4 com timbre institucional e a tabela de dados.
    Tabelas com muitas colunas são divididas em blocos legíveis (com quebra
    de página entre eles), cada um repetindo a coluna identificadora."""
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
            f"{len(df)} registro(s) — 100% dos dados filtrados, sem amostragem.",
            meta_style,
        )
    )
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
    elementos.append(Paragraph("ADMJESUSIA 107805", ParagraphStyle(

        "rodape", parent=styles["Normal"], fontSize=6, textColor=colors.HexColor("#B9C0CA"), alignment=2,
    )))

    doc.build(elementos)
    return buf.getvalue()


def gerar_html(nome_aba: str, df: pd.DataFrame, filtros: dict[str, str] | None = None) -> str:
    """Gera um HTML autônomo (CSS3 embutido, sem dependência externa) pronto
    para impressão direta do navegador ou 'Salvar como PDF'."""
    texto_filtros = _linha_filtros(filtros)
    bloco_filtros = (
        f"<div class='filtros'>Filtros aplicados: <strong>{html.escape(texto_filtros)}</strong></div>"
        if texto_filtros
        else ""
    )

    if df.empty:
        corpo = "<p class='vazio'>Nenhum registro encontrado para os filtros aplicados.</p>"
    else:
        thead = "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
        linhas_html = []
        for _, linha in df.iterrows():
            tds = "".join(
                f"<td>{html.escape(str(v)) if str(v).strip() else '—'}</td>" for v in linha
            )
            linhas_html.append(f"<tr>{tds}</tr>")
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
  <p class="meta">Emitido em {emitido} — {len(df)} registro(s) — 100% dos dados filtrados, sem amostragem.</p>
  <div class="tabela-wrap">
  {corpo}
  </div>
  <footer>ADMJESUSIA 107805</footer>
</body>
</html>
"""
