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
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ORGAO = "DIRETORIA DE RH DO TJMA"
SISTEMA = "TELETRABALHO"

AZUL_PDF = colors.HexColor("#1B3A6B")
CINZA_PDF = colors.HexColor("#F2F3F5")
CINZA_TEXTO_PDF = colors.HexColor("#5A6472")


def _linha_filtros(filtros: dict[str, str] | None) -> str:
    if not filtros:
        return ""
    partes = [f"{k}: {v}" for k, v in filtros.items() if str(v).strip()]
    return " · ".join(partes)


def gerar_pdf(nome_aba: str, df: pd.DataFrame, filtros: dict[str, str] | None = None) -> bytes:
    """Gera um PDF paisagem A4 com timbre institucional e a tabela de dados."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        topMargin=1.4 * cm,
        bottomMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        rightMargin=1.2 * cm,
        title=f"{SISTEMA} - {nome_aba}",
    )
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
        cabecalho = [str(c) for c in df.columns]
        corpo = df.astype(str).values.tolist()
        dados = [cabecalho] + corpo
        tabela = Table(dados, repeatRows=1)
        tabela.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), AZUL_PDF),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 6.5),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, CINZA_PDF]),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#C7CDD6")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        elementos.append(tabela)

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
  {corpo}
  <footer>ADMJESUSIA 107805</footer>
</body>
</html>
"""
