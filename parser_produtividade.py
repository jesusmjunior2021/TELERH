# MAT-TELETRABALHO-001 · parser_produtividade.py
#
# A aba "produtividade" da planilha de origem NÃO é uma tabela retangular:
# é um relatório manual, em blocos repetidos por servidor. Cada bloco tem:
#   1) uma linha de cabeçalho de período  (MATRÍCULA | PRAZO INICIAL | PRAZO
#      FINAL | PROCESSO), que só aparece por completo na abertura de um novo
#      servidor;
#   2) uma linha de dado do período (nome, matrícula, prazo inicial, prazo
#      final, processo do período vigente);
#   3) um sub-cabeçalho (MÊS/ANO | META | PRODUÇÃO | PROCESSO);
#   4) N linhas de série mensal (mês/ano, meta, produção, processo do mês,
#      às vezes uma observação na coluna seguinte).
# Um mesmo servidor pode ter VÁRIOS períodos (renovações), cada um com sua
# própria série mensal, antes de abrir o bloco do próximo servidor.
#
# Regra dura: nunca inventar linha nem completar lacuna. Mês sem meta/produção
# lançada aparece como NaN, não como zero.

from __future__ import annotations

import re
from typing import Optional

import pandas as pd

MES_RE = re.compile(r"^[a-zç]{3}\d{2}$", re.IGNORECASE)


def _is_month_token(v: str) -> bool:
    return bool(MES_RE.match((v or "").strip()))


def parse_produtividade(values: list[list[str]]) -> pd.DataFrame:
    """Recebe ws.get_all_values() da aba 'produtividade' e devolve um
    DataFrame normalizado (uma linha por servidor x período x mês), 100%
    fiel ao que está na planilha — sem preencher lacuna por inferência.
    """
    records = []
    cur_nome = None
    cur_matricula = None
    cur_lotacao = None
    cur_meta_declarada = None
    cur_prazo_inicial = None
    cur_prazo_final = None
    cur_processo_periodo = None
    period_id = 0

    n_rows = len(values)
    for i, row in enumerate(values):
        row = row + [""] * (30 - len(row)) if len(row) < 30 else row
        a, b, c, d, e, f, g = (row[0], row[1], row[2], row[3], row[4], row[5], row[6] if len(row) > 6 else "")

        # Abertura de bloco de NOVO servidor: cabeçalho completo de período.
        if b.strip().upper() == "MATRÍCULA" and c.strip().upper() == "PRAZO INICIAL":
            # a próxima linha não vazia traz nome/matrícula/prazo/processo do 1º período
            if i + 1 < n_rows:
                nrow = values[i + 1]
                cur_nome = (nrow[0] or "").strip() or None
                cur_matricula = (nrow[1] or "").strip() or None
                cur_prazo_inicial = (nrow[2] or "").strip() or None
                cur_prazo_final = (nrow[3] or "").strip() or None
                cur_processo_periodo = (nrow[4] or "").strip() or None
                period_id += 1
            continue

        # Novo período (renovação) do MESMO servidor: cabeçalho parcial,
        # sem repetir "MATRÍCULA" mas repetindo "PRAZO INICIAL".
        if c.strip().upper() == "PRAZO INICIAL" and d.strip().upper() == "PRAZO FINAL" and b.strip().upper() != "MATRÍCULA":
            if i + 1 < n_rows:
                nrow = values[i + 1]
                # texto de meta às vezes vem na coluna A desta linha de dado
                meta_txt = (nrow[0] or "").strip() or None
                cur_meta_declarada = meta_txt
                cur_prazo_inicial = (nrow[2] or "").strip() or None
                cur_prazo_final = (nrow[3] or "").strip() or None
                cur_processo_periodo = (nrow[4] or "").strip() or None
                period_id += 1
            continue

        # Sub-cabeçalho da série mensal.
        if c.strip().upper() == "MÊS/ANO":
            continue

        # Linha de série mensal: coluna C é um token tipo "jun26".
        if _is_month_token(c):
            # texto solto na coluna A, quando presente, costuma ser
            # lotação/meta/cargo do bloco — guardamos como contexto, não
            # como dado mensal.
            if a.strip() and not _is_month_token(a):
                if a.strip().lower().startswith("meta"):
                    cur_meta_declarada = a.strip()
                elif "vara" in a.strip().lower() or "comarca" in a.strip().lower():
                    cur_lotacao = a.strip()

            mes_ano = c.strip()
            meta = d.strip() if d.strip() else None
            producao = e.strip() if e.strip() else None
            processo_mes = f.strip() if f.strip() else None
            obs = g.strip() if g.strip() else None

            if meta is None and producao is None and processo_mes is None:
                # linha de mês futuro/sem lançamento ainda — mantém como
                # lacuna explícita, não descarta o mês.
                pass

            records.append(
                {
                    "servidor": cur_nome,
                    "matricula": cur_matricula,
                    "lotacao": cur_lotacao,
                    "periodo_id": period_id,
                    "prazo_inicial": cur_prazo_inicial,
                    "prazo_final": cur_prazo_final,
                    "processo_periodo": cur_processo_periodo,
                    "meta_declarada_periodo": cur_meta_declarada,
                    "mes_ano": mes_ano,
                    "meta": meta,
                    "producao": producao,
                    "processo_mes": processo_mes,
                    "observacao": obs,
                }
            )
            continue

        # Linha totalmente vazia = separador de bloco; não reseta o servidor
        # corrente (o próximo período do mesmo servidor pode vir a seguir).

    df = pd.DataFrame.from_records(records)
    if df.empty:
        return df

    df["meta"] = pd.to_numeric(df["meta"], errors="coerce")
    df["producao"] = pd.to_numeric(df["producao"], errors="coerce")
    df["superavit"] = df["producao"] - df["meta"]
    df["id"] = range(1, len(df) + 1)
    cols = ["id"] + [c for c in df.columns if c != "id"]
    return df[cols]
