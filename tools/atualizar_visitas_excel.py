# -*- coding: utf-8 -*-
"""Atualiza `records` e `stores` do Historico da Loja a partir do export
S3 INSTRUTORES - TREINAMENTO.xlsx (aba BASE), no historico_da_loja.html
(fonte) e no public/index.html (publicado, preservando as lojas do
roteiro que nao tem nenhum treinamento ainda)."""
import json
import re
from pathlib import Path

import pandas as pd

ROOT = Path(r"C:\dev\historico loja")
SRC_HTML = ROOT / "historico_da_loja.html"
PUB_HTML = ROOT / "public" / "index.html"
XLSX = r"\\servidor\pasta 194300 - ROKU\2026\2. RELATÓRIOS\Bases\S3 INSTRUTORES - TREINAMENTO.xlsx"

COL_MAP = {
    "ID": "id",
    "COD_LOJA": "cod_loja",
    "LOJA": "loja",
    "CIDADE": "cidade",
    "DES_UF": "uf",
    "REDE": "rede",
    "BANDEIRA": "bandeira",
    "CANAL": "canal",
    "END_LOGRADOURO": "endereco",
    "TITULO": "titulo",
    "CATEGORIAS": "categorias",
    "DES_STATUS": "status",
    "INSTRUTOR": "instrutor",
    "DES_FORMATO": "formato",
    "TREINAMENTO_REALIZADO": "realizado",
    "PESSOAS_PREVISTAS": "previstas",
    "QTD_PESSOAS_TREINADA": "treinadas",
    "MÉDIA %": "media",
    "OBSERVACOES": "observacoes",
    "PROGRAMADO": "programado",
}


def to_date_str(v):
    if v is None or pd.isna(v):
        return None
    return pd.Timestamp(v).strftime("%Y-%m-%d")


def main():
    df = pd.read_excel(XLSX, sheet_name="BASE")
    total_raw = len(df)
    df = df[df["COD_LOJA"].notna()].copy()
    dropped_no_loja = total_raw - len(df)

    df["_data"] = df["DATA_VISTA"]
    df.loc[df["_data"].isna(), "_data"] = df["DTA_ROTEIRO"]

    rows = df.to_dict("records")
    records = []
    for r in rows:
        rec_data = to_date_str(r.get("_data"))
        rec = {dst: (None if pd.isna(r.get(src)) else r.get(src)) for src, dst in COL_MAP.items()}
        def num(v, cast):
            if v is None:
                return None
            s = str(v).strip()
            return cast(s) if s else None

        rec["cod_loja"] = int(rec["cod_loja"])
        rec["previstas"] = num(rec.get("previstas"), float)
        rec["previstas"] = int(rec["previstas"]) if rec["previstas"] is not None else 0
        rec["treinadas"] = num(rec.get("treinadas"), float)
        rec["treinadas"] = int(rec["treinadas"]) if rec["treinadas"] is not None else 0
        rec["media"] = num(rec.get("media"), float)
        for k in ("id", "loja", "cidade", "uf", "rede", "bandeira", "canal", "endereco",
                  "titulo", "categorias", "status", "instrutor", "formato", "realizado",
                  "observacoes", "programado"):
            v = rec.get(k)
            rec[k] = None if v is None else str(v)
        rec["data"] = rec_data
        records.append(rec)

    print(f"linhas na planilha: {total_raw} | sem COD_LOJA (descartadas): {dropped_no_loja} | records finais: {len(records)}")

    # ---------- stores: 1 por COD_LOJA, mantendo o registro mais recente ----------
    by_store = {}
    for rec in records:
        cod = rec["cod_loja"]
        prev = by_store.get(cod)
        if prev is None or (rec["data"] or "") >= (prev["data"] or ""):
            by_store[cod] = rec
    stores = [
        {
            "cod_loja": r["cod_loja"],
            "loja": r["loja"],
            "cidade": r["cidade"],
            "uf": r["uf"],
            "rede": r["rede"],
            "bandeira": r["bandeira"],
            "canal": r["canal"],
            "endereco": r["endereco"],
        }
        for r in by_store.values()
    ]
    stores.sort(key=lambda s: s["cod_loja"])
    print("stores unicos:", len(stores))

    records_json = json.dumps(records, ensure_ascii=False, separators=(",", ":"))
    stores_json = json.dumps(stores, ensure_ascii=False, separators=(",", ":"))

    ANCHOR_START = '<div id="toast" class="toast"></div>\n\n<script>\n'
    ANCHOR_END = "/*__ATIVACOES_START__*/"

    def replace_block(content, var_kw, extra_stores_json):
        i0 = content.index(ANCHOR_START) + len(ANCHOR_START)
        i1 = content.index(ANCHOR_END)
        assert i0 < i1, "ancoras fora de ordem"
        block = f"{var_kw} records={records_json};\n{var_kw} stores={extra_stores_json};\n"
        return content[:i0] + block + content[i1:]

    # ---------- historico_da_loja.html (fonte) ----------
    content = SRC_HTML.read_text(encoding="utf-8")
    content = replace_block(content, "const", stores_json)
    SRC_HTML.write_text(content, encoding="utf-8")
    print("OK ->", SRC_HTML)

    # ---------- public/index.html (recalcula lojas do roteiro sem treinamento) ----------
    ROTEIRO_XLSX = r"C:\Users\usuario\AppData\Local\Temp\roteiro_test.xlsx"
    new_codes = {s["cod_loja"] for s in stores}
    extra_from_roteiro = []
    ro_wb = pd.ExcelFile(ROTEIRO_XLSX)
    ro_df = ro_wb.parse("Roteiro")
    seen = set()
    for _, row in ro_df.iterrows():
        cod = row.get("COD_LOJA")
        if pd.isna(cod):
            continue
        cod = int(cod)
        if cod in new_codes or cod in seen:
            continue
        seen.add(cod)
        extra_from_roteiro.append({
            "cod_loja": cod,
            "loja": row.get("NOME LOJA"),
            "cidade": row.get("CIDADE"),
            "uf": row.get("UF"),
            "rede": row.get("BANDEIRA"),
            "bandeira": row.get("BANDEIRA"),
            "canal": row.get("CANAL"),
            "endereco": row.get("ENDEREÇO"),
        })
    stores_pub = stores + extra_from_roteiro
    stores_pub.sort(key=lambda s: s["cod_loja"])
    print("lojas do roteiro preservadas (sem treinamento ainda):", len(extra_from_roteiro))
    stores_pub_json = json.dumps(stores_pub, ensure_ascii=False, separators=(",", ":"))

    pub = PUB_HTML.read_text(encoding="utf-8")

    pub = replace_block(pub, "let", stores_pub_json)
    PUB_HTML.write_text(pub, encoding="utf-8")
    print("OK ->", PUB_HTML)


if __name__ == "__main__":
    main()
