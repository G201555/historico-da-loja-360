# -*- coding: utf-8 -*-
"""
Atualiza o bloco `const ativacoes={...}` do Historico da Loja a partir de
tres planilhas exportadas (Rack Roku TV, Exposicao Especial, Movimentacao
de Testeira), em vez de consultar o bi_cliente ao vivo.

Mesma logica/campos do tools/montar_historico_loja.ps1, so que a fonte
dos dados agora e Excel. Loja e casada pelo nome normalizado (campo LOJA
da planilha == campo `loja` do array `stores` do app).
"""
import json
import re
import unicodedata
from pathlib import Path

import pandas as pd

ROOT = Path(r"C:\dev\historico loja")
HTML_FILES = [ROOT / "historico_da_loja.html", ROOT / "public" / "index.html"]

BASES = Path(r"K:\JOB 194300 - ROKU\2026\2. RELATÓRIOS\Bases")
RACK_XLSX = BASES / "RACK ROKU TV.xlsx"
EXP_XLSX = BASES / "Exposição Especial.xlsx"
TEST_XLSX = BASES / "MOVIMENTAÇÃO DE TESTEIRA.xlsx"


def norm(s):
    if s is None:
        return ""
    s = str(s).strip().upper()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s)
    return s


def fix_foto(v):
    s = "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
    if s.lower().startswith("http://"):
        return "https://" + s[7:]
    return s


def s(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def data_str(v):
    if v is None or pd.isna(v):
        return ""
    return pd.Timestamp(v).strftime("%Y-%m-%d")


def load_stores():
    with open(HTML_FILES[0], "r", encoding="utf-8") as f:
        for line in f:
            if line.startswith("const stores="):
                body = line.strip()
                return json.loads(body[len("const stores="):-1])
    raise RuntimeError("stores nao encontrado")


def build_store_index(stores):
    idx = {}
    for st in stores:
        idx.setdefault(norm(st["loja"]), []).append(st["cod_loja"])
    return idx


def first_nonempty(row, cols):
    for c in cols:
        if c in row and s(row[c]):
            return s(row[c])
    return ""


def process_sheet(path, sheet, status_col, extra):
    df = pd.read_excel(path, sheet_name=sheet)
    df = df[df[status_col].notna()].copy()
    df["_data"] = pd.to_datetime(df["DATA_VISITA"], errors="coerce")
    df = df.sort_values("_data", ascending=False)
    return df


def main():
    stores = load_stores()
    store_idx = build_store_index(stores)
    print(f"stores carregados: {len(stores)} ({len(store_idx)} nomes unicos)")

    ativacoes = {}
    unmatched = {"rack": set(), "exposicao": set(), "testeira": set()}
    stats = {}

    # ---------- Rack Roku TV ----------
    rack_foto_cols = [
        "Foto do Rack Finalizado:", "Foto do Rack",
        "Foto do Rack no Ambiente:", "Foto do Rack Ambiente",
        "Foto do Local ANTES",
    ]
    df = process_sheet(RACK_XLSX, "BASE", "Qual a situação atual do Rack?", None)
    codigo_cols = ["Código de Identificação do Rack Implementado:", "Código de Identificação do Rack:"]
    seen = set()
    rack_foto_fallback = {}
    for _, row in df.iterrows():
        key = norm(row["LOJA"])
        cods = store_idx.get(key)
        foto = fix_foto(first_nonempty(row, rack_foto_cols))
        if cods:
            for cod in cods:
                if foto and cod not in rack_foto_fallback:
                    rack_foto_fallback[cod] = foto
        else:
            unmatched["rack"].add(s(row["LOJA"]))
            continue
        if key in seen:
            continue
        seen.add(key)
        situacao = s(row["Qual a situação atual do Rack?"])
        entry = {
            "situacao": situacao,
            "presente": situacao == "OK, em perfeito estado.",
            "tipo": s(row.get("Qual modelo de Rack está sendo instalado?")),
            "codigo": first_nonempty(row, codigo_cols),
            "data": data_str(row["_data"]),
            "foto": foto,
        }
        for cod in cods:
            ativacoes.setdefault(str(cod), {})["rack"] = dict(entry)
    for cod, foto in rack_foto_fallback.items():
        e = ativacoes.get(str(cod), {}).get("rack")
        if e and not e["foto"]:
            e["foto"] = foto
    stats["rack"] = sum(1 for v in ativacoes.values() if "rack" in v)

    # ---------- Exposicao Especial ----------
    exp_foto_cols = [
        "Foto da Exposição Especial FINALIZADA:", "Foto da Exposição Especial",
        "Foto da Exposição Especial - Ambiente", "Foto do local ANTES da criação da exposição:",
        "Envie uma foto do local onde a exposição foi negociada:",
        "Envie uma foto da avaria/problema:", "Foto do ambiente sem a Exposição Especial",
    ]
    df = process_sheet(EXP_XLSX, "Base", "Qual a situação atual da Exposição Especial?", None)
    seen = set()
    exp_foto_fallback = {}
    for _, row in df.iterrows():
        key = norm(row["LOJA"])
        cods = store_idx.get(key)
        foto = fix_foto(first_nonempty(row, exp_foto_cols))
        if cods:
            for cod in cods:
                if foto and cod not in exp_foto_fallback:
                    exp_foto_fallback[cod] = foto
        else:
            unmatched["exposicao"].add(s(row["LOJA"]))
            continue
        if key in seen:
            continue
        seen.add(key)
        situacao = s(row["Qual a situação atual da Exposição Especial?"])
        entry = {
            "situacao": situacao,
            "presente": situacao != "Foi removida/descontinuada pela loja.",
            "tipo": s(row.get("Qual tipo de Exposição Especial foi criada?")),
            "data": data_str(row["_data"]),
            "foto": foto,
        }
        for cod in cods:
            ativacoes.setdefault(str(cod), {})["exposicao"] = dict(entry)
    for cod, foto in exp_foto_fallback.items():
        e = ativacoes.get(str(cod), {}).get("exposicao")
        if e and not e["foto"]:
            e["foto"] = foto
    stats["exposicao"] = sum(1 for v in ativacoes.values() if "exposicao" in v)

    # ---------- Testeira ----------
    df = process_sheet(TEST_XLSX, "file-04.08.26", "TIPO DE AÇÃO", None)
    seen = set()
    test_foto_fallback = {}
    for _, row in df.iterrows():
        key = norm(row["LOJA"])
        cods = store_idx.get(key)
        foto = fix_foto(s(row.get("FOTO DA TESTEIRA")))
        if cods:
            for cod in cods:
                if foto and cod not in test_foto_fallback:
                    test_foto_fallback[cod] = foto
        else:
            unmatched["testeira"].add(s(row["LOJA"]))
            continue
        if key in seen:
            continue
        seen.add(key)
        acao = s(row["TIPO DE AÇÃO"])
        entry = {
            "acao": acao,
            "presente": acao.upper().startswith("INSTALA"),
            "parceiro": s(row.get("QUAL PARCEIRO VOCÊ POSITIVOU?")),
            "quantidade": s(row.get("QUANTIDADE INSTALADA EM LOJA")),
            "data": data_str(row["_data"]),
            "foto": foto,
        }
        for cod in cods:
            ativacoes.setdefault(str(cod), {})["testeira"] = dict(entry)
    for cod, foto in test_foto_fallback.items():
        e = ativacoes.get(str(cod), {}).get("testeira")
        if e and not e["foto"]:
            e["foto"] = foto
    stats["testeira"] = sum(1 for v in ativacoes.values() if "testeira" in v)

    print(f"rack: {stats['rack']} lojas | exposicao: {stats['exposicao']} lojas | testeira: {stats['testeira']} lojas")
    for tipo, names in unmatched.items():
        if names:
            print(f"NAO CASADOS ({tipo}, {len(names)}):")
            for n in sorted(names):
                print("  -", n)

    js_block = "const ativacoes=" + json.dumps(ativacoes, ensure_ascii=False, separators=(",", ":")) + ";"
    start = "/*__ATIVACOES_START__*/"
    end = "/*__ATIVACOES_END__*/"
    wrapped = f"{start}\n{js_block}\n{end}"
    pattern = re.compile(re.escape(start) + r".*?" + re.escape(end), re.DOTALL)

    for html_path in HTML_FILES:
        content = html_path.read_text(encoding="utf-8")
        if start in content:
            content = pattern.sub(lambda m: wrapped, content, count=1)
        else:
            anchor = "const byStore=new Map();"
            if anchor not in content:
                raise RuntimeError(f"ponto de insercao nao encontrado em {html_path}")
            content = content.replace(anchor, f"{wrapped}\n{anchor}", 1)
        html_path.write_text(content, encoding="utf-8")
        print("OK ->", html_path)


if __name__ == "__main__":
    main()
