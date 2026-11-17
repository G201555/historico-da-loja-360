# -*- coding: utf-8 -*-
"""
Rotina semanal de atualizacao do Roku 360 - Historico da Loja.

Atualiza, nessa ordem:
  1) records/stores  <- S3 INSTRUTORES - TREINAMENTO.xlsx (aba BASE)
  2) ativacoes        <- RACK ROKU TV.xlsx / Exposicao Especial.xlsx / MOVIMENTACAO DE TESTEIRA.xlsx
  3) roteiro + lojas do roteiro sem treinamento <- Analise de Roteiro Q3.xlsx (aba Roteiro, so leitura)

Depois valida que os 2 arquivos HTML continuam com JSON valido e que o numero
de registros nao caiu de forma suspeita (>50%) antes de publicar no Firebase.
Se qualquer etapa falhar ou a validacao nao passar, a rotina para SEM publicar
e grava o motivo no log - nada quebrado vai pro ar sozinho.

Uso:
    python tools\\atualizar_tudo.py
"""
import datetime
import json
import re
import subprocess
import sys
import unicodedata
from pathlib import Path

import pandas as pd
import openpyxl

ROOT = Path(r"C:\dev\historico loja")
SRC_HTML = ROOT / "historico_da_loja.html"
PUB_HTML = ROOT / "public" / "index.html"
LOG_DIR = ROOT / "tools" / "logs"

S3_XLSX = r"\\servidor\pasta 194300 - ROKU\2026\2. RELATÓRIOS\Bases\S3 INSTRUTORES - TREINAMENTO.xlsx"
BASES = Path(r"K:\JOB 194300 - ROKU\2026\2. RELATÓRIOS\Bases")
RACK_XLSX = BASES / "RACK ROKU TV.xlsx"
EXP_XLSX = BASES / "Exposição Especial.xlsx"
TEST_XLSX = BASES / "MOVIMENTAÇÃO DE TESTEIRA.xlsx"
ROTEIRO_XLSX = r"C:\Users\usuario\Downloads\Analise de Roteiro Q3.xlsx"

ANCHOR_TOAST = '<div id="toast" class="toast"></div>\n\n<script>\n'
ANCHOR_ATIVACOES_START = "/*__ATIVACOES_START__*/"
ANCHOR_ATIVACOES_END = "/*__ATIVACOES_END__*/"
ANCHOR_ROTEIRO_START = "/*__ROTEIRO_START__*/"
ANCHOR_ROTEIRO_END = "/*__ROTEIRO_END__*/"

COL_MAP = {
    "ID": "id", "COD_LOJA": "cod_loja", "LOJA": "loja", "CIDADE": "cidade",
    "DES_UF": "uf", "REDE": "rede", "BANDEIRA": "bandeira", "CANAL": "canal",
    "END_LOGRADOURO": "endereco", "TITULO": "titulo", "CATEGORIAS": "categorias",
    "DES_STATUS": "status", "INSTRUTOR": "instrutor", "DES_FORMATO": "formato",
    "TREINAMENTO_REALIZADO": "realizado", "PESSOAS_PREVISTAS": "previstas",
    "QTD_PESSOAS_TREINADA": "treinadas", "MÉDIA %": "media",
    "OBSERVACOES": "observacoes", "PROGRAMADO": "programado",
}

ROTEIRO_ALIAS = {
    "DAIANE GERONIMO NOE": "DAIANE GERONIMO NOÉ",
    "DIEGO FONSECA CRUZ": "consultor",
    "HENRIQUE AMORIM": "HENRIQUE MARQUES DA LUZ AMORIM",
    "ANTONIA TEOMARIA DOS SANTOS": "ANTONIA TEOMÁRIA PEREIRA DOS SANTOS",
}

log_lines = []


def log(msg):
    line = f"{datetime.datetime.now().strftime('%H:%M:%S')}  {msg}"
    print(line)
    log_lines.append(line)


def norm(s):
    if s is None:
        return ""
    s = str(s).strip().upper()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s)


def s_(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def data_str(v):
    if v is None or pd.isna(v):
        return ""
    return pd.Timestamp(v).strftime("%Y-%m-%d")


def replace_between(content, start, end, inner):
    wrapped = f"{start}\n{inner}\n{end}"
    if start in content:
        pattern = re.compile(re.escape(start) + r".*?" + re.escape(end), re.DOTALL)
        return pattern.sub(lambda m: wrapped, content, count=1)
    anchor = "const byStore=new Map();"
    if anchor not in content:
        raise RuntimeError("ponto de insercao (const byStore) nao encontrado")
    return content.replace(anchor, f"{wrapped}\n{anchor}", 1)


def extract_json_block(content, pattern):
    m = re.search(pattern, content, re.DOTALL)
    if not m:
        raise RuntimeError(f"bloco nao encontrado: {pattern}")
    return json.loads(m.group(1))


# ============================================================
# 1) records / stores
# ============================================================
def build_records_stores():
    df = pd.read_excel(S3_XLSX, sheet_name="BASE")
    total_raw = len(df)
    df = df[df["COD_LOJA"].notna()].copy()
    dropped = total_raw - len(df)
    df["_data"] = df["DATA_VISTA"]
    df.loc[df["_data"].isna(), "_data"] = df["DTA_ROTEIRO"]

    def num(v, cast):
        if v is None:
            return None
        st = str(v).strip()
        return cast(st) if st else None

    records = []
    for r in df.to_dict("records"):
        rec = {dst: (None if pd.isna(r.get(src)) else r.get(src)) for src, dst in COL_MAP.items()}
        rec["cod_loja"] = int(rec["cod_loja"])
        p = num(rec.get("previstas"), float)
        rec["previstas"] = int(p) if p is not None else 0
        t = num(rec.get("treinadas"), float)
        rec["treinadas"] = int(t) if t is not None else 0
        rec["media"] = num(rec.get("media"), float)
        for k in ("id", "loja", "cidade", "uf", "rede", "bandeira", "canal", "endereco",
                  "titulo", "categorias", "status", "instrutor", "formato", "realizado",
                  "observacoes", "programado"):
            v = rec.get(k)
            rec[k] = None if v is None else str(v)
        rec["data"] = data_str(r.get("_data"))
        records.append(rec)

    by_store = {}
    for rec in records:
        cod = rec["cod_loja"]
        prev = by_store.get(cod)
        if prev is None or (rec["data"] or "") >= (prev["data"] or ""):
            by_store[cod] = rec
    stores = [
        {k: r[k] for k in ("cod_loja", "loja", "cidade", "uf", "rede", "bandeira", "canal", "endereco")}
        for r in by_store.values()
    ]
    stores.sort(key=lambda x: x["cod_loja"])
    log(f"records: planilha={total_raw} sem_cod_loja={dropped} finais={len(records)} | stores unicos={len(stores)}")
    return records, stores


def update_visits(records, stores, roteiro_extra_stores):
    records_json = json.dumps(records, ensure_ascii=False, separators=(",", ":"))
    stores_json = json.dumps(stores, ensure_ascii=False, separators=(",", ":"))

    stores_pub = stores + roteiro_extra_stores
    stores_pub.sort(key=lambda x: x["cod_loja"])
    stores_pub_json = json.dumps(stores_pub, ensure_ascii=False, separators=(",", ":"))

    def rewrite(path, var_kw, stores_j):
        content = path.read_text(encoding="utf-8")
        i0 = content.index(ANCHOR_TOAST) + len(ANCHOR_TOAST)
        i1 = content.index(ANCHOR_ATIVACOES_START)
        block = f"{var_kw} records={records_json};\n{var_kw} stores={stores_j};\n"
        path.write_text(content[:i0] + block + content[i1:], encoding="utf-8")

    rewrite(SRC_HTML, "const", stores_json)
    rewrite(PUB_HTML, "let", stores_pub_json)
    log(f"OK visitas -> {SRC_HTML.name} ({len(stores)} lojas) e {PUB_HTML.name} ({len(stores_pub)} lojas)")


# ============================================================
# 2) ativacoes
# ============================================================
def build_store_index(stores):
    idx = {}
    for st in stores:
        idx.setdefault(norm(st["loja"]), []).append(st["cod_loja"])
    return idx


def fix_foto(v):
    st = s_(v)
    return "https://" + st[7:] if st.lower().startswith("http://") else st


def first_nonempty(row, cols):
    for c in cols:
        if c in row and s_(row[c]):
            return s_(row[c])
    return ""


def process_sheet(path, sheet, status_col):
    df = pd.read_excel(path, sheet_name=sheet)
    df = df[df[status_col].notna()].copy()
    df["_data"] = pd.to_datetime(df["DATA_VISITA"], errors="coerce")
    return df.sort_values("_data", ascending=False)


def build_ativacoes(stores):
    store_idx = build_store_index(stores)
    ativacoes = {}
    unmatched = {"rack": set(), "exposicao": set(), "testeira": set()}

    rack_foto_cols = ["Foto do Rack Finalizado:", "Foto do Rack", "Foto do Rack no Ambiente:",
                       "Foto do Rack Ambiente", "Foto do Local ANTES"]
    codigo_cols = ["Código de Identificação do Rack Implementado:", "Código de Identificação do Rack:"]
    df = process_sheet(RACK_XLSX, "BASE", "Qual a situação atual do Rack?")
    seen, foto_fb = set(), {}
    for _, row in df.iterrows():
        key = norm(row["LOJA"])
        cods = store_idx.get(key)
        foto = fix_foto(first_nonempty(row, rack_foto_cols))
        if not cods:
            unmatched["rack"].add(s_(row["LOJA"])); continue
        for cod in cods:
            if foto and cod not in foto_fb: foto_fb[cod] = foto
        if key in seen: continue
        seen.add(key)
        situacao = s_(row["Qual a situação atual do Rack?"])
        entry = {"situacao": situacao, "presente": situacao == "OK, em perfeito estado.",
                 "tipo": s_(row.get("Qual modelo de Rack está sendo instalado?")),
                 "codigo": first_nonempty(row, codigo_cols), "data": data_str(row["_data"]), "foto": foto}
        for cod in cods:
            ativacoes.setdefault(str(cod), {})["rack"] = dict(entry)
    for cod, foto in foto_fb.items():
        e = ativacoes.get(str(cod), {}).get("rack")
        if e and not e["foto"]: e["foto"] = foto

    exp_foto_cols = ["Foto da Exposição Especial FINALIZADA:", "Foto da Exposição Especial",
                      "Foto da Exposição Especial - Ambiente", "Foto do local ANTES da criação da exposição:",
                      "Envie uma foto do local onde a exposição foi negociada:",
                      "Envie uma foto da avaria/problema:", "Foto do ambiente sem a Exposição Especial"]
    df = process_sheet(EXP_XLSX, "Base", "Qual a situação atual da Exposição Especial?")
    seen, foto_fb = set(), {}
    for _, row in df.iterrows():
        key = norm(row["LOJA"])
        cods = store_idx.get(key)
        foto = fix_foto(first_nonempty(row, exp_foto_cols))
        if not cods:
            unmatched["exposicao"].add(s_(row["LOJA"])); continue
        for cod in cods:
            if foto and cod not in foto_fb: foto_fb[cod] = foto
        if key in seen: continue
        seen.add(key)
        situacao = s_(row["Qual a situação atual da Exposição Especial?"])
        entry = {"situacao": situacao, "presente": situacao != "Foi removida/descontinuada pela loja.",
                 "tipo": s_(row.get("Qual tipo de Exposição Especial foi criada?")),
                 "data": data_str(row["_data"]), "foto": foto}
        for cod in cods:
            ativacoes.setdefault(str(cod), {})["exposicao"] = dict(entry)
    for cod, foto in foto_fb.items():
        e = ativacoes.get(str(cod), {}).get("exposicao")
        if e and not e["foto"]: e["foto"] = foto

    df = process_sheet(TEST_XLSX, "file-04.08.26", "TIPO DE AÇÃO")
    seen, foto_fb = set(), {}
    for _, row in df.iterrows():
        key = norm(row["LOJA"])
        cods = store_idx.get(key)
        foto = fix_foto(s_(row.get("FOTO DA TESTEIRA")))
        if not cods:
            unmatched["testeira"].add(s_(row["LOJA"])); continue
        for cod in cods:
            if foto and cod not in foto_fb: foto_fb[cod] = foto
        if key in seen: continue
        seen.add(key)
        acao = s_(row["TIPO DE AÇÃO"])
        entry = {"acao": acao, "presente": acao.upper().startswith("INSTALA"),
                 "parceiro": s_(row.get("QUAL PARCEIRO VOCÊ POSITIVOU?")),
                 "quantidade": s_(row.get("QUANTIDADE INSTALADA EM LOJA")),
                 "data": data_str(row["_data"]), "foto": foto}
        for cod in cods:
            ativacoes.setdefault(str(cod), {})["testeira"] = dict(entry)
    for cod, foto in foto_fb.items():
        e = ativacoes.get(str(cod), {}).get("testeira")
        if e and not e["foto"]: e["foto"] = foto

    for tipo, names in unmatched.items():
        if names:
            log(f"ativacoes: {len(names)} loja(s) nao casadas em {tipo}")
    return ativacoes


def update_ativacoes(ativacoes):
    js_block = "const ativacoes=" + json.dumps(ativacoes, ensure_ascii=False, separators=(",", ":")) + ";"
    for path in (SRC_HTML, PUB_HTML):
        content = path.read_text(encoding="utf-8")
        content = replace_between(content, ANCHOR_ATIVACOES_START, ANCHOR_ATIVACOES_END, js_block)
        path.write_text(content, encoding="utf-8")
    rack = sum(1 for v in ativacoes.values() if "rack" in v)
    exp = sum(1 for v in ativacoes.values() if "exposicao" in v)
    test = sum(1 for v in ativacoes.values() if "testeira" in v)
    log(f"OK ativacoes -> rack={rack} exposicao={exp} testeira={test}")


# ============================================================
# 3) roteiro
# ============================================================
def find_col(idx, candidates):
    for c in candidates:
        if c in idx:
            return c
    raise KeyError(f"nenhuma das colunas {candidates} encontrada (colunas atuais: {list(idx)})")


def build_roteiro(stores):
    new_codes = {s["cod_loja"] for s in stores}
    wb = openpyxl.load_workbook(ROTEIRO_XLSX, data_only=True, read_only=True)
    ro = wb["Roteiro"]
    at = wb["Atendimentos"]

    ro_rows = list(ro.iter_rows(values_only=True))
    ro_headers = list(ro_rows[0])
    ro_idx = {h: i for i, h in enumerate(ro_headers)}
    ro_rows = ro_rows[1:]
    col_instrutor = find_col(ro_idx, ["INSTRUTOR / REPRESENTANTE", "INSTRUTOR"])

    at_rows = list(at.iter_rows(values_only=True))
    at_headers = list(at_rows[0])
    at_idx = {h: i for i, h in enumerate(at_headers)}
    at_rows = at_rows[1:]
    visited_codes = {str(r[at_idx["COD_LOJA"]]) for r in at_rows if r[at_idx["COD_LOJA"]] is not None}

    # a coluna do instrutor as vezes vem suja (valores da coluna "LOJA VISITADA"
    # vazam pra ca em algumas linhas, tipo "SIM"/"NAO"/vazio) - filtra isso.
    GARBAGE_NOMES = {"SIM", "NAO", "N/A", "", "-"}

    roteiro = {}
    extra_stores = {}
    for r in ro_rows:
        nome = norm(r[ro_idx[col_instrutor]])
        cod_raw = r[ro_idx["COD_LOJA"]]
        if cod_raw is None or not nome or nome in GARBAGE_NOMES:
            continue
        cod = str(int(cod_raw))
        visitada = cod in visited_codes
        d = roteiro.setdefault(nome, {"total": 0, "visitadas": 0, "pendentes": []})
        d["total"] += 1
        if visitada:
            d["visitadas"] += 1
        else:
            d["pendentes"].append({
                "cod_loja": int(cod), "loja": r[ro_idx["NOME LOJA"]], "bandeira": r[ro_idx["BANDEIRA"]],
                "cidade": r[ro_idx["CIDADE"]], "uf": r[ro_idx["UF"]], "frequencia": r[ro_idx["Frequencia"]],
            })
        if int(cod) not in new_codes and int(cod) not in extra_stores:
            extra_stores[int(cod)] = {
                "cod_loja": int(cod), "loja": r[ro_idx["NOME LOJA"]], "cidade": r[ro_idx["CIDADE"]],
                "uf": r[ro_idx["UF"]], "rede": r[ro_idx["BANDEIRA"]], "bandeira": r[ro_idx["BANDEIRA"]],
                "canal": r[ro_idx["CANAL"]], "endereco": r[ro_idx["ENDEREÇO"]],
            }
    log(f"roteiro: {len(roteiro)} colaboradores | {len(extra_stores)} lojas do roteiro sem treinamento")
    return roteiro, list(extra_stores.values())


def update_roteiro(roteiro):
    js_block = "const roteiro=" + json.dumps(roteiro, ensure_ascii=False, separators=(",", ":")) + ";"
    content = PUB_HTML.read_text(encoding="utf-8")
    content = replace_between(content, ANCHOR_ROTEIRO_START, ANCHOR_ROTEIRO_END, js_block)
    PUB_HTML.write_text(content, encoding="utf-8")
    log(f"OK roteiro -> {PUB_HTML.name}")


def update_last_updated_stamp():
    today = datetime.date.today().strftime("%d/%m/%Y")
    pattern = re.compile(r'(id="sideUpdated">Atualizado: )[^<]*(<)')
    for path in (SRC_HTML, PUB_HTML):
        content = path.read_text(encoding="utf-8")
        new_content = pattern.sub(lambda m: f"{m.group(1)}{today}{m.group(2)}", content, count=1)
        if new_content != content:
            path.write_text(new_content, encoding="utf-8")
    log(f"OK carimbo de atualizacao -> {today}")


# ============================================================
# validacao + deploy
# ============================================================
def validate(prev_counts):
    ok = True
    for path in (SRC_HTML, PUB_HTML):
        content = path.read_text(encoding="utf-8")
        try:
            records = extract_json_block(content, r"(?:const|let) records=(\[.*?\]);\s*\n(?:const|let) stores=")
            stores = extract_json_block(content, r"(?:const|let) stores=(\[.*?\]);\n")
            ativ = extract_json_block(content, re.escape(ANCHOR_ATIVACOES_START) + r"\nconst ativacoes=(\{.*?\});\n")
        except Exception as e:
            log(f"VALIDACAO FALHOU em {path.name}: {e}")
            return False
        prev = prev_counts.get(path.name, {})
        if prev.get("records") and len(records) < prev["records"] * 0.5:
            log(f"VALIDACAO SUSPEITA em {path.name}: records caiu de {prev['records']} para {len(records)}")
            ok = False
        log(f"validado {path.name}: records={len(records)} stores={len(stores)} ativacoes={len(ativ)}")
    if PUB_HTML.name:
        content = PUB_HTML.read_text(encoding="utf-8")
        try:
            extract_json_block(content, re.escape(ANCHOR_ROTEIRO_START) + r"\nconst roteiro=(\{.*?\});\n")
        except Exception as e:
            log(f"VALIDACAO FALHOU (roteiro) em {PUB_HTML.name}: {e}")
            ok = False
    return ok


def previous_counts():
    counts = {}
    for path in (SRC_HTML, PUB_HTML):
        try:
            content = path.read_text(encoding="utf-8")
            records = extract_json_block(content, r"(?:const|let) records=(\[.*?\]);\s*\n(?:const|let) stores=")
            counts[path.name] = {"records": len(records)}
        except Exception:
            counts[path.name] = {}
    return counts


FIREBASE_CMD = r"C:\Users\usuario\AppData\Local\Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.19.0-win-x64\firebase.cmd"


def deploy():
    log("iniciando firebase deploy...")
    r = subprocess.run(
        [FIREBASE_CMD, "deploy", "--only", "hosting", "--project", "historico-loja-roku"],
        cwd=str(ROOT), capture_output=True, text=True, shell=True, timeout=600,
    )
    log(r.stdout[-3000:])
    if r.returncode != 0:
        log("DEPLOY FALHOU: " + r.stderr[-2000:])
        return False
    log("deploy concluido com sucesso")
    return True


def main():
    LOG_DIR.mkdir(exist_ok=True)
    log("=== rotina de atualizacao iniciada ===")
    prev = previous_counts()
    try:
        records, stores = build_records_stores()
        ativacoes = build_ativacoes(stores)
    except Exception as e:
        log(f"ERRO durante atualizacao (visitas/ativacoes): {e!r}")
        write_log()
        sys.exit(1)

    # roteiro vem de um arquivo que o usuario edita/atualiza por conta propria
    # (planilha com query table ao vivo) - o esquema de colunas dele ja mudou
    # sozinho antes, entao essa etapa e "best effort": se falhar, mantem o
    # roteiro publicado anteriormente e segue com o resto normalmente.
    try:
        roteiro, roteiro_extra_stores = build_roteiro(stores)
    except Exception as e:
        log(f"AVISO: roteiro nao atualizado ({e!r}) - mantendo o roteiro publicado anteriormente.")
        roteiro, roteiro_extra_stores = None, []

    try:
        update_visits(records, stores, roteiro_extra_stores)
        update_ativacoes(ativacoes)
        if roteiro is not None:
            update_roteiro(roteiro)
        update_last_updated_stamp()
    except Exception as e:
        log(f"ERRO durante atualizacao (gravacao dos arquivos): {e!r}")
        write_log()
        sys.exit(1)

    if not validate(prev):
        log("VALIDACAO FALHOU - deploy CANCELADO, arquivos locais podem estar incorretos.")
        write_log()
        sys.exit(1)

    if "--dry-run" in sys.argv:
        log("--dry-run: pulando o firebase deploy.")
        write_log()
        sys.exit(0)

    ok = deploy()
    write_log()
    sys.exit(0 if ok else 1)


def write_log():
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    (LOG_DIR / f"{ts}.log").write_text("\n".join(log_lines), encoding="utf-8")


if __name__ == "__main__":
    main()
