# -*- coding: utf-8 -*-
"""Injeta o bloco `const roteiro={...}` (cobertura de roteiro por colaborador)
e as lojas do roteiro que ainda nao existiam em `stores`, no public/index.html
(unico arquivo com sistema de identidade)."""
import json
from pathlib import Path

ROOT = Path(r"C:\dev\historico loja")
HTML = ROOT / "public" / "index.html"
DATA = ROOT / "tools" / "_roteiro_data.json"

with open(DATA, encoding="utf-8") as f:
    data = json.load(f)
roteiro = data["roteiro"]
missing_stores = data["missing_stores"]

content = HTML.read_text(encoding="utf-8")

# ---------- 1) injeta const roteiro={...} logo apos o bloco de ativacoes ----------
start, end = "/*__ROTEIRO_START__*/", "/*__ROTEIRO_END__*/"
roteiro_json = json.dumps(roteiro, ensure_ascii=False, separators=(",", ":"))
js_block = f"const roteiro={roteiro_json};"
wrapped = f"{start}\n{js_block}\n{end}"

import re
if start in content:
    pattern = re.compile(re.escape(start) + r".*?" + re.escape(end), re.DOTALL)
    content = pattern.sub(lambda m: wrapped, content, count=1)
else:
    anchor = "const byStore=new Map();"
    assert anchor in content, "ponto de insercao do roteiro nao encontrado"
    content = content.replace(anchor, f"{wrapped}\n{anchor}", 1)

# ---------- 2) acrescenta as lojas do roteiro que faltam em `stores` ----------
lines = content.split("\n")
for i, line in enumerate(lines):
    if line.startswith("let stores=[") or line.startswith("const stores=["):
        assert line.rstrip().endswith("];"), "formato inesperado da linha stores"
        extra = "".join(json.dumps(s, ensure_ascii=False, separators=(",", ":")) + "," for s in missing_stores)
        lines[i] = line[:-2] + "," + extra.rstrip(",") + "];" if extra else line
        print("lojas adicionadas a stores:", len(missing_stores))
        break
else:
    raise RuntimeError("linha const stores= nao encontrada")
content = "\n".join(lines)

HTML.write_text(content, encoding="utf-8")
print("OK ->", HTML)
