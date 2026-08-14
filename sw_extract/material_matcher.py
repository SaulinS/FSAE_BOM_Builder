"""
sw_extract/material_matcher.py

"Corretor" de nomenclatura: casa o nome de material nativo do SolidWorks
(ex: "AISI 1020") contra os nomes exatos do catalogo Materials do Cost
Report (ex: "Steel, Mild (per kg)"), usando fuzzy matching.

Usa difflib (biblioteca padrao do Python, sem dependencia externa) para nao
exigir instalacao extra na maquina de quem modela. Se quiserem resultados
melhores no futuro, a biblioteca `rapidfuzz` (pip install rapidfuzz) da
matches mais precisos com pouca mudanca de codigo.

IMPORTANTE: por design, este modulo so SUGERE candidatos -- nunca aplica
uma correcao sozinho. Decisao final e sempre de uma pessoa (ver GUI de
padronizacao/corretor combinada com o time).
"""

import difflib


def suggest_material_matches(raw_name, catalog_names, max_suggestions=3, min_score=0.35):
    """
    Retorna ate `max_suggestions` nomes do catalogo mais parecidos com
    `raw_name`, ordenados por score decrescente (0 a 1), cada um com seu
    score. Ex: [("Steel, Mild (per kg)", 0.62), ...]

    Nao aplica nada -- so retorna sugestoes para confirmacao manual.
    """
    if not raw_name:
        return []

    raw_lower = raw_name.lower()
    scored = []
    for name in catalog_names:
        score = difflib.SequenceMatcher(None, raw_lower, name.lower()).ratio()
        if score >= min_score:
            scored.append((name, round(score, 3)))

    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[:max_suggestions]


def load_catalog_material_names(wb):
    """Le todos os nomes de material da aba Materials (coluna C)."""
    ws = wb["Materials"]
    names = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=3):
        v = row[2].value
        if v:
            names.append(v)
    return names
