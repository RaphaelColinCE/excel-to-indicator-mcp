"""Service de parsing de formules Excel"""

import re
import logging
from typing import List, Tuple

from src.models.excel_models import (
    ExcelFormula,
    ExcelReference,
    FormulaToken,
    ReferenceType,
)

logger = logging.getLogger("mcp_server")

# Pattern combiné pour les références inter-feuilles (gère range et cellule)
# Sheet peut être un nom simple (ex: Sheet1) ou un nom entre guillemets simples
_RE_CROSS_SHEET = re.compile(
    r"(?:'(?P<sheet_q>[^']+)'|(?P<sheet_u>[A-Za-z0-9_\-.]+))"
    r"!(?P<start>\$?[A-Z]+\$?\d+)(?::(?P<end>\$?[A-Z]+\$?\d+))?",
    re.IGNORECASE,
)

# Ranges locaux : A1:B10, $A$1:$B$10, etc. (pas précédés de !)
_RE_LOCAL_RANGE = re.compile(
    r"(?<!!)\b(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)\b", re.IGNORECASE
)

# Cellules locales isolées
_RE_LOCAL_CELL = re.compile(r"(?<![!\w])(\$?[A-Z]+\$?\d+)(?![:A-Z\d])", re.IGNORECASE)

# Fonctions Excel : nom suivi d'une parenthèse ouvrante
_RE_FUNCTION = re.compile(r"\b([A-Z][A-Z0-9_]*)\s*\(", re.IGNORECASE)

_RE_STRING = re.compile(r'"[^"]*"')
_RE_NUMBER = re.compile(r"(?<![A-Z\$])\d+(\.\d+)?(?!\d)")


def parse_formula(
    formula: str,
    sheet: str,
    cell: str,
    file_path: str = "",
) -> ExcelFormula:
    """Parser une formule Excel et retourner un objet ExcelFormula structuré.

    Args:
        formula: Formule Excel brute (ex: =SUM(Sheet1!A1:A10))
        sheet: Feuille contenant la formule
        cell: Cellule contenant la formule
        file_path: Chemin du fichier Excel (optionnel)

    Returns:
        ExcelFormula avec tokens, références et fonctions extraits
    """
    references: List[ExcelReference] = []
    tokens: List[FormulaToken] = []
    functions_used: List[str] = []
    seen_refs: set = set()

    # Intervalles à exclure des recherches locales (déjà pris par cross-sheet)
    cross_sheet_spans: List[Tuple[int, int]] = []

    # ── Références inter-feuilles ────────────────────────────────────────────
    for m in _RE_CROSS_SHEET.finditer(formula):
        sheet_name = m.group("sheet_q") or m.group("sheet_u")
        start_cell = m.group("start")
        end_cell = m.group("end")
        ref_value = m.group(0)

        cross_sheet_spans.append((m.start(), m.end()))

        if ref_value in seen_refs:
            continue
        seen_refs.add(ref_value)

        ref_type = ReferenceType.EXTERNAL_RANGE if end_cell else ReferenceType.EXTERNAL_CELL
        references.append(
            ExcelReference(
                type=ref_type,
                value=ref_value,
                sheet=sheet_name,
                start_cell=start_cell,
                end_cell=end_cell,
            )
        )
        tokens.append(FormulaToken(type="reference", value=ref_value, position=m.start()))

    # ── Ranges locaux ────────────────────────────────────────────────────────
    for m in _RE_LOCAL_RANGE.finditer(formula):
        if _in_spans(m.start(), cross_sheet_spans):
            continue
        ref_value = m.group(1)
        if ref_value in seen_refs:
            continue
        seen_refs.add(ref_value)
        parts = ref_value.split(":")
        references.append(
            ExcelReference(
                type=ReferenceType.RANGE,
                value=ref_value,
                start_cell=parts[0],
                end_cell=parts[1] if len(parts) > 1 else None,
            )
        )
        tokens.append(FormulaToken(type="reference", value=ref_value, position=m.start()))

    # ── Fonctions ────────────────────────────────────────────────────────────
    for m in _RE_FUNCTION.finditer(formula):
        func_name = m.group(1).upper()
        if func_name not in functions_used:
            functions_used.append(func_name)
        tokens.append(FormulaToken(type="function", value=func_name, position=m.start()))

    # ── Constantes numériques ─────────────────────────────────────────────────
    for m in _RE_NUMBER.finditer(formula):
        tokens.append(FormulaToken(type="constant", value=m.group(0), position=m.start()))

    # ── Chaînes de caractères ─────────────────────────────────────────────────
    for m in _RE_STRING.finditer(formula):
        tokens.append(FormulaToken(type="constant", value=m.group(0), position=m.start()))

    tokens.sort(key=lambda t: t.position)

    has_cross_sheet = any(
        r.type in (ReferenceType.EXTERNAL_CELL, ReferenceType.EXTERNAL_RANGE)
        for r in references
    )

    return ExcelFormula(
        original_formula=formula,
        file_path=file_path,
        sheet=sheet,
        cell=cell,
        tokens=tokens,
        references=references,
        functions_used=functions_used,
        is_complex=len(formula) > 100 or len(references) > 5,
        has_cross_sheet_refs=has_cross_sheet,
        has_external_refs=False,
        depth_estimate=_estimate_depth(formula),
    )


def _in_spans(pos: int, spans: List[Tuple[int, int]]) -> bool:
    """Vérifier si une position est dans l'un des intervalles donnés."""
    return any(start <= pos < end for start, end in spans)


def _estimate_depth(formula: str) -> int:
    """Estimer la profondeur d'imbrication d'une formule (parenthèses)."""
    max_depth = 0
    current = 0
    for ch in formula:
        if ch == "(":
            current += 1
            max_depth = max(max_depth, current)
        elif ch == ")":
            current = max(0, current - 1)
    return max_depth


def extract_references_from_formula(formula: str) -> List[Tuple[str, str]]:
    """Extraire les paires (feuille, cellule/range) d'une formule.

    Utilitaire léger pour les besoins de résolution de dépendances.

    Returns:
        Liste de tuples (sheet_or_empty, reference)
    """
    result: List[Tuple[str, str]] = []
    seen: set = set()
    cross_sheet_spans: List[Tuple[int, int]] = []

    for m in _RE_CROSS_SHEET.finditer(formula):
        key = m.group(0)
        cross_sheet_spans.append((m.start(), m.end()))
        if key not in seen:
            seen.add(key)
            sheet_name = m.group("sheet_q") or m.group("sheet_u")
            start, end = m.group("start"), m.group("end")
            ref = f"{start}:{end}" if end else start
            result.append((sheet_name, ref))

    for m in _RE_LOCAL_RANGE.finditer(formula):
        if _in_spans(m.start(), cross_sheet_spans):
            continue
        key = m.group(1)
        if key not in seen:
            seen.add(key)
            result.append(("", key))

    return result

