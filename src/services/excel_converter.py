"""Conversion d'une feuille Excel entière en formules d'indicateurs.

Logique :
  1. Lire la feuille et extraire la carte des en-têtes (header_row).
  2. Extraire la carte des références absolues via les lignes #CELL du fichier.
  3. Pour chaque cellule de formula_row dont l'en-tête N'EST PAS dans le
     mapping (= indicateur de sortie nouveau), convertir la formule Excel
     en expression fact[IndicatorName=...; SRF_LEI=?; endDate=?].value.
  4. Retourner les enregistrements (source_excel_sheet, source_cell,
     source_formula, indicator_name, indicator_formula).
"""

from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Optional

import openpyxl

logger = logging.getLogger("mcp_server")

# Regex : référence de cellule Excel  ($A$1 | A$1 | $A1 | A1)
# Negative lookbehind pour éviter de matcher l'intérieur des identifiants
_CELL_REF_RE = re.compile(r"(?<![A-Za-z_])\$?([A-Z]{1,3})\$?(\d{1,7})", re.IGNORECASE)

# Regex : plage de cellules Excel  (COL1:COL2, $COL$ROW1:$COL$ROW2, mixtes)
# Utilisé pour détecter les agrégations sur toutes les lignes dynamiques (institution rows).
_RANGE_REF_RE = re.compile(
    r"(\$?[A-Z]{1,3})\$?(\d{1,7}):(\$?[A-Z]{1,3})\$?(\d{1,7})",
    re.IGNORECASE,
)

# Extraction de l'indicateur dans une cellule #CELL de configuration
_MC_COMP_RE = re.compile(r"(?:OUTPUT|INPUT)MCCOMP:\{[^}]*\}([^;\n]+)")

# Formules à ignorer pour la génération d'indicateurs intermédiaires.
# Ces fonctions retournent des métadonnées de contexte (entité, pays) inutiles
# dans les formules d'indicateurs finales.
_IGNORED_INTERMEDIATE_FORMULAS: frozenset[str] = frozenset([
    "SELCOMPANY()",
    "@SELCOMPANY()",
    'DOCUMENTVARIABLE_LABEL_UNIQUE("CC")',
    '@DOCUMENTVARIABLE_LABEL_UNIQUE("CC")',
])

# Post-processing : SUM(fact[...].value ...) → contenu sans wrapper SUM.
# Gère également les multi-plages séparées par ";" :
#   SUM(fact[A].value + fact[B].value; fact[C].value + fact[D].value)
#   → fact[A].value + fact[B].value + fact[C].value + fact[D].value
_SUM_FACT_RE = re.compile(
    r'SUM\((fact\[[^\]]+\]\.value(?:\s*\+\s*fact\[[^\]]+\]\.value)*'
    r'(?:\s*;\s*fact\[[^\]]+\]\.value(?:\s*\+\s*fact\[[^\]]+\]\.value)*)*)\)'
)

# Post-processing : COUNTIF résolus (plage → fact) en LIST_COUNT.
# Trois formes de critère :
#   1. String simple    : "LPS"   → value="LPS"
#   2. Comparaison inline: ">=0"  → value>=0
#   3. Comparaison concat: ">"&0  → value>0
_COUNTIF_STR_RE = re.compile(
    r'COUNTIF\(fact\[([^\]]+?)\]\.value;"([^"<>=!]*)"\)'
)
_COUNTIF_CMP_INLINE_RE = re.compile(
    r'COUNTIF\(fact\[([^\]]+?)\]\.value;"(>=|<=|<>|>|<)([^"]*)"\)'
)
_COUNTIF_CMP_CONCAT_RE = re.compile(
    r'COUNTIF\(fact\[([^\]]+?)\]\.value;"(>=|<=|<>|>|<)"&([^;)]+)\)'
)


class ExcelConverterService:
    """Convertit les formules d'une feuille Excel en formules d'indicateurs.

    Le mapping (mapping.json) sert de dictionnaire :
      - code_colonne → indicator_name + flag static
      - Les colonnes dont le code N'est PAS dans le mapping sont des
        indicateurs de sortie à définir (formulas converties → output).
    """

    def __init__(self, mapping_path: Path) -> None:
        with open(mapping_path, encoding="utf-8") as fh:
            data = json.load(fh)

        self.mapping: dict[str, dict] = data.get("mapping", {})

        # Index inverse : indicator_name → static flag (pour les entrées de cell_ref_map)
        self._indicator_static: dict[str, bool] = {}
        for entry in self.mapping.values():
            self._indicator_static[entry["indicator_name"]] = entry.get("static", False)

    # ─────────────────────────────────────────────────────────────────────────
    # API publique
    # ─────────────────────────────────────────────────────────────────────────

    def convert_sheet(
        self,
        file_path: Path,
        sheet_name: str,
        header_row: int = 25,
        formula_row: int = 26,
        output_columns: Optional[list[str]] = None,
    ) -> list[dict]:
        """Convertir toutes les formules d'une feuille Excel.

        Args:
            file_path: Chemin vers le fichier Excel.
            sheet_name: Nom de la feuille à traiter.
            header_row: Numéro de la ligne contenant les codes des colonnes.
            formula_row: Numéro de la ligne contenant les formules à convertir.
            output_columns: Lettres de colonnes à forcer comme sorties (optionnel).
                            Si None, on inclut les colonnes dont l'en-tête n'est
                            pas dans le mapping (= nouvel indicateur).

        Returns:
            Liste de dicts avec les clés :
              source_excel_sheet, source_cell, source_formula,
              indicator_name, indicator_formula.
        """
        wb = openpyxl.load_workbook(file_path, keep_links=False)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Feuille '{sheet_name}' introuvable dans {file_path.name}")
        ws = wb[sheet_name]

        # 1. Carte des en-têtes : lettre_colonne → code_colonne
        col_header: dict[str, str] = {}
        for cell in ws[header_row]:
            if cell.value is not None:
                col_header[cell.column_letter] = str(cell.value).strip()

        # 2. Carte des références absolues : "$COL$ROW" → indicator_name
        cell_ref_map = self._build_cell_ref_map(ws)

        # 3. Identifier les colonnes de sortie (nouveaux indicateurs)
        results: list[dict] = []

        # Pre-scan : générer les indicateurs intermédiaires pour les cellules
        # de formula_row qui ont une formule mais pas d'en-tête de colonne connu.
        # Ces cellules sont dynamiques (formula_row = ligne dynamique principale).
        _used_names: set[str] = set(cell_ref_map.values())
        _intermediate_keys: set[str] = set()
        for _cell in ws[formula_row]:
            if not _cell.value or not str(_cell.value).startswith("="):
                continue
            # Ignorer les formules de métadonnées (entité, pays)
            _formula_body = str(_cell.value)[1:].strip()
            if _formula_body in _IGNORED_INTERMEDIATE_FORMULAS:
                continue
            _col = _cell.column_letter
            if col_header.get(_col):
                continue  # Colonne avec indicateur connu, pas intermédiaire
            _abs_key = f"${_col}${formula_row}"
            if _abs_key in cell_ref_map:
                continue
            _label = self._find_cell_label_in_ws(ws, formula_row, _cell.column)
            _short = self._make_label_short(_label, 8) if _label else _col.upper()
            _base = f"{sheet_name}_{_short}_$"
            _name = _base
            _ctr = 2
            while _name in _used_names:
                _name = f"{_base}{_ctr}"
                _ctr += 1
            _used_names.add(_name)
            cell_ref_map[_abs_key] = _name
            cell_ref_map[f"{_col}{formula_row}"] = _name
            _intermediate_keys.add(_abs_key)

        for cell in ws[formula_row]:
            raw = cell.value
            if not raw or not str(raw).startswith("="):
                continue

            col = cell.column_letter
            abs_key = f"${col}${formula_row}"
            header_code = col_header.get(col)

            if header_code:
                indicator_name = header_code
                is_intermediate = False
                # Filtre : est-ce un indicateur de sortie ?
                if output_columns is not None:
                    if col not in output_columns:
                        continue
                else:
                    if header_code in self.mapping:
                        continue
            elif abs_key in _intermediate_keys:
                indicator_name = cell_ref_map[abs_key]
                is_intermediate = True
            else:
                continue  # Pas d'en-tête et pas intermédiaire connu

            formula_text = str(raw)[1:]  # retirer le "="
            converted = self._convert_formula(
                formula_text, col_header, cell_ref_map, formula_row
            )

            results.append(
                {
                    "source_excel_sheet": sheet_name,
                    "source_cell": f"{col}{formula_row}",
                    "source_formula": formula_text,
                    "indicator_name": indicator_name,
                    "indicator_formula": converted,
                    "intermediate": is_intermediate,
                }
            )

        return results

    def convert_multi_row_sheet(
        self,
        file_path: Path,
        sheet_name: str,
    ) -> list[dict]:
        """Convertir une feuille Excel à structure multi-lignes (DATA/CELL alternées).

        Utilisée pour des feuilles comme S1CRUF où chaque ligne #DATA est suivie
        d'une ligne #CELL définissant les indicateurs colonne par colonne.

        Logique :
          1. Construire cell_ref_map globale depuis TOUTES les lignes #CELL.
          2. Pour chaque cellule formule `$COL$ROW` :
             - L'indicator_name est cell_ref_map[$COL$ROW].
             - Si l'indicateur n'est pas dans le mapping (nouvel indicateur), convertir.
          3. Résoudre les références internes via cell_ref_map (multi-lignes).
          4. Les références non résolues (ex: plages sans indicateur) restent telles quelles.
        """
        wb = openpyxl.load_workbook(file_path, keep_links=False)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Feuille '{sheet_name}' introuvable dans {file_path.name}")
        ws = wb[sheet_name]

        cell_ref_map = self._build_cell_ref_map(ws)

        # Détecter la ligne institution (ligne #DATA avec les lignes dynamiques par LEI).
        # Priorité 1 : tag explicite "#DATA_DYNAMIQUE" en colonne A (déterministe, recommandé).
        # Priorité 2 (fallback) : heuristique — la ligne #DATA avec le plus d'entrées dans
        # cell_ref_map (rétrocompatibilité avec les anciens fichiers non taggués).
        institution_row: Optional[int] = None
        for row_num in range(1, ws.max_row + 1):
            a_val = ws.cell(row=row_num, column=1).value
            if isinstance(a_val, str) and a_val.strip().upper().startswith("#DATA_DYNAMIQUE"):
                institution_row = row_num
                break

        if institution_row is None:
            row_counts: dict[int, int] = {}
            for key in cell_ref_map:
                if key.startswith("$") and key.count("$") == 2:
                    try:
                        row_num_key = int(key.split("$")[2])
                        row_counts[row_num_key] = row_counts.get(row_num_key, 0) + 1
                    except (ValueError, IndexError):
                        pass
            institution_row = max(row_counts, key=row_counts.get) if row_counts else None

        results: list[dict] = []
        max_col = ws.max_column

        # Pre-scan : générer les indicateurs intermédiaires (cellules avec formule
        # mais sans entrée dans cell_ref_map) AVANT la boucle principale,
        # afin que les formules qui les référencent puissent les résoudre.
        intermediate_keys = self._prescan_intermediate_cells(
            ws, sheet_name, cell_ref_map, institution_row
        )

        for row_num in range(1, ws.max_row + 1):
            a_val = ws.cell(row=row_num, column=1).value
            if not isinstance(a_val, str) or not a_val.startswith("#DATA"):
                continue

            # col_header pour ce DATA row : extraire depuis la ligne #CELL suivante
            row_col_header: dict[str, str] = {}
            next_cell_row = row_num + 1
            if ws.cell(row=next_cell_row, column=1).value == "#CELL":
                for c in range(1, max_col + 1):
                    v = ws.cell(row=next_cell_row, column=c).value
                    if v and isinstance(v, str):
                        m = _MC_COMP_RE.search(v)
                        if m:
                            row_col_header[ws.cell(row=next_cell_row, column=c).column_letter] = (
                                m.group(1).strip().rstrip(";")
                            )

            for col_idx in range(1, max_col + 1):
                cell = ws.cell(row=row_num, column=col_idx)
                raw = cell.value
                if not raw or not str(raw).startswith("="):
                    continue

                col = cell.column_letter
                abs_key = f"${col}${row_num}"
                indicator_name = cell_ref_map.get(abs_key)
                if not indicator_name:
                    continue

                is_intermediate = abs_key in intermediate_keys

                # Pour les indicateurs connus du mapping : skip sauf si intermédiaire
                if indicator_name in self.mapping and not is_intermediate:
                    continue

                formula_text = str(raw)[1:]
                converted = self._convert_formula(
                    formula_text, row_col_header, cell_ref_map, row_num,
                    institution_row=institution_row,
                )
                results.append(
                    {
                        "source_excel_sheet": sheet_name,
                        "source_cell": f"{col}{row_num}",
                        "source_formula": formula_text,
                        "indicator_name": indicator_name,
                        "indicator_formula": converted,
                        "intermediate": is_intermediate,
                    }
                )

                # Auto-registration : s'assurer que les clés absolues et relatives
                # sont toutes deux enregistrées pour les formules suivantes.
                cell_ref_map[abs_key] = indicator_name
                cell_ref_map[f"{col}{row_num}"] = indicator_name

        return results

    # ─────────────────────────────────────────────────────────────────────────
    # Construction de la carte des références absolues
    # ─────────────────────────────────────────────────────────────────────────

    def _build_cell_ref_map(self, ws) -> dict[str, str]:
        """Construire la carte {$COL$ROW → indicator_name} depuis les lignes #CELL.

        Dans le format ISS Excel :
          - Ligne N-1 : codes indicateurs directement (institution-style, row before #DATA)
          - Ligne N   : #DATA_x  (données / formules)
          - Ligne N+1 : #CELL    (métadonnées OUTPUTMCCOMP / INPUTMCCOMP)
        La ligne #CELL définit les cellules de la ligne #DATA qui la précède.
        Pour les colonnes non couvertes par #CELL, on regarde la ligne N-1 pour
        un code indicateur direct (pattern : lettres majuscules + chiffres + _).
        """
        _IND_CODE_RE = re.compile(r"^[A-Z][A-Z0-9_]+$")

        # Repérer toutes les lignes #CELL (colonne A = "#CELL")
        cell_config_rows: set[int] = set()
        for (row_cells,) in ws.iter_rows(min_col=1, max_col=1):
            if row_cells.value == "#CELL":
                cell_config_rows.add(row_cells.row)

        ref_map: dict[str, str] = {}
        for cell_row_num in sorted(cell_config_rows):
            data_row = cell_row_num - 1  # La ligne données est juste avant

            # Passe 0 : pattern "formule à gauche / code à droite" dans data_row.
            # Ex : M21=<formule>, N21='S1CRU2_TALEC' → $M$21 = S1CRU2_TALEC.
            for cell in ws[data_row]:
                if not cell.value or not str(cell.value).startswith("="):
                    continue
                right = ws.cell(row=data_row, column=cell.column + 1)
                if right.value and isinstance(right.value, str):
                    v = right.value.strip()
                    if _IND_CODE_RE.match(v):
                        key = f"${cell.column_letter}${data_row}"
                        if key not in ref_map:
                            ref_map[key] = v
                            ref_map[f"{cell.column_letter}{data_row}"] = v

            # Passe 1 : lire la ligne d'en-tête au-dessus du DATA row (institution-style)
            # N'ajoute que pour les colonnes qui ont des formules sur data_row
            header_row = data_row - 1
            for cell in ws[header_row]:
                v = cell.value
                if not v or not isinstance(v, str):
                    continue
                v = v.strip()
                if not _IND_CODE_RE.match(v):
                    continue  # Pas un code indicateur (description avec espaces, etc.)
                col = cell.column_letter
                # Vérifier que data_row a bien une formule ici (sinon faux positif)
                data_cell = ws.cell(row=data_row, column=cell.column)
                if data_cell.value is None:
                    continue
                key = f"${col}${data_row}"
                if key not in ref_map:  # Ne pas écraser les entrées #CELL
                    ref_map[key] = v
                    ref_map[f"{col}{data_row}"] = v

            # Passe 2 : lire le #CELL row (OUTPUTMCCOMP / INPUTMCCOMP) — prioritaire
            for cell in ws[cell_row_num]:
                if not cell.value or not isinstance(cell.value, str):
                    continue
                m = _MC_COMP_RE.search(cell.value)
                if not m:
                    continue
                indicator = m.group(1).strip().rstrip(";")
                if not indicator:
                    continue
                col = cell.column_letter
                ref_map[f"${col}${data_row}"] = indicator  # Écrase si déjà présent
                ref_map[f"{col}{data_row}"] = indicator

        return ref_map

    # ─────────────────────────────────────────────────────────────────────────
    # Indicateurs intermédiaires (cellules sans header = calcul interne)
    # ─────────────────────────────────────────────────────────────────────────

    @staticmethod
    def _make_label_short(label: str, max_len: int = 8) -> str:
        """Construire un code court à partir d'un label (alphanum, uppercase)."""
        clean = re.sub(r"[^A-Za-z0-9]", "", label)
        result = clean[:max_len].upper()
        return result if result else "X"

    @staticmethod
    def _find_cell_label_in_ws(ws, row_num: int, col_idx: int) -> Optional[str]:
        """Remonter les lignes pour trouver un label textuel pour une colonne donnée."""
        for dr in range(1, 12):
            r = row_num - dr
            if r < 1:
                break
            v = ws.cell(row=r, column=col_idx).value
            if v is None:
                continue
            sv = str(v).strip()
            if not sv or sv.startswith("#") or sv.startswith("="):
                continue
            # Ignorer les codes indicateurs purs (ex: S1CRU2_TARG1)
            if re.match(r"^[A-Z][A-Z0-9_]+$", sv):
                continue
            return sv
        return None

    def _prescan_intermediate_cells(
        self,
        ws,
        sheet_name: str,
        cell_ref_map: dict[str, str],
        institution_row: Optional[int],
    ) -> set[str]:
        """Identifier les cellules intermédiaires et les ajouter à cell_ref_map.

        Une cellule est intermédiaire si elle contient une formule mais n'a pas
        d'indicateur dans cell_ref_map (pas de OUTPUTMCCOMP / header connu).
        Nom généré : {sheet_name}_{LabelShort}[_$] (avec _$ si ligne dynamique).
        Retourne l'ensemble des clés absolues ($COL$ROW) des cellules intermédiaires.
        """
        intermediate_keys: set[str] = set()
        used_names: set[str] = set(cell_ref_map.values())
        max_col = ws.max_column

        for row_num in range(1, ws.max_row + 1):
            a_val = ws.cell(row=row_num, column=1).value
            if not isinstance(a_val, str) or not a_val.startswith("#DATA"):
                continue
            is_dynamic = row_num == institution_row
            for col_idx in range(1, max_col + 1):
                raw = ws.cell(row=row_num, column=col_idx).value
                if not raw or not str(raw).startswith("="):
                    continue
                # Ignorer les formules de métadonnées (entité, pays)
                formula_body = str(raw)[1:].strip()
                if formula_body in _IGNORED_INTERMEDIATE_FORMULAS:
                    continue
                col = ws.cell(row=row_num, column=col_idx).column_letter
                abs_key = f"${col}${row_num}"
                if abs_key in cell_ref_map:
                    continue  # Déjà connu, pas intermédiaire
                label = self._find_cell_label_in_ws(ws, row_num, col_idx)
                short = self._make_label_short(label, 8) if label else col.upper()
                suffix = "_$" if is_dynamic else ""
                base_name = f"{sheet_name}_{short}{suffix}"
                # Éviter les doublons
                name = base_name
                ctr = 2
                while name in used_names:
                    name = f"{base_name}{ctr}"
                    ctr += 1
                used_names.add(name)
                cell_ref_map[abs_key] = name
                cell_ref_map[f"{col}{row_num}"] = name
                intermediate_keys.add(abs_key)

        return intermediate_keys

    # ─────────────────────────────────────────────────────────────────────────
    # Simplification SELFINYEAR
    # ─────────────────────────────────────────────────────────────────────────

    @staticmethod
    def _selfinyear_condition_is_always_true(op: str, year: int) -> bool:
        """Retourne True si la condition SELFINYEAR() OP YEAR est toujours vraie (année > 2023)."""
        # On suppose l'année de reporting toujours > 2023 (i.e. >= 2024)
        if op in (">", ">="):
            return year <= 2023   # SELFINYEAR() > 2023 → toujours TRUE si year=2023
        if op in ("<", "<="):
            return False          # SELFINYEAR() < YEAR ne peut pas être toujours vrai
        return False

    @staticmethod
    def _selfinyear_condition_is_always_false(op: str, year: int) -> bool:
        """Retourne True si la condition SELFINYEAR() OP YEAR est toujours fausse."""
        if op in ("<", "<="):
            return year <= 2024   # VALUE(SELFINYEAR()) < 2024 → toujours FALSE
        if op in (">", ">="):
            return False
        return False

    @classmethod
    def _simplify_selfinyear(cls, formula: str) -> str:
        """Simplifier les conditions basées sur SELFINYEAR().

        L'année de reporting est toujours > 2023 (≥ 2024), donc :
          SELFINYEAR() > 2023         → toujours TRUE  → IF(..., then, else) → then
          VALUE(SELFINYEAR()) < YEAR  (YEAR ≤ 2024) → toujours FALSE → IF(AND(...),then,else) → else

        L'itération permet de simplifier les IFs imbriqués.
        """
        # Patterns couverts :
        # 1. IF(SELFINYEAR()>YEAR, then, else)   [sans VALUE(), condition TRUE → garder then]
        # 2. IF(AND(VALUE(SELFINYEAR())<YEAR,...),then,else) [condition FALSE → garder else]
        patterns = [
            # Pattern 1 : SELFINYEAR() > YEAR  (toujours TRUE → garder branche THEN)
            (re.compile(
                r'(?:VALUE\()?SELFINYEAR\(\)(?:\))?\s*([><=!]+)\s*(\d{4})',
                re.IGNORECASE,
            ), "direct"),
            # Pattern 2 : AND(VALUE(SELFINYEAR()) < YEAR, ...) (toujours FALSE → garder ELSE)
            (re.compile(
                r'AND\(VALUE\(SELFINYEAR\(\)\)\s*([<>=!]+)\s*(\d{4})\s*,',
                re.IGNORECASE,
            ), "and_false"),
        ]

        result = formula
        for _ in range(30):  # max 30 simplifications
            changed = False
            for pattern, mode in patterns:
                m = pattern.search(result)
                if not m:
                    continue
                op = m.group(1)
                year = int(m.group(2))

                if mode == "direct":
                    if not cls._selfinyear_condition_is_always_true(op, year):
                        continue
                    # Chercher le IF( qui contient cette condition
                    before = result[: m.start()]
                    if_start = before.rfind("IF(")
                    if if_start < 0:
                        continue
                    # Parser IF(cond, then, else) → garder ELSE
                    # (le THEN "" est un garde désactivant la formule → on veut la branche else)
                    pos = if_start + 3
                    depth = 1
                    cond_end = -1
                    then_end = -1
                    i = pos
                    while i < len(result) and depth > 0:
                        ch = result[i]
                        if ch == "(":
                            depth += 1
                        elif ch == ")":
                            depth -= 1
                            if depth == 0:
                                break
                        elif ch == "," and depth == 1:
                            if cond_end < 0:
                                cond_end = i
                            elif then_end < 0:
                                then_end = i
                        i += 1
                    if cond_end < 0 or depth != 0:
                        continue
                    if_end = i
                    if then_end >= 0:
                        # IF a 3 arguments → garder ELSE (ignorer le garde THEN)
                        else_part = result[then_end + 1 : if_end]
                    else:
                        # IF a 2 arguments, pas d'ELSE → remplacer par vide
                        else_part = '""'
                    result = result[:if_start] + else_part + result[if_end + 1 :]
                    changed = True
                    break

                elif mode == "and_false":
                    if not cls._selfinyear_condition_is_always_false(op, year):
                        continue
                    # Chercher le IF( qui contient ce AND(
                    before = result[: m.start()]
                    if_start = before.rfind("IF(")
                    if if_start < 0:
                        continue
                    # Parser IF(cond, then, else) → garder ELSE
                    pos = if_start + 3
                    depth = 1
                    cond_end = -1
                    then_end = -1
                    i = pos
                    while i < len(result) and depth > 0:
                        ch = result[i]
                        if ch == "(":
                            depth += 1
                        elif ch == ")":
                            depth -= 1
                            if depth == 0:
                                break
                        elif ch == "," and depth == 1:
                            if cond_end < 0:
                                cond_end = i
                            elif then_end < 0:
                                then_end = i
                        i += 1
                    if cond_end < 0 or then_end < 0 or depth != 0:
                        continue
                    if_end = i
                    else_part = result[then_end + 1 : if_end]
                    result = result[:if_start] + else_part + result[if_end + 1 :]
                    changed = True
                    break

            if not changed:
                break

        return result

    # ─────────────────────────────────────────────────────────────────────────
    # Conversion d'une formule
    # ─────────────────────────────────────────────────────────────────────────

    def _convert_formula(
        self,
        formula: str,
        col_header: dict[str, str],
        cell_ref_map: dict[str, str],
        formula_row: int,
        institution_row: Optional[int] = None,
    ) -> str:
        """Remplacer les références de cellules Excel par des expressions fact[...].

        Algorithme :
          1. Segmenter la formule en parties textuelles et chaînes littérales.
          2. Dans les parties textuelles :
             a. Remplacer les plages institution (SUM(COL37:COL39)) par un fact
                sans SRF_LEI (agrégation applicative).
             b. Remplacer les références de cellules uniques par fact[...].
             c. Convertir les séparateurs d'arguments "," → ";".
          3. Réassembler.
          4. Post-processing :
             - SUM(fact[...institution...].value) → fact[...].value
               (l'appli agrège automatiquement, SUM est implicite)
             - COUNTIF(fact[...].value;"X") → LIST_COUNT({fact[...;SRF_LEI=?;value="X"].value})
        """
        # Pré-processing : simplifier les conditions SELFINYEAR() (année toujours > 2023)
        formula = self._simplify_selfinyear(formula)

        # Séparer les chaînes littérales (ex: "No", "") pour ne pas les modifier
        str_re = re.compile(r'"[^"]*"')
        segments = str_re.split(formula)
        literals = str_re.findall(formula)

        converted_segments: list[str] = []
        for i, segment in enumerate(segments):
            converted = self._replace_cell_refs(
                segment, col_header, cell_ref_map, formula_row, institution_row
            )
            converted = converted.replace(",", ";")
            converted_segments.append(converted)
            if i < len(literals):
                converted_segments.append(literals[i])

        result = "".join(converted_segments)

        # Post-processing 0 : ISBLANK(x) → x=""
        result = re.sub(r'\bISBLANK\(([^)]+)\)', r'\1=""', result)

        # Post-processing 1 : SUM(fact...) → supprimer le wrapper SUM.
        # Si multi-plages (;), les joindre avec + en ne remplaçant que les ';'
        # qui précèdent un 'fact[' (séparateurs de plages), pas ceux à l'intérieur
        # des crochets de fact[...].
        result = _SUM_FACT_RE.sub(
            lambda m: ' + '.join(re.split(r';\s*(?=fact\[)', m.group(1))),
            result
        )

        # Post-processing 2 : COUNTIF(fact[...].value;critère) → LIST_COUNT({fact[...;SRF_LEI=?;...].value})
        # Forme 1 : critère string simple "LPS" → value="LPS"
        result = _COUNTIF_STR_RE.sub(
            r'LIST_COUNT({fact[\1; SRF_LEI=?; value="\2"].value})', result
        )
        # Forme 2 : critère comparaison inline ">="&0 → value>=0
        result = _COUNTIF_CMP_INLINE_RE.sub(
            r'LIST_COUNT({fact[\1; SRF_LEI=?; value\2\3].value})', result
        )
        # Forme 3 : critère comparaison concaténée ">"&0 → value>0
        result = _COUNTIF_CMP_CONCAT_RE.sub(
            r'LIST_COUNT({fact[\1; SRF_LEI=?; value\2\3].value})', result
        )

        return result

    def _replace_cell_refs(
        self,
        text: str,
        col_header: dict[str, str],
        cell_ref_map: dict[str, str],
        formula_row: int,
        institution_row: Optional[int] = None,
    ) -> str:
        """Substituer les références de cellules dans un segment non-littéral.

        Traitement en deux passes :
          1. Plages → résolution selon le type de plage :
             a. Plage institution (COL37:COL39) → fact sans SRF_LEI (agrégation applicative).
             b. Plage summary même colonne (COL18:COL30) → expansion en
                fact[A].value + fact[B].value + ... pour tous les indicateurs trouvés.
          2. Cellules uniques → fact avec ou sans SRF_LEI selon type d'indicateur.
        """
        # Passe 1 : plages
        def _sub_range(m: re.Match) -> str:
            col1 = m.group(1).lstrip("$").upper()
            row1 = int(m.group(2))
            col2 = m.group(3).lstrip("$").upper()
            row2 = int(m.group(4))

            # Cas A : plage institution mono-colonne (COL37:COL39) → fact sans SRF_LEI
            if institution_row is not None and row1 == institution_row and col1 == col2:
                abs_key = f"${col1}${row1}"
                if abs_key not in cell_ref_map:
                    return m.group(0)  # Colonne sans indicateur → laisser tel quel
                col_code = cell_ref_map[abs_key]
                indicator_name, _ = self._resolve_col_code(col_code)
                # Agrégation sur tous les LEI → sans SRF_LEI
                return f'fact[IndicatorName="{indicator_name}"; endDate=?].value'

            # Cas D : plage multi-colonnes sur la ligne dynamique (T28:W28)
            # → somme de tous les facts avec SRF_LEI (une valeur par LEI)
            if institution_row is not None and row1 == institution_row and row2 == institution_row and col1 != col2:
                from openpyxl.utils import column_index_from_string, get_column_letter
                idx1 = column_index_from_string(col1)
                idx2 = column_index_from_string(col2)
                parts: list[str] = []
                for idx in range(min(idx1, idx2), max(idx1, idx2) + 1):
                    c = get_column_letter(idx)
                    abs_key = f"${c}${institution_row}"
                    if abs_key in cell_ref_map:
                        col_code = cell_ref_map[abs_key]
                        indicator_name, is_static = self._resolve_col_code(col_code)
                        parts.append(_make_fact(indicator_name, is_static))
                if parts:
                    return " + ".join(parts)
                return m.group(0)

            # Cas B : plage summary (même colonne, plusieurs lignes DATA au-dessus)
            # Expander en fact[A].value + fact[B].value + ...
            if col1 == col2:
                parts: list[str] = []
                for row in range(min(row1, row2), max(row1, row2) + 1):
                    abs_key = f"${col1}${row}"
                    if abs_key in cell_ref_map:
                        col_code = cell_ref_map[abs_key]
                        indicator_name, _ = self._resolve_col_code(col_code)
                        # Indicateurs summary/ratio → pas de SRF_LEI
                        parts.append(f'fact[IndicatorName="{indicator_name}"; endDate=?].value')
                if parts:
                    return " + ".join(parts)

            # Cas C : plage multi-colonnes ou non résolue → laisser tel quel
            return m.group(0)

        text = _RANGE_REF_RE.sub(_sub_range, text)

        # La ligne dynamique (par LEI) : institution_row si fourni, sinon formula_row.
        dynamic_row = institution_row if institution_row is not None else formula_row

        # Passe 2 : cellules uniques restantes
        def _sub(match: re.Match) -> str:
            col = match.group(1).upper()
            row = int(match.group(2))

            # 1. Recherche dans la carte des références absolues (#CELL metadata)
            #    La valeur peut être un code court (ex: "1A1") ou un indicateur
            #    direct (ex: "S0PARA_TALEC") → toujours passer par _resolve_col_code.
            abs_key = f"${col}${row}"
            if abs_key in cell_ref_map:
                col_code = cell_ref_map[abs_key]
                indicator_name, is_static = self._resolve_col_code(col_code)
                # Forcer statique si la cellule n'est pas sur la ligne dynamique :
                # les lignes de ratio/totaux au-dessus ne varient pas par LEI.
                if row != dynamic_row:
                    is_static = True
                return _make_fact(indicator_name, is_static)

            # 2. Référence à la ligne de formules → résolution via en-tête de colonne
            if row == formula_row:
                code = col_header.get(col)
                if code:
                    indicator_name, is_static = self._resolve_col_code(code)
                    return _make_fact(indicator_name, is_static)

            # 3. Non résolu → laisser tel quel
            return match.group(0)

        return _CELL_REF_RE.sub(_sub, text)

    # ─────────────────────────────────────────────────────────────────────────
    # Utilitaires
    # ─────────────────────────────────────────────────────────────────────────

    def _resolve_col_code(self, col_code: str) -> tuple[str, bool]:
        """Résoudre un code de colonne en (indicator_name, is_static).

        Si le code n'est pas dans le mapping, il est lui-même l'indicator_name
        (c'est un nouvel indicateur défini par la feuille en cours).
        """
        entry = self.mapping.get(col_code)
        if entry:
            return entry["indicator_name"], entry.get("static", False)
        return col_code, False  # Nouvel indicateur → dynamique par défaut

    def _is_static(self, indicator_name: str) -> bool:
        """Vérifier si un indicator_name correspond à un indicateur statique."""
        # Cherche d'abord comme clé directe (ex: S0PARA_TALEC)
        entry = self.mapping.get(indicator_name)
        if entry:
            return entry.get("static", False)
        # Cherche dans l'index inverse
        return self._indicator_static.get(indicator_name, False)


# ─────────────────────────────────────────────────────────────────────────────
# Fonctions pures
# ─────────────────────────────────────────────────────────────────────────────


def _make_fact(indicator_name: str, is_static: bool) -> str:
    """Construire l'expression fact[...] selon le type d'indicateur."""
    if is_static:
        return f'fact[IndicatorName="{indicator_name}"; endDate=?].value'
    return f'fact[IndicatorName="{indicator_name}"; SRF_LEI=?; endDate=?].value'


# Regex pour détecter les références Excel brutes non résolues dans une formule convertie.
# Matches : lettre(s) majuscules + chiffres, non précédés/suivis de lettres, chiffres, _  ou "
_RAW_REF_RE = re.compile(r'(?<![A-Za-z0-9_"=])\$?[A-Z]{1,3}\$?\d{1,7}(?![A-Za-z0-9_"])')


def formula_status(formula: str) -> str:
    """Retourner 'OK' si la formule est correctement convertie, 'KO' sinon.

    Une formule est KO si elle contient :
      - des références de cellules Excel brutes (ex: G22, $DV$37)
      - un COUNTIF non transformé
      - un SUM( non transformé (le wrapper SUM doit avoir été supprimé)
    """
    if _RAW_REF_RE.search(formula):
        return "KO"
    if re.search(r'\bCOUNTIF\b', formula, re.IGNORECASE):
        return "KO"
    if re.search(r'\bSUM\(', formula, re.IGNORECASE):
        return "KO"
    return "OK"
