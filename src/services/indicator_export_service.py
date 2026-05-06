"""Export des indicateurs convertis vers XML et spécification Excel.

Deux exports possibles depuis les records produits par ExcelConverterService :
  1. export_to_xml   → fichier XML importable dans l'appli ISS
  2. export_analysis_spec → fichier Excel de spécification de tableau d'analyse
"""

from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Optional

import openpyxl

logger = logging.getLogger("mcp_server")

# Regex pour extraire le code d'un OUTPUTMCCOMP/INPUTMCCOMP dans une cellule #CELL
_MC_COMP_RE = re.compile(r"(?:OUTPUT|INPUT)MCCOMP:\{[^}]*\}([^;\n]+)")

# Correspondance type numérique (IndicatorsDetails.xlsx) → chaîne XML
_TYPE_INT_MAP: dict[int, str] = {0: "MONETARY", 2: "REAL", 4: "DATE", 5: "TEXT"}


class IndicatorExportService:
    """Service d'export des indicateurs vers XML et spécification Excel."""

    def __init__(self, details_path: Path, mapping_path: Path) -> None:
        self.indicator_details: dict[str, dict] = self._load_details(details_path)
        with open(mapping_path, encoding="utf-8") as fh:
            data = json.load(fh)
        self.mapping: dict = data.get("mapping", {})
        # Charge les indicateurs RS/XBRL depuis les fichiers XML du répertoire indicators
        self.xml_indicator_index: dict[str, tuple[str, str]] = self._load_xml_indicators(
            details_path.parent.parent / "indicators"
        )

    # ─────────────────────────────────────────────────────────────────────────
    # Initialisation
    # ─────────────────────────────────────────────────────────────────────────

    def _load_details(self, path: Path) -> dict[str, dict]:
        """Charge IndicatorsDetails.xlsx → {code: {label, data_type}}."""
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        code_col = headers.index("Component code")
        label_col = headers.index("Component label")
        type_col = headers.index("Type")
        details: dict[str, dict] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            code = row[code_col]
            if not code:
                continue
            label = row[label_col] or str(code)
            t = row[type_col]
            details[str(code)] = {
                "label": str(label),
                "data_type": _TYPE_INT_MAP.get(t, "REAL"),
            }
        return details

    def _load_xml_indicators(self, indicators_dir: Path) -> dict[str, tuple[str, str]]:
        """Charge les XML d'indicateurs RS → {column_code: (full_code, class)}.

        - S0DE / S0PARA / S0PILA → class 'RS'
        - LOI / DGS              → class 'XBRL'
        Indexe à la fois par ColumnCode et par Code complet.
        """
        import xml.etree.ElementTree as ET

        _RS_FRAMEWORKS  = ("S0DE", "S0PARA", "S0PILA")
        _XBRL_FRAMEWORKS = ("LOI", "DGS")
        index: dict[str, tuple[str, str]] = {}

        if not indicators_dir.exists():
            return index

        for xml_file in indicators_dir.glob("*.xml"):
            try:
                root = ET.parse(xml_file).getroot()
            except ET.ParseError:
                continue
            for ind in root.iter("Indicator"):
                full_code = ind.findtext("Code", "").strip()
                col_code  = ind.findtext("ColumnCode", "").strip()
                fw_code   = ind.findtext("Framework/Code", "").strip()
                if not full_code:
                    continue
                # Déterminer la classe
                if fw_code in _RS_FRAMEWORKS or any(full_code.startswith(f) for f in _RS_FRAMEWORKS):
                    cls = "RS"
                elif fw_code in _XBRL_FRAMEWORKS or any(full_code.startswith(f) for f in _XBRL_FRAMEWORKS):
                    cls = "XBRL"
                else:
                    cls = "RS"  # défaut pour les fichiers indicateurs
                # Indexer par ColumnCode et par Code complet
                if col_code:
                    index.setdefault(col_code, (full_code, cls))
                index[full_code] = (full_code, cls)
        return index

    # ─────────────────────────────────────────────────────────────────────────
    # Résolution de codes
    # ─────────────────────────────────────────────────────────────────────────

    def _resolve_full_name(self, code: str) -> tuple[str, str]:
        """Résoudre un code court en (full_indicator_name, indicator_class).

        indicator_class : 'XBRL' | 'RS' | 'CUSTOM'
        Priorité :
          1. Mapping JSON direct (EACIND_/CMGT_ = XBRL, sinon RS)
          2. S0DE_DGS_xxx / S0DE_LOI_xxx → strip "S0DE_" puis re-résoudre
          3. Index XML des indicateurs (S0DE/S0PARA/S0PILA = RS, LOI/DGS = XBRL)
          4. Heuristique sur le préfixe du code
        """
        entry = self.mapping.get(code)
        if entry:
            full: str = entry["indicator_name"]
            if full.startswith(("EACIND_", "CMGT_")):
                return full, "XBRL"
            return full, "RS"
        # S0DE_DGS_xxx / S0DE_LOI_xxx → strip "S0DE_" et re-résoudre via mapping
        if code.startswith("S0DE_DGS_") or code.startswith("S0DE_LOI_"):
            return self._resolve_full_name(code[5:])   # "S0DE_DGS_2A3" → "DGS_2A3"
        # Index XML (S0DE, S0PARA, S0PILA…)
        if code in self.xml_indicator_index:
            return self.xml_indicator_index[code]
        # Heuristique sur le préfixe
        if code.startswith("EACIND_"):
            return code, "XBRL"
        for rs_prefix in ("S0PARA_", "S0PILA_", "S0DE_"):
            if code.startswith(rs_prefix):
                return code, "RS"
        for xbrl_prefix in ("LOI_", "DGS_", "CMGT_"):
            if code.startswith(xbrl_prefix):
                return code, "XBRL"
        return code, "CUSTOM"

    def _indicator_cell(self, code: str, is_dimension: bool = False) -> str:
        """Générer la chaîne `Indicator[TYPE]:name;` ou `Dimension:SRF_LEI;`."""
        if is_dimension:
            return "Dimension:SRF_LEI;"
        full, cls = self._resolve_full_name(code)
        return f"Indicator[{cls}]:{full};"

    # ─────────────────────────────────────────────────────────────────────────
    # Export 1 : XML
    # ─────────────────────────────────────────────────────────────────────────

    def export_to_xml(
        self,
        records: list[dict],
        output_path: Path,
        data_types: Optional[dict[str, str]] = None,
    ) -> int:
        """Exporter les indicateurs convertis au format XML pour l'appli ISS.

        Args:
            records    : liste de dicts avec clés indicator_name / indicator_formula.
            output_path: chemin du fichier XML à créer.
            data_types : surcharge optionnelle {indicator_name: 'MONETARY'|'REAL'|...}.

        Returns:
            Nombre d'indicateurs exportés.
        """
        lines = [
            '<?xml version="1.0" encoding="utf-8"?>',
            '<Indicators xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
            ' xmlns:xsd="http://www.w3.org/2001/XMLSchema">',
            "  <CustomIndicators>",
        ]

        # Regex pour détecter les références Excel brutes non résolues (ex: G22, CP18)
        _raw_cell_re = re.compile(r'(?<![A-Za-z0-9_"=])\$?[A-Z]{1,3}\$?\d{1,7}(?![A-Za-z0-9_"])')

        for rec in records:
            code = rec["indicator_name"]
            formula = rec["indicator_formula"]
            detail = self.indicator_details.get(code, {})
            label_raw = detail.get("label", code)
            label = (
                str(label_raw)
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
            )

            # Priorité : surcharge explicite > IndicatorsDetails > REAL par défaut
            if data_types and code in data_types:
                data_type = data_types[code]
            else:
                data_type = detail.get("data_type", "REAL")

            # Si la formule contient encore des refs Excel brutes (pas résolues),
            # mettre "0" pour permettre l'import sans erreur.
            if _raw_cell_re.search(formula) or not formula.strip():
                formula = "0"

            formula_esc = (
                formula.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
            )

            lines += [
                "    <Indicator>",
                f"      <Code>{code}</Code>",
                f"      <Label>{label}</Label>",
                "      <RowCode />",
                "      <ColumnCode />",
                "      <IndicatorType>CUSTOM</IndicatorType>",
                f"      <DataType>{data_type}</DataType>",
                f"      <Formula>{formula_esc}</Formula>",
                "    </Indicator>",
            ]

        lines += [
            "  </CustomIndicators>",
            "  <RSIndicators />",
            "  <AnalyzerIndicators />",
            "  <XBRLIndicators />",
            "</Indicators>",
        ]

        output_path.write_text("\n".join(lines), encoding="utf-8")
        logger.info("XML exporté: %s (%d indicateurs)", output_path.name, len(records))
        return len(records)

    # ─────────────────────────────────────────────────────────────────────────
    # Export 2 : Spécification Excel
    # ─────────────────────────────────────────────────────────────────────────

    def export_analysis_spec(
        self,
        input_excel_path: Path,
        sheet_name: str,
        output_path: Path,
        header_row: int = 25,   # conservé pour compatibilité API
        formula_row: int = 26,  # conservé pour compatibilité API
    ) -> int:
        """Générer le fichier de spécification Excel pour le tableau d'analyse.

        Approche : copier l'Excel input à l'identique, puis :
          1. Supprimer les lignes violettes (FF9400D3) et la ligne rouge (FF800000)
          2. Supprimer la colonne B (violet en ligne 1)
          3. Remplacer les cellules orange (#CELL) avec la notation Indicator[...]
          4. Transformer la ligne #DATA_DYNAMIQUE en ligne verte Dynamic:SRF_LEI;
          5. Ajouter la feuille #TEMPLATE

        Seuls les cellules orange et vertes sont modifiées.
        Tout le reste (couleurs, tailles, valeurs) est conservé à l'identique.
        """
        import shutil
        import openpyxl.worksheet.cell_range
        from openpyxl.styles import PatternFill, Font

        PURPLE   = "FF9400D3"
        RED_ROW  = "FF800000"
        GREEN    = "008000"

        # ── 1. Copier le fichier source ──────────────────────────────────────
        shutil.copy(input_excel_path, output_path)
        wb = openpyxl.load_workbook(output_path)
        ws = wb[sheet_name]

        # ── 2. Identifier les lignes à supprimer (violet + rouge en col A) ──
        rows_to_delete: list[int] = []
        for r in range(1, ws.max_row + 1):
            cell_a = ws.cell(r, 1)
            try:
                bg = cell_a.fill.fgColor.rgb
            except Exception:
                bg = "00000000"
            if bg in (PURPLE, RED_ROW):
                rows_to_delete.append(r)

        # ── 3. Avant suppression : capturer la config des #CELL rows ─────────
        # On mémorise {row_original: {col_orig: cell_value}} pour les lignes #CELL
        # (les indices de lignes changent après delete_rows)
        cell_configs: dict[int, dict[int, str]] = {}
        dyn_row_orig: Optional[int] = None
        dyn_cell_row_orig: Optional[int] = None
        # Codes de colonnes issus de la ligne header juste avant #DATA_DYNAMIQUE
        # (ex: S0DE_DGS_2A3, S0DE_LOI_IPS3) → source de vérité pour les indicateurs
        header_col_codes: dict[int, str] = {}

        for r in range(1, ws.max_row + 1):
            val_a = ws.cell(r, 1).value
            if isinstance(val_a, str):
                if val_a == "#DATA_DYNAMIQUE":
                    dyn_row_orig = r
                    # La ligne juste au-dessus est la ligne header avec les codes indicateurs
                    prev_r = r - 1
                    for c in range(2, ws.max_column + 1):
                        cv = ws.cell(prev_r, c).value
                        if cv and isinstance(cv, str) and not cv.startswith("="):
                            header_col_codes[c] = cv.strip()
                elif val_a == "#CELL":
                    config: dict[int, str] = {}
                    for c in range(2, ws.max_column + 1):
                        cv = ws.cell(r, c).value
                        if cv and isinstance(cv, str) and (
                            "OUTPUTMCCOMP" in cv or "INPUTMCCOMP" in cv
                        ):
                            config[c] = cv
                    if config:
                        cell_configs[r] = config

        if dyn_row_orig:
            # La ligne #CELL dynamique suit immédiatement #DATA_DYNAMIQUE
            dyn_cell_row_orig = dyn_row_orig + 1

        # ── 4. Supprimer les lignes (du bas vers le haut) ────────────────────
        for r in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(r)

        # ── 5. Supprimer la colonne B ────────────────────────────────────────
        ws.delete_cols(2)

        # ── 5b. Démerger toutes les cellules fusionnées ───────────────────────
        # openpyxl.unmerge_cells peut lever KeyError sur des plages partiellement
        # supprimées → on vide directement le set des merges.
        ws.merged_cells = openpyxl.worksheet.cell_range.MultiCellRange()

        # ── 6. Calculer les nouveaux indices après suppressions ──────────────
        n_deleted_rows_before = {
            orig_r: sum(1 for dr in rows_to_delete if dr < orig_r)
            for orig_r in list(cell_configs.keys()) + (
                [dyn_row_orig, dyn_cell_row_orig]
                if dyn_row_orig else []
            )
        }

        def new_row(orig_r: int) -> int:
            return orig_r - n_deleted_rows_before.get(orig_r, 0)

        # Après delete_cols(2) : toutes les colonnes >= 3 sont décalées de -1
        # orig col 1 → new col 1 (A)
        # orig col 2 (B, supprimée) → n/a
        # orig col 3 (C) → new col 2 (B)
        # orig col 4 (D) → new col 3 (C)  ← position Dynamic:SRF_LEI;
        # orig col 5 (E) → new col 4 (D)  ← 1er indicateur data
        def new_col(orig_c: int) -> int:
            return orig_c - 1 if orig_c >= 3 else orig_c

        from openpyxl.cell.cell import MergedCell

        def _set_cell(ws, r: int, c: int, value, fill=None, font=None) -> None:
            """Modifier une cellule en ignorant silencieusement les MergedCell."""
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                return
            cell.value = value
            if fill is not None:
                cell.fill = fill
            if font is not None:
                cell.font = font

        # ── 6b. Nettoyer col B (#SHEET) et col C (#ROW) ─────────────────────
        # Après delete_cols(2) :
        #   col B (2) = ancienne col C (INPUTMCACC) → à vider → #SHEET
        #   col C (3) = ancienne col D (codes indicateurs) → à vider → #ROW
        # La seule valeur autorisée en col C est Dynamic:SRF_LEI; sur la ligne verte.
        for r in range(1, ws.max_row + 1):
            _set_cell(ws, r, 2, None)   # vider col B
            _set_cell(ws, r, 3, None)   # vider col C
        # En-têtes de colonnes métadonnées (ligne 1) + #SHEET en A2
        _set_cell(ws, 1, 2, "#SHEET")
        _set_cell(ws, 1, 3, "#ROW")
        _set_cell(ws, 2, 1, "#SHEET")

        # ── 7. Transformer la ligne #DATA_DYNAMIQUE ──────────────────────────
        # Ligne verte : seulement col C porte Dynamic:SRF_LEI; en vert.
        # Toutes les autres cellules de la ligne → blanc (pas de couleur globale).
        if dyn_row_orig:
            nr = new_row(dyn_row_orig)
            green_fill = PatternFill("solid", fgColor=GREEN)
            white_fill = PatternFill("solid", fgColor="FFFFFF")
            white_font = Font(color="000000", bold=False)
            for c in range(1, ws.max_column + 1):
                if c == 3:
                    continue  # col C traitée séparément
                _set_cell(ws, nr, c, None, fill=white_fill, font=white_font)
            # Dynamic:SRF_LEI; uniquement en col C (#ROW), fond vert
            _set_cell(ws, nr, 3, "Dynamic:SRF_LEI;",
                      fill=green_fill, font=Font(color="FFFFFF", bold=True))

        # ── 8. Transformer les lignes #CELL ──────────────────────────────────
        for orig_r, config in cell_configs.items():
            nr = new_row(orig_r)
            is_dyn_cell = (orig_r == dyn_cell_row_orig)

            for orig_c, cv_str in config.items():
                nc = new_col(orig_c)

                # DIMINPUTCODE → dimension LEI
                if "DIMINPUTCODE" in cv_str:
                    _set_cell(ws, nr, nc, "Dimension:SRF_LEI;")
                    continue

                # Pour la ligne #CELL dynamique uniquement : utiliser en priorité
                # le code de la ligne header (#DATA juste avant #DATA_DYNAMIQUE)
                # Ex: S0DE_DGS_2A3 → Indicator[XBRL]:CMGT_dgs__C02A3$;
                if is_dyn_cell:
                    header_code = header_col_codes.get(orig_c)
                    if header_code:
                        full, cls = self._resolve_full_name(header_code)
                        _set_cell(ws, nr, nc, f"Indicator[{cls}]:{full};")
                        continue

                # Extraction du code via OUTPUTMCCOMP / INPUTMCCOMP
                mc_m = _MC_COMP_RE.search(cv_str)
                if not mc_m:
                    _set_cell(ws, nr, nc, None)
                    continue

                comp_code = mc_m.group(1).strip().rstrip(";")
                full, cls = self._resolve_full_name(comp_code)
                _set_cell(ws, nr, nc, f"Indicator[{cls}]:{full};")

        # ── 9. Normaliser tous les tags #DATA* → #DATA dans toute la feuille ──
        # Exception : #DATA_DYNAMIQUE → None (la ligne verte n'a pas de tag en col A)
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                if isinstance(cell, MergedCell):
                    continue
                val = cell.value
                if not isinstance(val, str):
                    continue
                if val == "#DATA_DYNAMIQUE":
                    cell.value = None   # ligne verte → pas de tag
                elif val.startswith("#DATA") and val != "#DATA":
                    cell.value = "#DATA"

        # ── 10. Effacer les formules Excel des lignes #DATA ──────────────────
        # Les cellules de données contenant "=..." sont des formules Excel brutes
        # qui doivent être supprimées (remplacées par l'indicateur CUSTOM calculé).
        for r in range(1, ws.max_row + 1):
            val_a = ws.cell(r, 1).value
            if isinstance(val_a, str) and val_a == "#DATA":
                for c in range(2, ws.max_column + 1):
                    cell = ws.cell(r, c)
                    if isinstance(cell, MergedCell):
                        continue
                    v = cell.value
                    if isinstance(v, str) and v.startswith("="):
                        cell.value = None

        # ── 11. Afficher les lignes masquées + supprimer le # des labels ────────
        # Les tags techniques (#DATA, #CELL, etc.) sont en cols A, B, C.
        # En col D+, tout `#` en début de valeur est un label et doit être épuré.
        _TECH_TAGS = {"#DATA", "#CELL", "#SHEET", "#ROW", "#COLUMN",
                      "#TEMPLATE", "#VERSION", "#CODE", "#LABEL[en]"}
        for r in range(1, ws.max_row + 1):
            # Afficher les lignes masquées
            if ws.row_dimensions[r].hidden:
                ws.row_dimensions[r].hidden = False
            # Row 1 = ligne d'en-têtes techniques → ne pas toucher
            if r == 1:
                continue
            # Supprimer le # en début de label dans les colonnes de données (col D+)
            for c in range(4, ws.max_column + 1):
                cell = ws.cell(r, c)
                if isinstance(cell, MergedCell):
                    continue
                val = cell.value
                if isinstance(val, str) and val.startswith("#") and val not in _TECH_TAGS:
                    cell.value = val[1:]

        # ── 12. Insérer la ligne #COLUMN en A3 ───────────────────────────────
        # Structure finale : A1=vide / A2=#SHEET / A3=#COLUMN / A4+=#DATA...
        ws.insert_rows(3)
        ws.cell(3, 1).value = "#COLUMN"

        # ── 12b. Insérer une colonne #DATA vide en col D (entre #ROW et données)
        ws.insert_cols(4)
        ws.cell(1, 4).value = "#DATA"   # en-tête col D

        # ── 13. Supprimer les feuilles #TEMPLATE* héritées de l'input ────────
        for name in list(wb.sheetnames):
            if name.startswith("#TEMPLATE"):
                del wb[name]

        # ── 14. Ajouter la feuille #TEMPLATE ──────────────────────────────────
        blue_fill  = PatternFill("solid", fgColor="0000FF")
        white_font = Font(color="FFFFFF", bold=True)
        ws_tpl = wb.create_sheet("#TEMPLATE", 0)
        tpl_data = [
            [None,       "#DATA"],
            ["#VERSION", 3],
            ["#CODE",    sheet_name],
            ["#LABEL[en]", sheet_name],
        ]
        for r_idx, row_vals in enumerate(tpl_data, 1):
            for c_idx, val in enumerate(row_vals, 1):
                cell = ws_tpl.cell(r_idx, c_idx)
                cell.value = val
                cell.fill = blue_fill
                cell.font = white_font

        # ── 15. Renommer la feuille de données → "Data" ──────────────────────
        ws.title = "Data"

        wb.save(output_path)
        logger.info("Spec exportée: %s", output_path.name)
        return ws.max_column - 3  # approx n_data_cols
