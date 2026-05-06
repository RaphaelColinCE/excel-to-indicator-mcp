"""Service de conversion Excel → Formule d'indicateur"""

import logging
import re
from typing import List, Optional

from src.models.excel_models import ExcelFormula, ReferenceType
from src.models.result_models import (
    ConversionResult,
    ConversionStatus,
    DependencyNode,
    DependencyTrace,
    MappingUsage,
    ValidationReport,
    ValidationError as ValidationErrorModel,
)
from src.services.formula_parser import parse_formula
from src.services.mapping_service import MappingService

logger = logging.getLogger("mcp_server")

# Séparateur d'arguments utilisé dans les formules de l'application cible
FORMULA_SEP = "; "


class ConversionService:
    """Convertit des formules Excel en formules d'indicateurs via le mapping."""

    def __init__(self, mapping_service: MappingService, max_depth: int = 5):
        self.mapping_service = mapping_service
        self.max_depth = max_depth

    # ──────────────────────────────────────────────────────────────────────────
    # API principale
    # ──────────────────────────────────────────────────────────────────────────

    def convert_formula(
        self,
        formula: str,
        sheet: str,
        cell: str,
        file_path: str = "",
    ) -> ConversionResult:
        """Convertir une formule Excel en formule d'indicateur.

        Args:
            formula: Formule Excel brute (ex: =SUM(Sheet1!A1:A10))
            sheet: Feuille contenant la formule
            cell: Cellule contenant la formule
            file_path: Chemin du fichier Excel (optionnel)

        Returns:
            ConversionResult avec la formule mappée et les métadonnées
        """
        try:
            from src.services.excel_converter import ExcelConverterService
            formula = ExcelConverterService._simplify_selfinyear(formula)
            parsed = parse_formula(formula, sheet, cell, file_path)
            return self._convert_parsed(parsed)
        except Exception as e:
            logger.error(f"Erreur lors de la conversion: {e}")
            return ConversionResult(
                status=ConversionStatus.ERROR,
                original_formula=formula,
                explanation=f"Erreur inattendue: {str(e)}",
                errors=[str(e)],
                excel_file=file_path,
                sheet=sheet,
                cell=cell,
            )

    def validate_formula(self, formula: str) -> ValidationReport:
        """Valider une formule d'indicateur (post-conversion).

        Args:
            formula: Formule à valider

        Returns:
            ValidationReport avec les erreurs et avertissements
        """
        errors: List[ValidationErrorModel] = []
        warnings: List[ValidationErrorModel] = []
        rules_applied: List[str] = []

        # Règle 1 : parenthèses équilibrées
        rules_applied.append("balanced_parentheses")
        if not _check_balanced(formula):
            errors.append(
                ValidationErrorModel(
                    type="syntax",
                    message="Parenthèses déséquilibrées",
                    severity="error",
                )
            )

        # Règle 2 : références Excel brutes résiduelles (ex: A1, Sheet1!B2)
        rules_applied.append("no_raw_excel_refs")
        raw_refs = re.findall(r"'?[A-Za-z0-9 _]+'?![A-Z]+\d+|(?<![A-Za-z])[A-Z]+\d+(?![A-Za-z])", formula)
        if raw_refs:
            warnings.append(
                ValidationErrorModel(
                    type="unresolved_reference",
                    message=f"Références Excel brutes détectées: {raw_refs}",
                    severity="warning",
                )
            )

        # Règle 3 : formule non vide
        rules_applied.append("non_empty")
        if not formula.strip():
            errors.append(
                ValidationErrorModel(
                    type="empty",
                    message="La formule est vide",
                    severity="error",
                )
            )

        complexity = min(1.0, len(formula) / 500)
        is_valid = len(errors) == 0

        return ValidationReport(
            formula=formula,
            is_valid=is_valid,
            errors=errors,
            warnings=warnings,
            validation_rules_applied=rules_applied,
            complexity_score=complexity,
            confidence_score=1.0 - 0.3 * len(warnings) - 0.5 * len(errors),
        )

    def trace_dependencies(
        self,
        formula: str,
        sheet: str,
        cell: str,
        file_path: str = "",
        current_depth: int = 0,
    ) -> DependencyTrace:
        """Tracer l'arbre de dépendances d'une formule.

        Args:
            formula: Formule Excel brute
            sheet: Feuille contenant la formule
            cell: Cellule contenant la formule
            file_path: Chemin du fichier (optionnel)
            current_depth: Profondeur actuelle (pour la récursion)

        Returns:
            DependencyTrace avec le graphe de dépendances
        """
        parsed = parse_formula(formula, sheet, cell, file_path)
        root = self._build_dependency_node(parsed, current_depth)

        all_refs = self._collect_all_refs(root)
        unresolvable = [
            r for r in all_refs
            if self.mapping_service.get_reference_mapping(r) is None
        ]

        return DependencyTrace(
            original_formula=formula,
            root_node=root,
            total_depth=_max_depth_node(root),
            total_references=len(all_refs),
            unresolvable_references=unresolvable,
        )

    # ──────────────────────────────────────────────────────────────────────────
    # Méthodes internes
    # ──────────────────────────────────────────────────────────────────────────

    def _convert_parsed(self, parsed: ExcelFormula) -> ConversionResult:
        """Effectuer la conversion à partir d'une formule parsée."""
        mapped_formula = parsed.original_formula
        mapping_used: List[MappingUsage] = []
        unresolved: List[str] = []
        warnings: List[str] = []

        for ref in parsed.references:
            ref_key = ref.value
            mapping = self.mapping_service.get_reference_mapping(ref_key)

            if mapping is None and ref.sheet:
                # Essayer avec le nom de la feuille uniquement
                mapping = self.mapping_service.get_reference_mapping(ref.sheet)

            if mapping:
                mapped_formula = mapped_formula.replace(ref_key, mapping.concept)
                mapping_used.append(
                    MappingUsage(
                        excel_reference=ref_key,
                        concept=mapping.concept,
                        indicator=mapping.indicator,
                        dimensions=mapping.dimensions,
                        mapped_to=mapping.concept,
                    )
                )
            else:
                unresolved.append(ref_key)

        # Appliquer les mappings de fonctions Excel
        for func in parsed.functions_used:
            func_mapping = self.mapping_service.get_function_mapping(func)
            if func_mapping:
                target = func_mapping.get("target_function", func)
                mapped_formula = re.sub(
                    rf"\b{re.escape(func)}\b", target, mapped_formula, flags=re.IGNORECASE
                )

        # Déterminer le statut
        if unresolved and mapping_used:
            status = ConversionStatus.PARTIALLY_RESOLVED
        elif unresolved and not mapping_used:
            status = ConversionStatus.UNRESOLVED
        else:
            status = ConversionStatus.RESOLVED

        if unresolved:
            warnings.append(f"Références non résolues: {unresolved}")

        explanation = self._build_explanation(status, mapping_used, unresolved)

        return ConversionResult(
            status=status,
            original_formula=parsed.original_formula,
            mapped_formula=mapped_formula,
            explanation=explanation,
            mapping_used=mapping_used,
            unresolved_references=unresolved,
            excel_file=parsed.file_path,
            sheet=parsed.sheet,
            cell=parsed.cell,
            warnings=warnings,
        )

    def _build_dependency_node(
        self, parsed: ExcelFormula, depth: int
    ) -> DependencyNode:
        """Construire un nœud de dépendances récursivement."""
        children: List[DependencyNode] = []

        if depth < self.max_depth:
            for ref in parsed.references:
                child = DependencyNode(
                    reference=ref.value,
                    sheet=ref.sheet,
                    cell=ref.start_cell,
                    type=ref.type if isinstance(ref.type, str) else ref.type.value,
                    depth=depth + 1,
                )
                children.append(child)

        return DependencyNode(
            reference=f"{parsed.sheet}!{parsed.cell}",
            sheet=parsed.sheet,
            cell=parsed.cell,
            type="formula",
            depth=depth,
            children=children,
        )

    def _collect_all_refs(self, node: DependencyNode) -> List[str]:
        """Collecter toutes les références dans l'arbre."""
        refs = [node.reference]
        for child in node.children:
            refs.extend(self._collect_all_refs(child))
        return refs

    @staticmethod
    def _build_explanation(
        status: ConversionStatus,
        mapping_used: List[MappingUsage],
        unresolved: List[str],
    ) -> str:
        if status == ConversionStatus.RESOLVED:
            concepts = [m.concept for m in mapping_used]
            return f"Conversion complète. Concepts utilisés: {', '.join(concepts)}."
        elif status == ConversionStatus.PARTIALLY_RESOLVED:
            return (
                f"{len(mapping_used)} référence(s) résolue(s), "
                f"{len(unresolved)} non résolue(s): {unresolved}."
            )
        elif status == ConversionStatus.UNRESOLVED:
            return f"Impossible de résoudre les références: {unresolved}."
        return "Erreur lors de la conversion."


# ──────────────────────────────────────────────────────────────────────────────
# Fonctions utilitaires
# ──────────────────────────────────────────────────────────────────────────────

def _check_balanced(formula: str) -> bool:
    """Vérifier que les parenthèses sont équilibrées."""
    count = 0
    in_string = False
    for ch in formula:
        if ch == '"':
            in_string = not in_string
        if in_string:
            continue
        if ch == "(":
            count += 1
        elif ch == ")":
            count -= 1
        if count < 0:
            return False
    return count == 0


def _max_depth_node(node: DependencyNode) -> int:
    """Calculer la profondeur maximale de l'arbre."""
    if not node.children:
        return node.depth
    return max(_max_depth_node(c) for c in node.children)
