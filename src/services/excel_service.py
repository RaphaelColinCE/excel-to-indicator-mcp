"""Service de lecture et extraction des formules Excel"""

import logging
from pathlib import Path
from typing import List, Optional, Dict, Any
from openpyxl import load_workbook

from src.models.excel_models import ExcelFormula, FormulaToken, ReferenceType, ExcelReference
from src.utils.exceptions import ExcelFileNotFoundError, ExcelParsingError

logger = logging.getLogger("mcp_server")


class ExcelService:
    """Service pour lire et extraire les données des fichiers Excel"""

    def __init__(self, excel_dir: Path):
        """Initialiser le service Excel
        
        Args:
            excel_dir: Répertoire contenant les fichiers Excel
        """
        self.excel_dir = Path(excel_dir)
        self.workbooks: Dict[str, Any] = {}  # Cache des workbooks chargés

    def load_file(self, file_name: str) -> Any:
        """Charger un fichier Excel
        
        Args:
            file_name: Nom du fichier (ex: "data.xlsx")
            
        Returns:
            Workbook openpyxl
            
        Raises:
            ExcelFileNotFoundError: Si le fichier n'existe pas
            ExcelParsingError: Si le fichier ne peut pas être parsé
        """
        # Vérifier le cache
        if file_name in self.workbooks:
            return self.workbooks[file_name]

        file_path = self.excel_dir / file_name
        
        if not file_path.exists():
            raise ExcelFileNotFoundError(f"Fichier non trouvé: {file_path}")

        try:
            workbook = load_workbook(file_path, data_only=False)
            self.workbooks[file_name] = workbook
            logger.info(f"Fichier Excel chargé: {file_name}")
            return workbook
        except Exception as e:
            raise ExcelParsingError(f"Erreur lors du parsing du fichier {file_name}: {str(e)}")

    def get_sheet_names(self, file_name: str) -> List[str]:
        """Récupérer les noms des feuilles
        
        Args:
            file_name: Nom du fichier Excel
            
        Returns:
            Liste des noms de feuilles
        """
        workbook = self.load_file(file_name)
        return workbook.sheetnames

    def get_formula_at_cell(self, file_name: str, sheet_name: str, cell_ref: str) -> Optional[str]:
        """Récupérer la formule à une cellule
        
        Args:
            file_name: Nom du fichier Excel
            sheet_name: Nom de la feuille
            cell_ref: Référence de cellule (ex: "B5")
            
        Returns:
            La formule si elle existe, None sinon
        """
        workbook = self.load_file(file_name)
        
        if sheet_name not in workbook.sheetnames:
            return None

        sheet = workbook[sheet_name]
        cell = sheet[cell_ref]
        
        # Retourner la formule si elle existe
        if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
            return cell.value
        
        return None

    def get_all_formulas_in_sheet(self, file_name: str, sheet_name: str) -> Dict[str, str]:
        """Récupérer toutes les formules d'une feuille
        
        Args:
            file_name: Nom du fichier Excel
            sheet_name: Nom de la feuille
            
        Returns:
            Dictionnaire {cellule: formule}
        """
        workbook = self.load_file(file_name)
        
        if sheet_name not in workbook.sheetnames:
            return {}

        sheet = workbook[sheet_name]
        formulas = {}

        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas[cell.coordinate] = cell.value

        return formulas

    def extract_formula_context(self, file_name: str, sheet_name: str, cell_ref: str) -> Optional[ExcelFormula]:
        """Extraire le contexte complet d'une formule
        
        Args:
            file_name: Nom du fichier Excel
            sheet_name: Nom de la feuille
            cell_ref: Référence de cellule
            
        Returns:
            ExcelFormula avec parsing basique
        """
        formula_str = self.get_formula_at_cell(file_name, sheet_name, cell_ref)
        
        if not formula_str:
            return None

        # Créer un objet ExcelFormula basique (parsing plus détaillé ailleurs)
        excel_formula = ExcelFormula(
            original_formula=formula_str,
            file_path=str(self.excel_dir / file_name),
            sheet=sheet_name,
            cell=cell_ref,
            tokens=[],
            references=[],
            functions_used=[],
            is_complex=len(formula_str) > 100,
            has_cross_sheet_refs="!" in formula_str,
            has_external_refs=False,
        )

        logger.info(f"Contexte extrait: {file_name}!{sheet_name}!{cell_ref}")
        return excel_formula

    def get_cell_value(self, file_name: str, sheet_name: str, cell_ref: str) -> Any:
        """Récupérer la valeur calculée d'une cellule
        
        Args:
            file_name: Nom du fichier Excel
            sheet_name: Nom de la feuille
            cell_ref: Référence de cellule
            
        Returns:
            La valeur de la cellule
        """
        workbook = self.load_file(file_name)
        
        if sheet_name not in workbook.sheetnames:
            return None

        sheet = workbook[sheet_name]
        return sheet[cell_ref].value

    def close_file(self, file_name: str) -> None:
        """Fermer un fichier Excel
        
        Args:
            file_name: Nom du fichier
        """
        if file_name in self.workbooks:
            self.workbooks[file_name].close()
            del self.workbooks[file_name]
            logger.info(f"Fichier fermé: {file_name}")

    def close_all(self) -> None:
        """Fermer tous les fichiers ouverts"""
        for file_name in list(self.workbooks.keys()):
            self.close_file(file_name)
