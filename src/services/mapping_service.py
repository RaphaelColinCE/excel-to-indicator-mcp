"""Service de gestion du mapping Excel → Concepts applicatifs"""

import json
import logging
from pathlib import Path
from typing import Optional, Dict, List, Any

from src.models.mapping_models import MappingContext, MappingReference, MappedIndicator
from src.utils.exceptions import MappingFileNotFoundError, MappingParsingError
from src.utils.cache import simple_cache

logger = logging.getLogger("mcp_server")


class MappingService:
    """Service pour gérer les mappings Excel → Concepts applicatifs"""

    def __init__(self, mapping_file: Path):
        """Initialiser le service de mapping
        
        Args:
            mapping_file: Chemin du fichier de mapping JSON
        """
        self.mapping_file = Path(mapping_file)
        self.mapping_context: Optional[MappingContext] = None
        self._load_mapping()

    def _load_mapping(self) -> None:
        """Charger le fichier de mapping
        
        Raises:
            MappingFileNotFoundError: Si le fichier n'existe pas
            MappingParsingError: Si le parsing échoue
        """
        if not self.mapping_file.exists():
            logger.warning(f"Fichier de mapping non trouvé: {self.mapping_file}")
            self.mapping_context = MappingContext()
            return

        try:
            with open(self.mapping_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.mapping_context = MappingContext(**data)
                logger.info(f"Mapping chargé: {self.mapping_file}")
        except json.JSONDecodeError as e:
            raise MappingParsingError(f"Erreur JSON dans le mapping: {str(e)}")
        except Exception as e:
            raise MappingParsingError(f"Erreur lors du chargement du mapping: {str(e)}")

    @simple_cache
    def get_reference_mapping(self, excel_ref: str) -> Optional[MappingReference]:
        """Récupérer le mapping d'une référence Excel
        
        Args:
            excel_ref: Référence Excel (ex: "Sheet1!A1", "Revenue", etc)
            
        Returns:
            MappingReference si trouvée, None sinon
        """
        if not self.mapping_context:
            return None

        # Chercher dans les références
        if excel_ref in self.mapping_context.references:
            return self.mapping_context.references[excel_ref]

        # Chercher par pattern (ex: "Sheet1!A1*" pour une plage)
        for ref_key, ref_mapping in self.mapping_context.references.items():
            if excel_ref.startswith(ref_key.rstrip("*")):
                return ref_mapping

        return None

    def get_indicator(self, sheet: str, indicator_name: str) -> Optional[MappedIndicator]:
        """Récupérer un indicateur mappé
        
        Args:
            sheet: Nom de la feuille
            indicator_name: Nom de l'indicateur
            
        Returns:
            MappedIndicator si trouvé, None sinon
        """
        if not self.mapping_context or sheet not in self.mapping_context.sheets:
            return None

        sheet_data = self.mapping_context.sheets[sheet]
        indicators = sheet_data.get("indicators", [])

        for ind_data in indicators:
            if ind_data.get("name") == indicator_name:
                return MappedIndicator(**ind_data)

        return None

    def get_all_indicators(self, sheet: Optional[str] = None) -> List[MappedIndicator]:
        """Récupérer tous les indicateurs
        
        Args:
            sheet: Feuille spécifique (optionnel)
            
        Returns:
            Liste des indicateurs
        """
        if not self.mapping_context:
            return []

        indicators = []
        
        sheets_to_process = (
            [sheet] if sheet else self.mapping_context.sheets.keys()
        )

        for sheet_name in sheets_to_process:
            if sheet_name in self.mapping_context.sheets:
                sheet_data = self.mapping_context.sheets[sheet_name]
                for ind_data in sheet_data.get("indicators", []):
                    indicators.append(MappedIndicator(**ind_data))

        return indicators

    def get_function_mapping(self, excel_func: str) -> Optional[Dict[str, Any]]:
        """Récupérer le mapping d'une fonction Excel
        
        Args:
            excel_func: Fonction Excel (ex: "SUM", "VLOOKUP")
            
        Returns:
            Mapping de la fonction si existant
        """
        if not self.mapping_context:
            return None

        return self.mapping_context.excel_functions.get(excel_func.upper())

    def search_references(self, keyword: str) -> List[MappingReference]:
        """Rechercher des références par mot-clé
        
        Args:
            keyword: Mot-clé à chercher
            
        Returns:
            Liste des références correspondantes
        """
        if not self.mapping_context:
            return []

        results = []
        keyword_lower = keyword.lower()

        for ref_mapping in self.mapping_context.references.values():
            if (
                keyword_lower in ref_mapping.excel_reference.lower()
                or keyword_lower in ref_mapping.concept.lower()
                or (ref_mapping.description and keyword_lower in ref_mapping.description.lower())
            ):
                results.append(ref_mapping)

        return results

    def search_indicators(self, keyword: str) -> List[MappedIndicator]:
        """Rechercher des indicateurs par mot-clé
        
        Args:
            keyword: Mot-clé à chercher
            
        Returns:
            Liste des indicateurs correspondants
        """
        indicators = self.get_all_indicators()
        keyword_lower = keyword.lower()

        return [
            ind for ind in indicators
            if (
                keyword_lower in ind.name.lower()
                or keyword_lower in ind.concept.lower()
                or (ind.description and keyword_lower in ind.description.lower())
            )
        ]

    def save_mapping(self, output_path: Optional[Path] = None) -> None:
        """Sauvegarder le mapping en JSON
        
        Args:
            output_path: Chemin de sortie (défaut: fichier original)
        """
        if not self.mapping_context:
            logger.warning("Pas de contexte de mapping à sauvegarder")
            return

        path = Path(output_path or self.mapping_file)
        path.parent.mkdir(parents=True, exist_ok=True)

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(
                    self.mapping_context.model_dump(exclude_none=True),
                    f,
                    indent=2,
                    ensure_ascii=False,
                )
            logger.info(f"Mapping sauvegardé: {path}")
        except Exception as e:
            logger.error(f"Erreur lors de la sauvegarde du mapping: {str(e)}")
            raise
