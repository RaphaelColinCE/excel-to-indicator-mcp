"""Modèles de données pour Excel"""

from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any
from enum import Enum


class ReferenceType(str, Enum):
    """Types de références dans une formule"""
    CELL = "cell"                          # A1
    RANGE = "range"                        # A1:B10
    NAMED_RANGE = "named_range"            # Revenue
    EXTERNAL_CELL = "external_cell"        # Sheet1!A1
    EXTERNAL_RANGE = "external_range"      # Sheet1!A1:B10
    FUNCTION = "function"                  # SUM, VLOOKUP, etc
    CONSTANT = "constant"                  # 0.5, "text", etc


class ExcelReference(BaseModel):
    """Représente une référence Excel"""
    type: ReferenceType
    value: str                             # La référence brute (A1, Sheet1!A1, etc)
    sheet: Optional[str] = None            # Feuille si applicable
    start_cell: Optional[str] = None       # Cellule de début (pour les ranges)
    end_cell: Optional[str] = None         # Cellule de fin (pour les ranges)
    is_relative: bool = True               # Référence relative ou absolue
    description: Optional[str] = None      # Description optionnelle

    class Config:
        use_enum_values = True


class FormulaToken(BaseModel):
    """Représente un token dans une formule"""
    type: str                              # "operator", "reference", "function", "constant"
    value: str                             # La valeur du token
    position: int                          # Position dans la formule originale


class ExcelFormula(BaseModel):
    """Représente une formule Excel parsée"""
    original_formula: str                  # Formule brute (ex: =SUM(A1:A10))
    file_path: Optional[str] = None        # Chemin du fichier Excel
    sheet: str                             # Nom de la feuille
    cell: str                              # Cellule contenant la formule (ex: B5)
    
    # Éléments parsés
    tokens: List[FormulaToken]             # Tokens de la formule
    references: List[ExcelReference]       # Références trouvées
    functions_used: List[str]              # Fonctions utilisées (SUM, VLOOKUP, etc)
    
    # Métadonnées
    is_complex: bool                       # Formule complexe ?
    has_cross_sheet_refs: bool             # Contient des références entre feuilles ?
    has_external_refs: bool                # Contient des références externes ?
    depth_estimate: int = 1                # Estimation de la profondeur de dépendances
    
    # Annotations optionnelles
    name: Optional[str] = None             # Nom de l'indicateur si applicable
    description: Optional[str] = None      # Description de la formule

    class Config:
        use_enum_values = True


class ExcelContext(BaseModel):
    """Contexte Excel pour la résolution"""
    current_sheet: str                     # Feuille actuelle
    current_cell: str                      # Cellule actuelle
    available_sheets: List[str]            # Feuilles disponibles
    named_ranges: Dict[str, str]           # Ranges nommés {nom: référence}
    cell_values: Dict[str, Any] = Field(default_factory=dict)  # Valeurs connues
