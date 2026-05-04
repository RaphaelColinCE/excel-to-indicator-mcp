"""Modèles de données pour le mapping"""

from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any
from enum import Enum


class ReferenceClassification(str, Enum):
    """Classification des références Excel"""
    STATIC = "static"                      # Valeur constante
    PRECALCULATED = "precalculated"        # Tableau précalculé
    INDICATOR = "indicator"                # Indicateur créé précédemment
    UNRESOLVED = "unresolved"              # Non résoluble


class MappingReference(BaseModel):
    """Référence mappée (Excel → Concept)"""
    excel_reference: str                   # La référence Excel brute
    sheet: str                             # Feuille Excel
    cell_or_range: str                     # Cellule ou range
    classification: ReferenceClassification  # Type de référence
    
    concept: str                           # Concept métier
    indicator: Optional[str] = None        # Indicateur si applicable
    dimensions: List[str] = Field(default_factory=list)  # Dimensions
    
    description: Optional[str] = None
    validation_rules: List[str] = Field(default_factory=list)
    
    class Config:
        use_enum_values = True


class MappedIndicator(BaseModel):
    """Indicateur mappé"""
    name: str                              # Nom unique de l'indicateur
    cell: str                              # Cellule Excel
    sheet: str                             # Feuille Excel
    
    original_formula: str                  # Formule Excel brute
    mapped_formula: str                    # Formule mappée (ex: SUM(Income[Product]))
    
    concept: str                           # Concept métier principal
    indicator_type: str                    # Type d'indicateur
    dimensions: List[str] = Field(default_factory=list)  # Dimensions impliquées
    
    dependencies: List[str] = Field(default_factory=list)  # Autres indicateurs utilisés
    description: Optional[str] = None
    validation_rules: List[str] = Field(default_factory=list)


class MappingContext(BaseModel):
    """Contexte global de mapping"""
    version: str = "1.0"
    description: Optional[str] = None
    
    # Indicateurs par feuille
    sheets: Dict[str, Dict[str, Any]] = Field(default_factory=dict)
    
    # Références globales
    references: Dict[str, MappingReference] = Field(default_factory=dict)
    
    # Équivalences de fonctions Excel
    excel_functions: Dict[str, Dict[str, Any]] = Field(default_factory=dict)
    
    class Config:
        use_enum_values = True


class ConversionContext(BaseModel):
    """Contexte pour une conversion"""
    excel_formula: str                     # Formule Excel brute
    sheet: str                             # Feuille Excel
    cell: str                              # Cellule
    
    available_mappings: List[MappingReference]  # Mappings disponibles
    available_indicators: List[MappedIndicator] # Indicateurs précédents
    
    excel_functions: Dict[str, Dict[str, Any]] = Field(default_factory=dict)  # Équivalences
