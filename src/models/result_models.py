"""Modèles de résultats"""

from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any
from enum import Enum
from datetime import datetime


class ConversionStatus(str, Enum):
    """Statut de conversion"""
    RESOLVED = "resolved"                          # Conversion complète
    PARTIALLY_RESOLVED = "partially_resolved"      # Partie non mappée
    UNRESOLVED = "unresolved"                      # Impossible de convertir
    ERROR = "error"                                # Erreur pendant la conversion


class MappingUsage(BaseModel):
    """Trace d'utilisation d'un mapping"""
    excel_reference: str
    concept: str
    indicator: Optional[str] = None
    dimensions: List[str] = Field(default_factory=list)
    mapped_to: str                                  # Ce vers quoi il a été mappé


class ConversionResult(BaseModel):
    """Résultat de conversion Excel → Indicator"""
    status: ConversionStatus
    
    original_formula: str
    mapped_formula: Optional[str] = None
    explanation: str
    
    # Détails de la conversion
    mapping_used: List[MappingUsage] = Field(default_factory=list)
    unresolved_references: List[str] = Field(default_factory=list)
    
    # Métadonnées
    excel_file: Optional[str] = None
    sheet: Optional[str] = None
    cell: Optional[str] = None
    
    errors: List[str] = Field(default_factory=list)
    warnings: List[str] = Field(default_factory=list)
    
    # Timing
    conversion_timestamp: datetime = Field(default_factory=datetime.now)
    
    class Config:
        use_enum_values = True


class ValidationError(BaseModel):
    """Erreur de validation"""
    type: str                                      # Type d'erreur
    message: str                                   # Message
    severity: str = "error"                        # error, warning, info
    location: Optional[str] = None                 # Où l'erreur survient


class ValidationReport(BaseModel):
    """Rapport de validation"""
    formula: str
    is_valid: bool
    
    errors: List[ValidationError] = Field(default_factory=list)
    warnings: List[ValidationError] = Field(default_factory=list)
    
    # Informations de validation
    validation_timestamp: datetime = Field(default_factory=datetime.now)
    validation_rules_applied: List[str] = Field(default_factory=list)
    
    # Métriques
    complexity_score: float = 0.0          # 0.0 à 1.0
    readability_score: float = 0.0         # 0.0 à 1.0
    confidence_score: float = 0.0          # 0.0 à 1.0


class DependencyNode(BaseModel):
    """Nœud dans l'arbre de dépendances"""
    reference: str
    sheet: Optional[str] = None
    cell: Optional[str] = None
    type: str                              # cell, range, indicator, constant
    value: Optional[str] = None            # Valeur si c'est une constante
    children: List['DependencyNode'] = Field(default_factory=list)
    depth: int = 0


class DependencyTrace(BaseModel):
    """Trace complète des dépendances"""
    original_formula: str
    root_node: DependencyNode
    
    total_depth: int
    total_references: int
    circular_dependencies: List[str] = Field(default_factory=list)
    
    # Informations utiles
    precalculated_ranges: List[str] = Field(default_factory=list)
    external_references: List[str] = Field(default_factory=list)
    unresolvable_references: List[str] = Field(default_factory=list)
