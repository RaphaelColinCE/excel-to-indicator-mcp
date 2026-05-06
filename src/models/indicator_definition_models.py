"""Modèles pour les fichiers de définition d'indicateurs (JSON/XBRL et XML)"""

from pydantic import BaseModel, Field
from typing import List, Optional


# ──────────────────────────────────────────────────────────────────────────────
# Modèle unifié (utilisé en interne par le service)
# ──────────────────────────────────────────────────────────────────────────────

class IndicatorEntry(BaseModel):
    """Représentation normalisée d'un indicateur, indépendante du format source.

    Produit par le parseur JSON (EACIND) ou XML (S0DE, S0PARA, …).
    C'est ce modèle que le reste du code manipule.
    """
    excel_column_code: str          # Code de la colonne Excel (ex: S0DE_SR_YN, EACIND_1A2)
    indicator_name: str             # Nom/Code complet de l'indicateur
    framework: str                  # Framework (ex: S0DE, EACIND)
    label: str = ""                 # Label lisible
    data_type: str = ""             # TEXT, NUMERIC, …
    table_code: str = ""            # Table (ex: SmallAndRisky, f01)
    row_code: str = ""              # Code ligne (optionnel)
    column_code_xbrl: str = ""      # ColumnCode XBRL (ex: C0001)
    formula: Optional[str] = None   # Formule applicative (présente dans XML)
    source_format: str = "json"     # "json" ou "xml"


# ──────────────────────────────────────────────────────────────────────────────
# Modèles bruts JSON (EACIND et similaires)
# ──────────────────────────────────────────────────────────────────────────────

class IndicatorAxis(BaseModel):
    """Axes XBRL d'un datapoint JSON"""
    IndicatorTable: str = ""
    IndicatorZAxis: str = ""
    IndicatorXAxis: str = ""
    IndicatorYAxis: str = ""


class IndicatorDimension(BaseModel):
    DimensionName: str
    DimensionValue: str


class IndicatorDatapoint(BaseModel):
    DPM: str = ""
    SchemaRef: str = ""
    ElementName: str = ""
    Axis: IndicatorAxis = Field(default_factory=IndicatorAxis)
    Dimensions: List[IndicatorDimension] = Field(default_factory=list)
    IsDynamicUnit: bool = False


class IndicatorDefinition(BaseModel):
    """Définition d'un indicateur au format JSON (ex: EACIND.json)"""
    Domain: str = ""
    Framework: str
    Type: str = ""
    Name: str
    Label: str = ""
    Path: str = ""
    TableCode: str = ""
    TableCodeVariant: str = ""
    TableCodeFull: str = ""
    RowCode: str = ""
    ColumnCode: str = ""
    XbrlDataType: str = ""
    DataType: str = ""
    Datapoints: List[IndicatorDatapoint] = Field(default_factory=list)

    def to_entry(self) -> IndicatorEntry:
        """Convertir vers le modèle unifié."""
        # Code colonne : Framework_XAxis (ex: EACIND_1A2)
        excel_col = ""
        if self.Datapoints:
            x_axis = self.Datapoints[0].Axis.IndicatorXAxis
            if x_axis:
                excel_col = f"{self.Framework}_{x_axis}"
        if not excel_col and self.RowCode:
            excel_col = f"{self.Framework}_{self.RowCode}"
        if not excel_col:
            excel_col = self.Name

        return IndicatorEntry(
            excel_column_code=excel_col,
            indicator_name=self.Name,
            framework=self.Framework,
            label=self.Label,
            data_type=self.DataType,
            table_code=self.TableCode,
            row_code=self.RowCode,
            column_code_xbrl=self.ColumnCode,
            formula=None,
            source_format="json",
        )


# ──────────────────────────────────────────────────────────────────────────────
# Modèles bruts XML (S0DE.xml, S0PARA.xml, …)
# ──────────────────────────────────────────────────────────────────────────────

class XmlDatapoint(BaseModel):
    """Datapoint extrait d'un fichier XML"""
    schema_ref: str = ""
    filing_indicator: str = ""
    filing_type: str = ""
    formula: str = ""               # Formule applicative (ex: data[COLUMN=...].VALUE)


class XmlIndicatorDefinition(BaseModel):
    """Définition d'un indicateur au format XML (ex: S0DE.xml)"""
    code: str                       # <Code> complet
    label: str = ""
    publisher: str = ""
    framework: str                  # <Framework><Code>
    table: str = ""                 # <Table><Code>
    row_code: str = ""
    column_code: str                # <ColumnCode> = nom de colonne Excel
    indicator_type: str = ""
    data_type: str = ""
    datapoints: List[XmlDatapoint] = Field(default_factory=list)

    def to_entry(self) -> IndicatorEntry:
        """Convertir vers le modèle unifié."""
        formula = self.datapoints[0].formula if self.datapoints else None
        return IndicatorEntry(
            excel_column_code=self.column_code,
            indicator_name=self.code,
            framework=self.framework,
            label=self.label,
            data_type=self.data_type,
            table_code=self.table,
            row_code=self.row_code,
            column_code_xbrl=self.column_code,
            formula=formula,
            source_format="xml",
        )

