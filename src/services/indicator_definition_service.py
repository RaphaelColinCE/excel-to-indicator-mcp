"""Service de chargement et d'indexation des définitions d'indicateurs XBRL"""

import json
import logging
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional

from src.models.indicator_definition_models import (
    IndicatorDefinition,
    IndicatorEntry,
    XmlDatapoint,
    XmlIndicatorDefinition,
)

logger = logging.getLogger("mcp_server")


class IndicatorDefinitionService:
    """Charge les fichiers de définitions d'indicateurs (JSON et XML) et construit
    un index : code colonne Excel → IndicatorEntry.

    Convention de nommage des fichiers :
      - EACIND.json  → indicateurs du framework EACIND (format JSON)
      - S0DE.xml     → indicateurs du framework S0DE   (format XML)
      - S0PARA.xml   → indicateurs du framework S0PARA (format XML)
      - etc.

    La correspondance colonne Excel → fichier est déterminée par le préfixe :
      S0DE_SR_YN  → framework S0DE → S0DE.xml
      EACIND_1A2  → framework EACIND → EACIND.json
    """

    def __init__(self, indicators_dir: Path):
        self.indicators_dir = Path(indicators_dir)
        self._by_column_code: Dict[str, IndicatorEntry] = {}
        self._by_name: Dict[str, IndicatorEntry] = {}
        self._by_framework: Dict[str, List[IndicatorEntry]] = {}
        self._loaded_files: List[str] = []
        self._load_all()

    # ──────────────────────────────────────────────────────────────────────────
    # Chargement
    # ──────────────────────────────────────────────────────────────────────────

    def _load_all(self) -> None:
        if not self.indicators_dir.exists():
            logger.warning(f"Répertoire d'indicateurs introuvable: {self.indicators_dir}")
            return

        json_files = list(self.indicators_dir.rglob("*.json"))
        xml_files = list(self.indicators_dir.rglob("*.xml"))

        if not json_files and not xml_files:
            logger.warning(f"Aucun fichier JSON/XML trouvé dans: {self.indicators_dir}")
            return

        for f in json_files:
            self._load_json(f)
        for f in xml_files:
            self._load_xml(f)

        logger.info(
            f"Indicateurs chargés: {len(self._by_column_code)} entrées "
            f"depuis {len(self._loaded_files)} fichier(s)"
        )

    def _load_json(self, file_path: Path) -> None:
        """Parser un fichier JSON (format EACIND)."""
        try:
            with open(file_path, encoding="utf-8") as f:
                raw = json.load(f)
        except Exception as e:
            logger.error(f"Impossible de lire {file_path}: {e}")
            return

        # Normaliser en liste
        if isinstance(raw, dict):
            list_value = next((v for v in raw.values() if isinstance(v, list)), None)
            items = list_value if list_value is not None else [raw]
        elif isinstance(raw, list):
            items = raw
        else:
            logger.warning(f"Format JSON inattendu dans {file_path}")
            return

        count = 0
        for item in items:
            if not isinstance(item, dict) or "Name" not in item:
                continue
            try:
                indicator = IndicatorDefinition(**item)
                self._index(indicator.to_entry())
                count += 1
            except Exception as e:
                logger.debug(f"Indicateur JSON ignoré dans {file_path.name}: {e}")

        if count:
            self._loaded_files.append(file_path.name)
            logger.info(f"  {file_path.name}: {count} indicateur(s) JSON chargé(s)")

    def _load_xml(self, file_path: Path) -> None:
        """Parser un fichier XML (format S0DE, S0PARA, …).

        Structure attendue :
            <Indicators>
              <Indicator>
                <Code>…</Code>
                <Framework><Code>S0DE</Code></Framework>
                <Table><Code>SmallAndRisky</Code></Table>
                <ColumnCode>S0DE_SR_YN</ColumnCode>
                <Datapoints>
                  <Datapoint SchemaRef="V1" …>
                    <Formula>data[…].VALUE</Formula>
                  </Datapoint>
                </Datapoints>
              </Indicator>
            </Indicators>
        """
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
        except ET.ParseError as e:
            logger.error(f"Erreur XML dans {file_path.name}: {e}")
            return
        except Exception as e:
            logger.error(f"Impossible de lire {file_path.name}: {e}")
            return

        # Le root peut être <Indicators> ou directement <Indicator>
        indicators_nodes = (
            list(root.iter("Indicator"))
            if root.tag != "Indicator"
            else [root]
        )

        count = 0
        for node in indicators_nodes:
            try:
                xml_ind = _parse_xml_indicator(node)
                if xml_ind:
                    self._index(xml_ind.to_entry())
                    count += 1
            except Exception as e:
                logger.debug(f"Indicateur XML ignoré dans {file_path.name}: {e}")

        if count:
            self._loaded_files.append(file_path.name)
            logger.info(f"  {file_path.name}: {count} indicateur(s) XML chargé(s)")

    def _index(self, entry: IndicatorEntry) -> None:
        col = entry.excel_column_code
        if col and col not in self._by_column_code:
            self._by_column_code[col] = entry

        name = entry.indicator_name
        if name and name not in self._by_name:
            self._by_name[name] = entry

        fw = entry.framework
        if fw:
            self._by_framework.setdefault(fw, []).append(entry)

    # ──────────────────────────────────────────────────────────────────────────
    # API de consultation
    # ──────────────────────────────────────────────────────────────────────────

    def resolve_column_code(self, column_code: str) -> Optional[IndicatorEntry]:
        """Résoudre un code de colonne Excel vers son IndicatorEntry.

        Args:
            column_code: Ex: "S0DE_SR_YN" ou "EACIND_1A2"
        """
        return self._by_column_code.get(column_code)

    def get_indicator_name(self, column_code: str) -> Optional[str]:
        """Retourner le nom/code complet depuis un code de colonne Excel."""
        entry = self.resolve_column_code(column_code)
        return entry.indicator_name if entry else None

    def get_formula(self, column_code: str) -> Optional[str]:
        """Retourner la formule applicative depuis un code de colonne Excel.

        Disponible uniquement pour les indicateurs XML.
        """
        entry = self.resolve_column_code(column_code)
        return entry.formula if entry else None

    def get_by_name(self, name: str) -> Optional[IndicatorEntry]:
        return self._by_name.get(name)

    def get_all(self, framework: Optional[str] = None) -> List[IndicatorEntry]:
        if framework:
            return self._by_framework.get(framework, [])
        return list(self._by_name.values())

    def search(self, keyword: str) -> List[IndicatorEntry]:
        kw = keyword.lower()
        results: List[IndicatorEntry] = []
        seen: set = set()
        for entry in self._by_name.values():
            if (
                kw in entry.indicator_name.lower()
                or kw in entry.label.lower()
                or kw in entry.excel_column_code.lower()
            ):
                if entry.indicator_name not in seen:
                    seen.add(entry.indicator_name)
                    results.append(entry)
        return results

    def get_loaded_files(self) -> List[str]:
        return list(self._loaded_files)

    def get_frameworks(self) -> List[str]:
        return sorted(self._by_framework.keys())

    @property
    def total_count(self) -> int:
        return len(self._by_name)


# ──────────────────────────────────────────────────────────────────────────────
# Fonctions utilitaires XML
# ──────────────────────────────────────────────────────────────────────────────

def _text(node: ET.Element, tag: str, default: str = "") -> str:
    """Extraire le texte d'un sous-élément XML (retourne default si absent)."""
    child = node.find(tag)
    if child is not None and child.text:
        return child.text.strip()
    return default


def _parse_xml_indicator(node: ET.Element) -> Optional[XmlIndicatorDefinition]:
    """Parser un nœud XML <Indicator> → XmlIndicatorDefinition."""
    code = _text(node, "Code")
    column_code = _text(node, "ColumnCode")

    if not code or not column_code:
        return None

    framework = _text(node, "Framework/Code")
    table = _text(node, "Table/Code")

    datapoints: List[XmlDatapoint] = []
    dp_container = node.find("Datapoints")
    if dp_container is not None:
        for dp_node in dp_container.findall("Datapoint"):
            datapoints.append(
                XmlDatapoint(
                    schema_ref=dp_node.get("SchemaRef", ""),
                    filing_indicator=dp_node.get("FilingIndicator", ""),
                    filing_type=dp_node.get("FilingType", ""),
                    formula=_text(dp_node, "Formula"),
                )
            )

    return XmlIndicatorDefinition(
        code=code,
        label=_text(node, "Label"),
        publisher=_text(node, "Publisher"),
        framework=framework,
        table=table,
        row_code=_text(node, "RowCode"),
        column_code=column_code,
        indicator_type=_text(node, "IndicatorType"),
        data_type=_text(node, "DataType"),
        datapoints=datapoints,
    )

