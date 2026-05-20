"""Serveur MCP Excel-to-Indicator

Point d'entrée principal. Expose les outils MCP pour la conversion de formules
Excel en formules d'indicateurs métier.

Usage:
    python main.py
"""

import asyncio
import json
import logging
from pathlib import Path
from typing import Any

import mcp.server.stdio
import mcp.types as types
from mcp.server import NotificationOptions, Server
from mcp.server.models import InitializationOptions

from src.config.logging import setup_logging
from src.config.settings import settings
from src.services.conversion_service import ConversionService
from src.services.excel_converter import ExcelConverterService
from src.services.excel_service import ExcelService
from src.services.formula_parser import parse_formula
from src.services.indicator_definition_service import IndicatorDefinitionService
from src.services.indicator_export_service import IndicatorExportService
from src.services.mapping_service import MappingService

# ── Initialisation ────────────────────────────────────────────────────────────

setup_logging()
logger = logging.getLogger("mcp_server")

settings.ensure_directories()

_excel_service = ExcelService(settings.EXCEL_SOURCES_DIR)
_mapping_service = MappingService(settings.MAPPING_FILE)
_indicator_service = IndicatorDefinitionService(settings.INDICATORS_DIR)
_conversion_service = ConversionService(_mapping_service, settings.MAX_DEPTH_RESOLUTION)
_excel_converter = ExcelConverterService(settings.MAPPING_FILE)
_indicator_exporter = IndicatorExportService(
    details_path=settings.MAPPING_DIR / "IndicatorsDetails.xlsx",
    mapping_path=settings.MAPPING_FILE,
)

server = Server("excel-to-indicator")

# ── Liste des outils ──────────────────────────────────────────────────────────


@server.list_tools()
async def handle_list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="list_excel_files",
            description="Lister tous les fichiers Excel disponibles dans le répertoire de sources.",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": [],
            },
        ),
        types.Tool(
            name="get_sheet_names",
            description="Récupérer les noms des feuilles d'un fichier Excel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_name": {
                        "type": "string",
                        "description": "Nom du fichier Excel (ex: data.xlsx)",
                    }
                },
                "required": ["file_name"],
            },
        ),
        types.Tool(
            name="get_formula",
            description="Récupérer la formule présente dans une cellule Excel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_name": {"type": "string", "description": "Nom du fichier Excel"},
                    "sheet_name": {"type": "string", "description": "Nom de la feuille"},
                    "cell_ref": {
                        "type": "string",
                        "description": "Référence de cellule (ex: B5)",
                    },
                },
                "required": ["file_name", "sheet_name", "cell_ref"],
            },
        ),
        types.Tool(
            name="get_all_formulas",
            description="Récupérer toutes les formules d'une feuille Excel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_name": {"type": "string", "description": "Nom du fichier Excel"},
                    "sheet_name": {"type": "string", "description": "Nom de la feuille"},
                },
                "required": ["file_name", "sheet_name"],
            },
        ),
        types.Tool(
            name="parse_formula",
            description=(
                "Parser une formule Excel pour en extraire les références, "
                "fonctions et tokens sans effectuer la conversion."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "formula": {"type": "string", "description": "Formule Excel brute"},
                    "sheet": {
                        "type": "string",
                        "description": "Feuille contenant la formule",
                    },
                    "cell": {
                        "type": "string",
                        "description": "Cellule contenant la formule",
                    },
                },
                "required": ["formula", "sheet", "cell"],
            },
        ),
        types.Tool(
            name="convert_formula",
            description=(
                "Convertir une formule Excel en formule d'indicateur métier "
                "en utilisant le mapping Excel ↔ Concepts."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "formula": {"type": "string", "description": "Formule Excel brute"},
                    "sheet": {
                        "type": "string",
                        "description": "Feuille contenant la formule",
                    },
                    "cell": {
                        "type": "string",
                        "description": "Cellule contenant la formule",
                    },
                    "file_name": {
                        "type": "string",
                        "description": "Nom du fichier Excel (optionnel)",
                    },
                },
                "required": ["formula", "sheet", "cell"],
            },
        ),
        types.Tool(
            name="validate_formula",
            description="Valider une formule d'indicateur (vérification syntaxique et cohérence).",
            inputSchema={
                "type": "object",
                "properties": {
                    "formula": {
                        "type": "string",
                        "description": "Formule d'indicateur à valider",
                    }
                },
                "required": ["formula"],
            },
        ),
        types.Tool(
            name="trace_dependencies",
            description="Tracer l'arbre de dépendances d'une formule Excel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "formula": {"type": "string", "description": "Formule Excel brute"},
                    "sheet": {"type": "string", "description": "Feuille contenant la formule"},
                    "cell": {"type": "string", "description": "Cellule contenant la formule"},
                    "file_name": {
                        "type": "string",
                        "description": "Nom du fichier Excel (optionnel)",
                    },
                },
                "required": ["formula", "sheet", "cell"],
            },
        ),
        types.Tool(
            name="search_mappings",
            description="Rechercher des références ou concepts dans le fichier de mapping.",
            inputSchema={
                "type": "object",
                "properties": {
                    "keyword": {
                        "type": "string",
                        "description": "Mot-clé à rechercher dans les références et concepts",
                    }
                },
                "required": ["keyword"],
            },
        ),
        types.Tool(
            name="get_indicators",
            description="Récupérer tous les indicateurs mappés, avec filtre optionnel par feuille.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet": {
                        "type": "string",
                        "description": "Nom de la feuille (optionnel, tous si absent)",
                    }
                },
                "required": [],
            },
        ),
        types.Tool(
            name="convert_cell",
            description=(
                "Lire la formule d'une cellule Excel depuis un fichier, "
                "puis la convertir en formule d'indicateur en une seule opération."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_name": {"type": "string", "description": "Nom du fichier Excel"},
                    "sheet_name": {"type": "string", "description": "Nom de la feuille"},
                    "cell_ref": {
                        "type": "string",
                        "description": "Référence de cellule (ex: B5)",
                    },
                },
                "required": ["file_name", "sheet_name", "cell_ref"],
            },
        ),
        types.Tool(
            name="list_indicator_files",
            description="Lister les fichiers de définitions d'indicateurs chargés et les frameworks disponibles.",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        types.Tool(
            name="resolve_column_code",
            description=(
                "Résoudre un code de colonne Excel (ex: EACIND_1A2) vers "
                "le IndicatorName complet (ex: EACIND_f01_R01A2_C0001) "
                "et les métadonnées de l'indicateur."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "column_code": {
                        "type": "string",
                        "description": "Code de colonne Excel (ex: EACIND_1A2)",
                    }
                },
                "required": ["column_code"],
            },
        ),
        types.Tool(
            name="search_indicator_definitions",
            description="Rechercher des indicateurs dans les définitions chargées (par Name, Label ou code colonne).",
            inputSchema={
                "type": "object",
                "properties": {
                    "keyword": {"type": "string", "description": "Mot-clé à rechercher"},
                    "framework": {
                        "type": "string",
                        "description": "Filtrer par framework (ex: EACIND) — optionnel",
                    },
                },
                "required": ["keyword"],
            },
        ),
        types.Tool(
            name="convert_excel_sheet",
            description=(
                "Convertir toutes les formules d'une feuille Excel en formules d'indicateurs. "
                "Lit la ligne d'en-tête (header_row, défaut 25) pour identifier les codes de colonnes, "
                "puis convertit les cellules de formula_row (défaut 26) dont l'en-tête n'est pas encore "
                "dans le mapping (= nouveaux indicateurs à définir). "
                "Retourne les enregistrements (source_excel_sheet, source_cell, source_formula, "
                "indicator_name, indicator_formula)."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_name": {
                        "type": "string",
                        "description": "Nom du fichier Excel dans le répertoire de sources (ex: MySheet.xlsx)",
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Nom de la feuille à traiter (ex: S1CRU1)",
                    },
                    "header_row": {
                        "type": "integer",
                        "description": "Numéro de la ligne contenant les codes de colonnes (défaut: 25)",
                        "default": 25,
                    },
                    "formula_row": {
                        "type": "integer",
                        "description": "Numéro de la ligne contenant les formules (défaut: 26)",
                        "default": 26,
                    },
                    "output_columns": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Lettres de colonnes à forcer comme sorties (ex: [\"M\",\"N\"]). "
                                       "Si absent, détection automatique (en-tête pas dans mapping).",
                    },
                },
                "required": ["file_name", "sheet_name"],
            },
        ),
        types.Tool(
            name="export_indicators_to_excel",
            description=(
                "Exporter les résultats d'une conversion en fichier Excel "
                "au format attendu par l'application (colonnes: source_excel_sheet, "
                "source_cell, source_formula, indicator_name, indicator_formula)."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_name": {
                        "type": "string",
                        "description": "Fichier Excel source",
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Feuille à traiter",
                    },
                    "output_file_name": {
                        "type": "string",
                        "description": "Nom du fichier de sortie (ex: S1CRU1_output.xlsx). "
                                       "Sera créé dans le répertoire de sources.",
                    },
                    "header_row": {
                        "type": "integer",
                        "description": "Ligne d'en-tête (défaut: 25)",
                        "default": 25,
                    },
                    "formula_row": {
                        "type": "integer",
                        "description": "Ligne de formules (défaut: 26)",
                        "default": 26,
                    },
                    "multi_row": {
                        "type": "boolean",
                        "description": "Si true, utilise le mode multi-lignes (DATA/CELL alternées) pour des feuilles comme S1CRUF. "
                                       "Dans ce mode, header_row et formula_row sont ignorés.",
                        "default": False,
                    },
                },
                "required": ["file_name", "sheet_name", "output_file_name"],
            },
        ),
        types.Tool(
            name="convert_multi_row_sheet",
            description=(
                "Convertir une feuille Excel à structure multi-lignes (DATA/CELL alternées), "
                "comme S1CRUF. Chaque ligne #DATA est suivie d'une ligne #CELL définissant les "
                "indicateurs. Construit la carte globale de toutes les lignes #CELL, puis convertit "
                "toutes les cellules formule dont l'indicateur n'est pas encore dans le mapping "
                "(= nouveaux indicateurs à définir). Les références non résolues (plages sans "
                "indicateur nommé) sont laissées telles quelles. "
                "Retourne les enregistrements (source_excel_sheet, source_cell, source_formula, "
                "indicator_name, indicator_formula)."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_name": {
                        "type": "string",
                        "description": "Nom du fichier Excel dans le répertoire de sources (ex: S1CRUF_input.xlsx)",
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Nom de la feuille à traiter (ex: S1CRUF)",
                    },
                },
                "required": ["file_name", "sheet_name"],
            },
        ),
        types.Tool(
            name="export_indicators_to_xml",
            description=(
                "Convertir une feuille Excel et exporter les indicateurs au format XML "
                "importable dans l'application ISS. "
                "Génère un fichier XML avec <CustomIndicators> contenant Code, Label, "
                "DataType (depuis IndicatorsDetails.xlsx) et Formula pour chaque indicateur. "
                "Supporte les feuilles standard (header_row/formula_row) et multi-lignes (multi_row=true)."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_name": {"type": "string", "description": "Fichier Excel source"},
                    "sheet_name": {"type": "string", "description": "Feuille à traiter"},
                    "output_file_name": {
                        "type": "string",
                        "description": "Nom du fichier XML de sortie (ex: S1CRU1_Indicators.xml)",
                    },
                    "header_row": {
                        "type": "integer",
                        "description": "Ligne d'en-tête pour format standard (défaut: 25)",
                        "default": 25,
                    },
                    "formula_row": {
                        "type": "integer",
                        "description": "Ligne de formules pour format standard (défaut: 26)",
                        "default": 26,
                    },
                    "multi_row": {
                        "type": "boolean",
                        "description": "Si true, utilise le mode multi-lignes (DATA/CELL alternées). "
                                       "header_row et formula_row sont ignorés dans ce mode.",
                        "default": False,
                    },
                    "data_types": {
                        "type": "object",
                        "description": "Surcharge optionnelle des DataType par indicateur "
                                       "(ex: {\"S1CRU1_LUMPS\": \"MONETARY\"}). "
                                       "Prioritaire sur IndicatorsDetails.xlsx. Valeurs: MONETARY, REAL, DATE, TEXT.",
                        "additionalProperties": {"type": "string"},
                    },
                },
                "required": ["file_name", "sheet_name", "output_file_name"],
            },
        ),
        types.Tool(
            name="export_analysis_table_spec",
            description=(
                "Générer le fichier de spécification Excel pour le tableau d'analyse ISS. "
                "Lit la structure du fichier Excel source (lignes statiques + section dynamique) "
                "et produit un fichier .xlsx avec deux feuilles : #TEMPLATE et Data. "
                "La feuille Data contient les lignes #DATA/#CELL avec les types "
                "Indicator[XBRL], Indicator[RS], Indicator[CUSTOM] et Dimension:SRF_LEI. "
                "Supporte uniquement le format standard (header_row/formula_row)."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "file_name": {"type": "string", "description": "Fichier Excel source (input ISS)"},
                    "sheet_name": {"type": "string", "description": "Nom de la feuille à traiter"},
                    "output_file_name": {
                        "type": "string",
                        "description": "Nom du fichier de spécification de sortie (ex: S1CRU1_Spec.xlsx)",
                    },
                    "header_row": {
                        "type": "integer",
                        "description": "Ligne des codes courts (défaut: 25)",
                        "default": 25,
                    },
                    "formula_row": {
                        "type": "integer",
                        "description": "Ligne des formules / données dynamiques (défaut: 26)",
                        "default": 26,
                    },
                },
                "required": ["file_name", "sheet_name", "output_file_name"],
            },
        ),
    ]


# ── Implémentation des outils ─────────────────────────────────────────────────


@server.call_tool()
async def handle_call_tool(
    name: str, arguments: dict[str, Any] | None
) -> list[types.TextContent]:
    args = arguments or {}

    try:
        if name == "list_excel_files":
            files = settings.get_excel_files()
            result = {
                "excel_files": [f.name for f in files],
                "count": len(files),
                "directory": str(settings.EXCEL_SOURCES_DIR),
            }

        elif name == "get_sheet_names":
            sheets = _excel_service.get_sheet_names(args["file_name"])
            result = {"file_name": args["file_name"], "sheets": sheets}

        elif name == "get_formula":
            formula = _excel_service.get_formula_at_cell(
                args["file_name"], args["sheet_name"], args["cell_ref"]
            )
            result = {
                "file_name": args["file_name"],
                "sheet": args["sheet_name"],
                "cell": args["cell_ref"],
                "formula": formula,
                "has_formula": formula is not None,
            }

        elif name == "get_all_formulas":
            formulas = _excel_service.get_all_formulas_in_sheet(
                args["file_name"], args["sheet_name"]
            )
            result = {
                "file_name": args["file_name"],
                "sheet": args["sheet_name"],
                "formulas": formulas,
                "count": len(formulas),
            }

        elif name == "parse_formula":
            parsed = parse_formula(
                args["formula"], args["sheet"], args["cell"]
            )
            result = parsed.model_dump()

        elif name == "convert_formula":
            conversion = _conversion_service.convert_formula(
                formula=args["formula"],
                sheet=args["sheet"],
                cell=args["cell"],
                file_path=args.get("file_name", ""),
            )
            result = conversion.model_dump()

        elif name == "validate_formula":
            report = _conversion_service.validate_formula(args["formula"])
            result = report.model_dump()

        elif name == "trace_dependencies":
            trace = _conversion_service.trace_dependencies(
                formula=args["formula"],
                sheet=args["sheet"],
                cell=args["cell"],
                file_path=args.get("file_name", ""),
            )
            result = trace.model_dump()

        elif name == "search_mappings":
            refs = _mapping_service.search_references(args["keyword"])
            indicators = _mapping_service.search_indicators(args["keyword"])
            result = {
                "keyword": args["keyword"],
                "references": [r.model_dump() for r in refs],
                "indicators": [i.model_dump() for i in indicators],
                "total_found": len(refs) + len(indicators),
            }

        elif name == "get_indicators":
            sheet = args.get("sheet")
            indicators = _mapping_service.get_all_indicators(sheet)
            result = {
                "sheet_filter": sheet,
                "indicators": [i.model_dump() for i in indicators],
                "count": len(indicators),
            }

        elif name == "convert_cell":
            formula = _excel_service.get_formula_at_cell(
                args["file_name"], args["sheet_name"], args["cell_ref"]
            )
            if formula is None:
                result = {
                    "error": f"Aucune formule trouvée dans {args['sheet_name']}!{args['cell_ref']}",
                    "file_name": args["file_name"],
                }
            else:
                conversion = _conversion_service.convert_formula(
                    formula=formula,
                    sheet=args["sheet_name"],
                    cell=args["cell_ref"],
                    file_path=args["file_name"],
                )
                result = conversion.model_dump()

        elif name == "list_indicator_files":
            result = {
                "loaded_files": _indicator_service.get_loaded_files(),
                "frameworks": _indicator_service.get_frameworks(),
                "total_indicators": _indicator_service.total_count,
                "indicators_dir": str(settings.INDICATORS_DIR),
            }

        elif name == "resolve_column_code":
            indicator = _indicator_service.resolve_column_code(args["column_code"])
            if indicator:
                result = {
                    "column_code": args["column_code"],
                    "found": True,
                    "indicator_name": indicator.indicator_name,
                    "label": indicator.label,
                    "framework": indicator.framework,
                    "data_type": indicator.data_type,
                    "table_code": indicator.table_code,
                    "row_code": indicator.row_code,
                    "formula": indicator.formula,
                    "source_format": indicator.source_format,
                }
            else:
                result = {
                    "column_code": args["column_code"],
                    "found": False,
                    "indicator_name": None,
                }

        elif name == "search_indicator_definitions":
            matches = _indicator_service.search(args["keyword"])
            if args.get("framework"):
                matches = [m for m in matches if m.Framework == args["framework"]]
            result = {
                "keyword": args["keyword"],
                "framework_filter": args.get("framework"),
                "count": len(matches),
                "results": [
                    {
                        "column_code": m.excel_column_code,
                        "indicator_name": m.indicator_name,
                        "label": m.label,
                        "framework": m.framework,
                        "data_type": m.data_type,
                        "formula": m.formula,
                        "source_format": m.source_format,
                    }
                    for m in matches[:50]
                ],
            }

        elif name == "convert_excel_sheet":
            file_path = settings.EXCEL_SOURCES_DIR / args["file_name"]
            if not file_path.exists():
                # Chercher aussi dans les exemples
                alt = settings.PROJECT_ROOT / "data" / "examples" / args["file_name"]
                if alt.exists():
                    file_path = alt
                else:
                    raise FileNotFoundError(f"Fichier introuvable: {args['file_name']}")

            # Mettre à jour le registre des dimensions créées puis rafraîchir la conversion.
            _indicator_exporter.register_created_dimensions(file_path, args["sheet_name"])
            _excel_converter.refresh_created_dimensions()

            records = _excel_converter.convert_sheet(
                file_path=file_path,
                sheet_name=args["sheet_name"],
                header_row=args.get("header_row", 25),
                formula_row=args.get("formula_row", 26),
                output_columns=args.get("output_columns"),
            )
            result = {
                "file_name": args["file_name"],
                "sheet_name": args["sheet_name"],
                "count": len(records),
                "indicators": records,
                "known_created_dimensions": _indicator_exporter.get_created_dimensions_registry(),
            }

        elif name == "export_indicators_to_excel":
            import openpyxl as _openpyxl
            file_path = settings.EXCEL_SOURCES_DIR / args["file_name"]
            if not file_path.exists():
                alt = settings.PROJECT_ROOT / "data" / "examples" / args["file_name"]
                if alt.exists():
                    file_path = alt
                else:
                    raise FileNotFoundError(f"Fichier introuvable: {args['file_name']}")

            _indicator_exporter.register_created_dimensions(file_path, args["sheet_name"])
            _excel_converter.refresh_created_dimensions()

            multi_row = args.get("multi_row", False)
            if multi_row:
                records = _excel_converter.convert_multi_row_sheet(
                    file_path=file_path,
                    sheet_name=args["sheet_name"],
                )
            else:
                records = _excel_converter.convert_sheet(
                    file_path=file_path,
                    sheet_name=args["sheet_name"],
                    header_row=args.get("header_row", 25),
                    formula_row=args.get("formula_row", 26),
                )

            from src.services.excel_converter import formula_status as _formula_status
            created_dims = _indicator_exporter.collect_created_dimensions(
                input_excel_path=file_path,
                sheet_name=args["sheet_name"],
            )
            dims_by_indicator: dict[str, list[dict]] = {}
            for d in created_dims:
                ind = d.get("indicator_name", "")
                if not ind:
                    continue
                dims_by_indicator.setdefault(ind, []).append(d)

            out_path = settings.OUTPUT_DIR / args["output_file_name"]
            wb_out = _openpyxl.Workbook()
            ws_out = wb_out.active
            headers = [
                "source_excel_sheet",
                "source_cell",
                "source_formula",
                "indicator_name",
                "indicator_formula",
                "status",
                "creates_dimension",
                "new_dimension_to_create",
                "created_dimension_name",
                "created_dimension_code",
                "created_dimension_source_cell",
                "dimension_creation_note",
            ]
            ws_out.append(headers)
            for rec in records:
                status = _formula_status(rec["indicator_formula"])
                if rec.get("intermediate"):
                    status = status + "_INT"
                dim_entries = dims_by_indicator.get(rec["indicator_name"], [])
                if dim_entries:
                    creates_dim = "YES"
                    has_new_dim = any(bool(d.get("is_new_dimension")) for d in dim_entries)
                    new_dim = "YES" if has_new_dim else ""
                    dim_names = ", ".join(d["dimension_name"] for d in dim_entries)
                    dim_codes = ", ".join(d["dimension_code"] for d in dim_entries)
                    dim_cells = ", ".join(d["source_cell"] for d in dim_entries)
                    if has_new_dim:
                        dim_note = f"Create dimension(s): {dim_names}"
                    else:
                        dim_note = ""
                else:
                    creates_dim = ""
                    new_dim = ""
                    dim_names = ""
                    dim_codes = ""
                    dim_cells = ""
                    dim_note = ""
                ws_out.append([rec["source_excel_sheet"], rec["source_cell"], rec["source_formula"],
                                rec["indicator_name"], rec["indicator_formula"], status,
                                creates_dim, new_dim, dim_names, dim_codes, dim_cells, dim_note])
            wb_out.save(out_path)
            result = {
                "output_file": str(out_path),
                "count": len(records),
                "indicators": [r["indicator_name"] for r in records],
                "created_dimensions": created_dims,
                "known_created_dimensions": _indicator_exporter.get_created_dimensions_registry(),
            }

        elif name == "convert_multi_row_sheet":
            file_path = settings.EXCEL_SOURCES_DIR / args["file_name"]
            if not file_path.exists():
                alt = settings.PROJECT_ROOT / "data" / "examples" / args["file_name"]
                if alt.exists():
                    file_path = alt
                else:
                    raise FileNotFoundError(f"Fichier introuvable: {args['file_name']}")

            _indicator_exporter.register_created_dimensions(file_path, args["sheet_name"])
            _excel_converter.refresh_created_dimensions()

            records = _excel_converter.convert_multi_row_sheet(
                file_path=file_path,
                sheet_name=args["sheet_name"],
            )
            result = {
                "file_name": args["file_name"],
                "sheet_name": args["sheet_name"],
                "count": len(records),
                "indicators": records,
                "known_created_dimensions": _indicator_exporter.get_created_dimensions_registry(),
            }

        elif name == "export_indicators_to_xml":
            file_path = settings.EXCEL_SOURCES_DIR / args["file_name"]
            if not file_path.exists():
                alt = settings.PROJECT_ROOT / "data" / "examples" / args["file_name"]
                if alt.exists():
                    file_path = alt
                else:
                    raise FileNotFoundError(f"Fichier introuvable: {args['file_name']}")

            _indicator_exporter.register_created_dimensions(file_path, args["sheet_name"])
            _excel_converter.refresh_created_dimensions()

            multi_row = args.get("multi_row", False)
            if multi_row:
                records = _excel_converter.convert_multi_row_sheet(
                    file_path=file_path,
                    sheet_name=args["sheet_name"],
                )
            else:
                records = _excel_converter.convert_sheet(
                    file_path=file_path,
                    sheet_name=args["sheet_name"],
                    header_row=args.get("header_row", 25),
                    formula_row=args.get("formula_row", 26),
                )

            out_path = settings.OUTPUT_DIR / args["output_file_name"]
            count = _indicator_exporter.export_to_xml(
                records=records,
                output_path=out_path,
                data_types=args.get("data_types"),
            )
            result = {
                "output_file": str(out_path),
                "count": count,
                "indicators": [r["indicator_name"] for r in records],
                "known_created_dimensions": _indicator_exporter.get_created_dimensions_registry(),
            }

        elif name == "export_analysis_table_spec":
            file_path = settings.EXCEL_SOURCES_DIR / args["file_name"]
            if not file_path.exists():
                alt = settings.PROJECT_ROOT / "data" / "examples" / args["file_name"]
                if alt.exists():
                    file_path = alt
                else:
                    raise FileNotFoundError(f"Fichier introuvable: {args['file_name']}")

            out_path = settings.OUTPUT_DIR / args["output_file_name"]
            _indicator_exporter.register_created_dimensions(file_path, args["sheet_name"])
            created_dims = _indicator_exporter.collect_created_dimensions(
                input_excel_path=file_path,
                sheet_name=args["sheet_name"],
            )
            n_cols = _indicator_exporter.export_analysis_spec(
                input_excel_path=file_path,
                sheet_name=args["sheet_name"],
                output_path=out_path,
                header_row=args.get("header_row", 25),
                formula_row=args.get("formula_row", 26),
            )
            result = {
                "output_file": str(out_path),
                "data_columns": n_cols,
                "sheet_name": args["sheet_name"],
                "created_dimensions": created_dims,
                "known_created_dimensions": _indicator_exporter.get_created_dimensions_registry(),
            }

        else:
            result = {"error": f"Outil inconnu: {name}"}

    except Exception as e:
        logger.error(f"Erreur dans l'outil '{name}': {e}", exc_info=True)
        result = {"error": str(e), "tool": name}

    return [
        types.TextContent(
            type="text",
            text=json.dumps(result, indent=2, default=str, ensure_ascii=False),
        )
    ]


# ── Point d'entrée ────────────────────────────────────────────────────────────


async def main() -> None:
    logger.info("Démarrage du serveur MCP Excel-to-Indicator")
    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name="excel-to-indicator",
                server_version="0.1.0",
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )


if __name__ == "__main__":
    asyncio.run(main())
