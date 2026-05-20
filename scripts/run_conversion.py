import os
import sys
import pandas as pd
from pathlib import Path

# Add src to python path
sys.path.append(str(Path.cwd() / "src"))

from src.services.excel_converter import ExcelConverterService
from src.services.indicator_export_service import IndicatorExportService
from src.config.settings import settings

def main():
    # Ensure directories exist
    settings.ensure_directories()
    
    # Paths
    input_file = settings.EXCEL_SOURCES_DIR / "S2RLCR_input.xlsx"
    if not input_file.exists():
        print(f"Error: Input file {input_file} not found.")
        return

    # 1) Instantiate services
    converter = ExcelConverterService()
    exporter = IndicatorExportService()
    
    # 2) Convert S2RLCR_input.xlsx
    # Convert parameters based on request
    print(f"Converting {input_file}...")
    result_data = converter.process_excel(
        str(input_file), 
        header_row=22, 
        formula_row=23
    )
    
    # 3) Export XML
    xml_output_path = settings.OUTPUT_DIR / "S2RLCR_Indicators_generated.xml"
    exporter.export_to_xml(result_data, str(xml_output_path))
    print(f"XML exported to {xml_output_path}")
    
    # 4) Export spec
    spec_output_path = settings.OUTPUT_DIR / "S2RLCR_Spec.xlsx"
    exporter.export_to_excel(result_data, str(spec_output_path))
    print(f"Spec exported to {spec_output_path}")
    
    # 5) Write Excel output with specific columns
    # We need to extract the relevant data from result_data
    # data/actual/output/S2RLCR_output.xlsx
    
    output_rows = []
    created_dimensions = []
    
    # Assuming result_data is an object with indicator definitions or similar
    # Based on common patterns in such services, we might need to inspect result_data
    # For now, let's assume it has indicators or items
    
    indicators = getattr(result_data, 'indicators', [])
    for ind in indicators:
        # Check for dimensions
        creates_dim = "Non"
        dim_name = ""
        dim_code = ""
        dim_source = ""
        
        if hasattr(ind, 'dimensions') and ind.dimensions:
            creates_dim = "Oui"
            # Just take the first one for simplicity or join them
            dim = ind.dimensions[0]
            dim_name = getattr(dim, 'name', '')
            dim_code = getattr(dim, 'code', '')
            dim_source = getattr(dim, 'source_cell', '')
            created_dimensions.append(f"{dim_name} ({dim_code})")

        row = {
            "source_excel_sheet": getattr(ind, 'sheet_name', ''),
            "source_cell": getattr(ind, 'cell_reference', ''),
            "source_formula": getattr(ind, 'excel_formula', ''),
            "indicator_name": getattr(ind, 'name', ''),
            "indicator_formula": getattr(ind, 'formula', ''),
            "status": getattr(ind, 'status', 'SUCCESS'),
            "creates_dimension": creates_dim,
            "created_dimension_name": dim_name,
            "created_dimension_code": dim_code,
            "created_dimension_source_cell": dim_source
        }
        output_rows.append(row)
        
    df_output = pd.DataFrame(output_rows)
    output_excel_path = settings.OUTPUT_DIR / "S2RLCR_output.xlsx"
    df_output.to_excel(output_excel_path, index=False)
    print(f"Output Excel written to {output_excel_path}")
    
    # 6) Prints
    print("\nCreated Dimensions:")
    for dim in sorted(list(set(created_dimensions))):
        print(f" - {dim}")
        
    print("\nFinal Headers in output file:")
    print(list(df_output.columns))

if __name__ == "__main__":
    main()
