# ==============================================================================
# tools/07_hara_export.py
# Tool for exporting the complete HARA (HAZOP, Situations, E/S/C, ASIL, SGs)
# to a single, formatted Excel file.
# ==============================================================================

from cat.mad_hatter.decorators import tool
from cat.log import log
import sys
import os
import json
from datetime import datetime
from ..generators.HARA.hara_excel_generator import HaraExcelGenerator 

# Setup paths
current_file = os.path.abspath(__file__)
tools_folder = os.path.dirname(current_file)
plugin_folder = os.path.dirname(tools_folder)
code_folder = os.path.dirname(tools_folder)

# Add core and generators module path
sys.path.insert(0, os.path.join(plugin_folder, 'core'))
sys.path.insert(0, os.path.join(code_folder, 'generators'))


# Import the new generator
try:
    from ..generators.HARA.hara_excel_generator import HaraExcelGenerator, EXCEL_AVAILABLE
except ImportError as e:
    log.error(f"Failed to import HaraExcelGenerator: {e}")
    EXCEL_AVAILABLE = False


# ==============================================================================
# Cat-facing @tool decorator
# ==============================================================================

@tool(
    return_direct=True,
    examples=[
        "export hara to excel",
        "generate hara report",
        "create hara excel file",
        "save hara as excel"
    ]
)
def export_hara_to_excel(tool_input, cat):
    """
    Generates a complete HARA (Hazard Analysis and Risk Assessment) report
    as a formatted Excel file, styled like 'hara_dev_xls.py'.
    
    The report includes all data from the HARA workflow:
    - Sheet 1: Summary (Statistics)
    - Sheet 2: Operational Situations
    - Sheet 3: HARA Table (HAZOP, E/S/C, ASIL, Safety Goals)
    - Sheet 4: Safety Goals Summary
    
    The file is saved in the 'generated_documents/03_HARA' folder.
    
    Args:
        tool_input: Ignored.
        cat: Cheshire Cat instance.
        
    Returns:
        JSON string with the success message and file path.
    """
    
    log.info("🔧 TOOL CALLED: export_hara_to_excel (hara_dev_xls style)")
    
    if not EXCEL_AVAILABLE:
        log.error("HARA Excel export failed: openpyxl is not installed.")
        return json.dumps({
            "status": "error",
            "message": "Excel export failed: `openpyxl` library not found.",
            "suggestion": "Please install the library: `pip install openpyxl`"
        })

    # Get data from working memory
    try:
        item_name = cat.working_memory['hara_item_name']
        sit_data = cat.working_memory['operational_situations_complete_output']
        # Use asil_determination output as the main HARA table data
        hara_data = cat.working_memory['asil_determination_complete_output']
        sg_data = cat.working_memory['safety_goals_complete_output']
        
        # Simple validation
        if not all([item_name, sit_data, hara_data, sg_data]):
            raise KeyError("One or more required HARA data components are missing.")
            
    except KeyError as e:
        log.warning(f"HARA export failed: Missing data in working memory - {e}")
        missing_keys = [
            k for k in ['hara_item_name', 'operational_situations_complete_output', 
                        'asil_determination_complete_output', 'safety_goals_complete_output']
            if k not in cat.working_memory
        ]
        return json.dumps({
            "status": "error",
            "message": "Cannot generate HARA report: Missing required data.",
            "missing_data_keys": missing_keys,
            "suggestions": [
                "Complete the full HARA workflow first:",
                "1. `define item [name]`",
                "2. `apply hazop analysis`",
                "3. `define operational situations`",
                "4. `assess esc for all hazards`",
                "5. `determine asil`",
                "6. `derive safety goals`",
                "Then run `export hara to excel`"
            ]
        })
        
    try:
        # Setup output directory
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(plugin_folder, "..\generated_documents", "03_HARA")
        os.makedirs(output_dir, exist_ok=True)
        
        # Sanitize system name for filename
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in item_name).replace(" ", "_")
        
        filename = f"HARA_Report_{safe_name}_{timestamp}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        # Generate workbook
        log.info(f"📊 Creating HARA Excel generator for: {item_name}")
        generator = HaraExcelGenerator(item_name, sit_data, hara_data, sg_data)
        wb = generator.generate_workbook()
        
        # Save workbook
        wb.save(filepath)
        log.info(f"✅ HARA Excel saved: {filepath}")


        return f"""✅ **HARA Excel Report Generated Successfully!!** 

       File: `{filename}`
       Location: `generated_documents/03_HARA/`

        ✅ HARA WORKFLOW COMPLETED
        Next steps 
        1. 📖 Review the FSRs
        2. ✍️ Complete the approvals section
        """


        # Build response
        response = {
            "status": "success",
            "message": "HARA Excel Report Generated Successfully!",
            "file_name": filename,
            "location": "generated_documents/03_HARA/",
            "full_path": filepath,
            "sheets": [sheet.title for sheet in wb.worksheets],
            "total_hazards": len(hara_data.get('hazards', [])),
            "total_safety_goals": len(sg_data.get('safety_goals', []))
        }
        return json.dumps(response, indent=2)

    except Exception as e:
        log.error(f"❌ HARA Excel generation failed: {e}")
        import traceback
        log.error(traceback.format_exc())
        return json.dumps({
            "status": "error",
            "message": f"An unexpected error occurred: {str(e)}",
            "traceback": traceback.format_exc()
        })