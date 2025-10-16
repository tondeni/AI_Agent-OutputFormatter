# ==============================================================================
# fsc_file_tools.py (NEW FILE)
# Tools for FSC file generation (Excel and Word)
# Place in: AI_Agent-OutputFormatter/fsc_file_tools.py
# ==============================================================================

from cat.mad_hatter.decorators import tool
from cat.log import log
import os
from datetime import datetime


@tool(return_direct=True)
def create_excel_file(tool_input, cat):
    """
    Generate Excel file with FSC data (FSRs, allocation matrix, traceability).
    
    Creates a comprehensive Excel workbook with multiple sheets:
    - FSRs listing
    - Allocation matrix
    - Traceability to safety goals
    - Statistics and summaries
    
    Examples: 
    - "create excel file"
    - "generate excel spreadsheet"
    - "export FSRs to excel"
    """
    
    log.info("✅ TOOL CALLED: create_excel_file")
    
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    system_name = cat.working_memory.get("system_name", "System")
    
    if not fsrs_data:
        return """❌ No FSC data available to export.

**Required:** Derive FSRs first using: `derive FSRs for all goals`

Then you can generate Excel files with the FSC data."""
    
    try:
        # Import Excel generator
        from .fsr_excel_generator import generate_fsr_excel
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        plugin_folder = os.path.dirname(__file__)
        output_dir = os.path.join(plugin_folder, "generated_documents", "06_FSC")
        os.makedirs(output_dir, exist_ok=True)
        
        # Sanitize system name for filename
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in system_name).replace(" ", "_")
        
        filename = f"FSC_{safe_name}_{timestamp}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        # Create Excel workbook
        log.info(f"📊 Generating Excel file: {filename}")
        wb = generate_fsr_excel(fsrs_data, system_name)
        wb.save(filepath)
        
        # Calculate statistics
        allocated = sum(1 for f in fsrs_data if f.get('allocated_to'))
        by_asil = {}
        for f in fsrs_data:
            asil = f.get('asil', 'QM')
            by_asil[asil] = by_asil.get(asil, 0) + 1
        
        asil_summary = ', '.join([f"ASIL {asil}: {count}" for asil, count in sorted(by_asil.items(), reverse=True)])
        
        return f"""✅ **Excel file generated successfully!**

**File:** `{filename}`
**Location:** `generated_documents/06_FSC/`

**Contents:**
- 📋 **FSRs Sheet**: {len(fsrs_data)} Functional Safety Requirements
- 🗂️ **Allocation Matrix**: {allocated}/{len(fsrs_data)} FSRs allocated
- 🔗 **Traceability**: FSR → Safety Goal mapping
- 📊 **Statistics**: Distribution by ASIL and type

**ASIL Distribution:** {asil_summary}

**Worksheets:**
1. FSRs - Complete listing with all details
2. Allocation Matrix - FSRs grouped by component
3. Traceability - Mapping to safety goals
4. Statistics - Summary and charts

The file is ready for review and can be shared with your safety team.
"""
        
    except ImportError as e:
        log.error(f"❌ Import error: {e}")
        return f"""❌ Excel generator not available.

**Error:** {str(e)}

**Solution:** Ensure the FSR Excel generator module is installed:
- Check `fsr_excel_generator.py` exists in the plugin folder
- Verify `openpyxl` is installed: `pip install openpyxl`
"""
    except Exception as e:
        log.error(f"❌ Excel generation failed: {e}")
        import traceback
        log.error(traceback.format_exc())
        return f"""❌ Failed to generate Excel file.

**Error:** {str(e)}

**Troubleshooting:**
1. Check the plugin logs for details
2. Verify FSR data is valid in working memory
3. Ensure write permissions for generated_documents folder

Try again or contact support if the issue persists.
"""


@tool(return_direct=True)
def create_word_document(tool_input, cat):
    """
    Generate Word document with complete FSC report.
    
    Creates a professional ISO 26262 compliant FSC document with:
    - Executive summary
    - Safety goals from HARA
    - Functional Safety Requirements
    - Allocation to architecture
    - Compliance documentation
    
    Examples:
    - "create word document"
    - "generate word report"
    - "create FSC document"
    """
    
    log.info("✅ TOOL CALLED: create_word_document")
    
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    goals_data = cat.working_memory.get("fsc_safety_goals", [])
    strategies_data = cat.working_memory.get("fsc_strategies", [])
    system_name = cat.working_memory.get("system_name", "System")
    
    if not fsrs_data and not goals_data:
        return """❌ No FSC data available to export.

**Required:** Complete the FSC development workflow:
1. `load HARA for [system name]`
2. `develop safety strategy for all goals`
3. `derive FSRs for all goals`

Then you can generate Word documents with the complete FSC.
"""
    
    try:
        # Import Word generator
        from .fsc_word_generator import generate_fsc_word
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        plugin_folder = os.path.dirname(__file__)
        output_dir = os.path.join(plugin_folder, "generated_documents", "06_FSC")
        os.makedirs(output_dir, exist_ok=True)
        
        # Sanitize system name for filename
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in system_name).replace(" ", "_")
        
        filename = f"FSC_Report_{safe_name}_{timestamp}.docx"
        filepath = os.path.join(output_dir, filename)
        
        # Create Word document
        log.info(f"📄 Generating Word document: {filename}")
        doc = generate_fsc_word(system_name, goals_data, fsrs_data, strategies_data)
        doc.save(filepath)
        
        # Calculate document statistics
        pages_estimate = len(goals_data) * 2 + len(fsrs_data) // 3 + 10
        
        return f"""✅ **Word document generated successfully!**

**File:** `{filename}`
**Location:** `generated_documents/06_FSC/`
**Estimated Pages:** ~{pages_estimate} pages

**Document Structure:**
1. 📋 **Title Page** - System and document information
2. 📝 **Executive Summary** - Overview and key metrics
3. 🎯 **Safety Goals** ({len(goals_data)}) - From HARA analysis
4. 🛡️ **Safety Strategies** - 9 strategy types per ISO 26262
5. 📌 **Functional Safety Requirements** ({len(fsrs_data)}) - Complete FSR specifications
6. 🏗️ **Allocation** - FSRs mapped to system architecture
7. ✅ **Compliance** - ISO 26262-3:2018 Clause 7 requirements

**ISO 26262 Compliance:**
- ✅ Clause 7.4.2 - FSR specification
- ✅ Clause 7.4.2.8 - Architectural allocation
- ✅ Clause 7.5 - Work product requirements

The document is ready for safety assessment and can be used in your safety case.
"""
        
    except ImportError as e:
        log.error(f"❌ Import error: {e}")
        return f"""❌ Word generator not available.

**Error:** {str(e)}

**Solution:** Ensure the FSC Word generator module is installed:
- Check `fsc_word_generator.py` exists in the plugin folder
- Verify `python-docx` is installed: `pip install python-docx`
"""
    except Exception as e:
        log.error(f"❌ Word generation failed: {e}")
        import traceback
        log.error(traceback.format_exc())
        return f"""❌ Failed to generate Word document.

**Error:** {str(e)}

**Troubleshooting:**
1. Check the plugin logs for details
2. Verify FSC data is valid in working memory
3. Ensure write permissions for generated_documents folder

Try again or contact support if the issue persists.
"""


@tool(return_direct=True)
def generate_fsc_files(tool_input, cat):
    """
    Generate both Excel and Word files for FSC.
    
    Creates complete documentation package:
    - Excel spreadsheet with FSRs and allocation
    - Word document with FSC report
    
    Examples:
    - "generate fsc files"
    - "create all fsc documents"
    - "export complete fsc"
    """
    
    log.info("✅ TOOL CALLED: generate_fsc_files")
    
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    
    if not fsrs_data:
        return """❌ No FSC data available.

**Required:** Derive FSRs first: `derive FSRs for all goals`
"""
    
    # Generate Excel
    excel_result = create_excel_file("", cat)
    
    # Generate Word
    word_result = create_word_document("", cat)
    
    # Combine results
    return f"""✅ **Complete FSC documentation package generated!**

---

{excel_result}

---

{word_result}

---

**Next Steps:**
1. Review the generated files in `generated_documents/06_FSC/`
2. Share with your safety team for review
3. Include in your ISO 26262 work product documentation
"""


@tool(return_direct=False)
def set_output_format(tool_input, cat):
    """
    Set output format preference for agent responses.
    
    Options:
    - "minimal" - Plain text, no markdown/tables
    - "standard" - Normal formatting (default)
    - "detailed" - Extra information and context
    
    Examples:
    - "set output format to minimal"
    - "format output as detailed"
    - "use standard formatting"
    """
    
    log.info(f"✅ TOOL CALLED: set_output_format with input: {tool_input}")
    
    input_lower = str(tool_input).lower()
    
    if "minimal" in input_lower or "plain" in input_lower or "simple" in input_lower:
        format_type = "minimal"
        description = "Plain text without tables or markdown"
    elif "detailed" in input_lower or "verbose" in input_lower:
        format_type = "detailed"
        description = "Detailed formatting with extra context"
    else:
        format_type = "standard"
        description = "Normal formatting with tables and markdown"
    
    # Store preference
    cat.working_memory["output_format"] = format_type
    
    log.info(f"📝 Output format set to: {format_type}")
    
    return {
        "output": f"Output format preference updated to: **{format_type}**. "
                  f"{description}. This will apply to future responses."
    }