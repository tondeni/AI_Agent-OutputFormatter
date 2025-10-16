# ==============================================================================
# AI_Agent-OutputFormatter/tools/fsr_file_tools.py
# Tools for FSR and FSC file generation
# ==============================================================================

from cat.mad_hatter.decorators import tool
from cat.log import log
import os
import sys
from datetime import datetime

# Add parent directory to path for imports
# Calculate plugin root directory
# File is at: AI_Agent-OutputFormatter/code/tools/fsr_file_tools.py
# We need to go up to: AI_Agent-OutputFormatter/
current_file = os.path.abspath(__file__)
tools_folder = os.path.dirname(current_file)   # .../code/tools
code_folder = os.path.dirname(tools_folder)    # .../code
plugin_folder = os.path.dirname(code_folder)   # .../AI_Agent-OutputFormatter



@tool(
    return_direct=True,
    examples=[
        "create fsr excel",
        "generate FSR spreadsheet",
        "export FSRs to excel"
    ]
)
def create_fsr_excel(tool_input, cat):
    """
    Generate Excel file with FSR listing and allocation matrix.
    
    Creates comprehensive Excel workbook with:
    - FSRs Sheet: Complete FSR listing
    - Allocation Matrix: FSRs mapped to components
    - Traceability: FSR to Safety Goal mapping
    - Statistics: Distribution charts and summaries
    
    Examples:
    - "create fsr excel"
    - "generate FSR spreadsheet"
    """
    
    log.info("✅ TOOL CALLED: create_fsr_excel")
    
    # Get data from working memory
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    goals_data = cat.working_memory.get("fsc_safety_goals", [])
    system_name = cat.working_memory.get("system_name", "System")
    
    if not fsrs_data:
        return """❌ No FSR data available to export.

**Required:** Derive FSRs first using: `derive FSRs for all goals`

Then you can generate Excel files with the FSC data."""
    
    try:
        # Import Excel generator
        from ..generators import generate_fsr_excel
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(plugin_folder, "generated_documents", "06_FSRs")
        os.makedirs(output_dir, exist_ok=True)
        
        # Sanitize system name for filename
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in system_name).replace(" ", "_")
        
        filename = f"FSR_{safe_name}_{timestamp}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        # Create Excel workbook
        log.info(f"📊 Generating Excel file: {filename}")
        wb = generate_fsr_excel(fsrs_data, goals_data, system_name)
        wb.save(filepath)
        
        # Calculate statistics
        allocated = sum(1 for f in fsrs_data 
                       if f.get('allocated_to') and f.get('allocated_to') != 'TBD')
        
        by_asil = {}
        for f in fsrs_data:
            asil = f.get('asil', 'QM')
            by_asil[asil] = by_asil.get(asil, 0) + 1
        
        asil_summary = ', '.join([f"ASIL {asil}: {count}" 
                                 for asil, count in sorted(by_asil.items(), reverse=True)])
        
        return f"""✅ **Excel file generated successfully!**

**File:** `{filename}`
**Location:** `{output_dir}`

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
4. Statistics - Summary charts

The file is ready for review and can be shared with your safety team.
"""
        
    except ImportError as e:
        log.error(f"❌ Import error: {e}")
        return f"""❌ Excel generator not available.

**Error:** {str(e)}

**Solution:**
- Check `generators/fsr_excel_generator.py` exists
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
"""


@tool(
    return_direct=True,
    examples=[
        "create allocation report",
        "show allocation analysis",
        "allocation summary"
    ]
)
def create_allocation_report(tool_input, cat):
    """
    Generate detailed allocation analysis report (text-based, in chat).
    
    Creates allocation report with:
    - Component allocation matrix
    - ASIL distribution per component
    - Coverage analysis
    - Unallocated FSRs list
    
    This is displayed in chat, not saved as a file.
    """
    
    log.info("✅ TOOL CALLED: create_allocation_report")
    
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    system_name = cat.working_memory.get("system_name", "System")
    
    if not fsrs_data:
        return "❌ No FSRs available. Please derive FSRs first."
    
    # Analyze allocation
    by_component = {}
    unallocated = []
    
    for fsr in fsrs_data:
        component = fsr.get('allocated_to', 'TBD')
        if component in ['TBD', 'NOT ALLOCATED', 'N/A', '', None]:
            unallocated.append(fsr)
        else:
            if component not in by_component:
                by_component[component] = []
            by_component[component].append(fsr)
    
    # Build report
    report = f"""📊 **FSR Allocation Report: {system_name}**

**Overall Statistics:**
- Total FSRs: {len(fsrs_data)}
- Allocated: {len(fsrs_data) - len(unallocated)}
- Unallocated: {len(unallocated)}
- Components: {len(by_component)}

---

**Allocation by Component:**

"""
    
    for component, fsrs in sorted(by_component.items()):
        by_asil = {}
        for fsr in fsrs:
            asil = fsr.get('asil', 'QM')
            by_asil[asil] = by_asil.get(asil, 0) + 1
        
        asil_dist = ', '.join([f"ASIL {asil}: {count}" 
                              for asil, count in sorted(by_asil.items(), reverse=True)])
        
        report += f"""### {component}
- Total FSRs: {len(fsrs)}
- ASIL Distribution: {asil_dist}

"""
    
    if unallocated:
        report += f"""---

⚠️ **Unallocated FSRs ({len(unallocated)}):**

"""
        for fsr in unallocated[:5]:
            report += f"- {fsr.get('id')}: {fsr.get('description')[:80]}...\n"
        
        if len(unallocated) > 5:
            report += f"\n... and {len(unallocated) - 5} more\n"
    
    report += """

**Next Steps:**
- Allocate remaining FSRs: `allocate all FSRs`
- Generate Excel: `create fsr excel`
"""
    
    return report


@tool(
    return_direct=True,
    examples=[
        "create fsc document",
        "generate word report",
        "create FSC word document"
    ]
)
def create_fsc_document(tool_input, cat):
    """
    Generate complete Word document with FSC report.
    
    Creates professional ISO 26262 compliant FSC document with:
    - Executive summary
    - Safety goals from HARA
    - Functional Safety Requirements
    - Allocation to architecture
    - Compliance documentation
    
    Examples:
    - "create fsc document"
    - "generate word report"
    """
    
    log.info("✅ TOOL CALLED: create_fsc_document")
    
    goals_data = cat.working_memory.get("fsc_safety_goals", [])
    strategies_data = cat.working_memory.get("fsc_safety_strategies", [])
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    system_name = cat.working_memory.get("system_name", "System")
    
    if not fsrs_data or not goals_data:
        return """❌ Insufficient FSC data to generate document.

**Required:** Complete the FSC development workflow:
1. `load HARA for [system name]`
2. `develop safety strategy for all goals`
3. `derive FSRs for all goals`

Then you can generate Word documents with the complete FSC.
"""
    
    try:
        # Import Word generator
        from generators.fsc_word_generator import generate_fsc_word
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(plugin_folder, "generated_documents", "06_FSC")
        os.makedirs(output_dir, exist_ok=True)
        
        # Sanitize system name
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in system_name).replace(" ", "_")
        
        filename = f"FSC_Report_{safe_name}_{timestamp}.docx"
        filepath = os.path.join(output_dir, filename)
        
        # Create Word document
        log.info(f"📄 Generating Word document: {filename}")
        doc = generate_fsc_word(system_name, goals_data, strategies_data, fsrs_data)
        doc.save(filepath)
        
        # Estimate pages
        pages_estimate = len(goals_data) * 2 + len(fsrs_data) // 3 + 15
        
        return f"""✅ **Word document generated successfully!**

**File:** `{filename}`
**Location:** `{output_dir}`
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

**Solution:**
- Check `generators/fsc_word_generator.py` exists
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
"""


@tool(
    return_direct=False,
    examples=[
        "set output format to minimal",
        "use detailed output",
        "format as standard"
    ]
)
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
    """
    
    log.info(f"✅ TOOL CALLED: set_output_format")
    
    input_lower = str(tool_input).lower()
    
    if "minimal" in input_lower or "plain" in input_lower:
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
    
    return f"Output format preference updated to: **{format_type}**. {description}. This will apply to future responses."