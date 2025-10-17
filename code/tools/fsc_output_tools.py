# ==============================================================================
# AI_Agent-OutputFormatter/code/tools/fsc_output_tools.py
# Simple tools that call self-contained generators with completeness warnings
# ==============================================================================

from cat.mad_hatter.decorators import tool
from cat.log import log
import os
import sys
from datetime import datetime
from ..generators.Functional_Safety_Concept.fsc_word_generator import FSCWordGenerator 
from ..generators.Functional_Safety_Concept.fsc_excel_generator import FSCExcelGenerator
from ..generators.Functional_Safety_Concept.fsr_excel_generator import generate_fsr_excel

# Setup paths
current_file = os.path.abspath(__file__)
tools_folder = os.path.dirname(current_file)
code_folder = os.path.dirname(tools_folder)
plugin_folder = os.path.dirname(code_folder)

@tool(
    return_direct=True,
    examples=[
        "create fsc word document",
        "generate fsc document",
        "export fsc to word"
    ]
)
def create_fsc_word_document(tool_input, cat):
    """
    Generate FSC Word document.
    
    Reads structured content from working memory (contract-based).
    No dependencies on FSC Developer plugin.
    """
    
    log.info("✅ TOOL CALLED: create_fsc_word_document")
    
    # Check for structured content (contract v1.0)
    structured_content = cat.working_memory.get("fsc_structured_content")
    
    if structured_content:
        log.info("✅ Found structured content (contract v1.0)")
        system_name = structured_content.get('system_name', 'Unknown')
        
        # Validate schema version
        metadata = structured_content.get('metadata', {})
        schema_version = metadata.get('schema_version', 'unknown')
        
        if schema_version != '1.0':
            return f"""⚠️ Schema version mismatch

**Expected:** 1.0
**Found:** {schema_version}

Content may be from incompatible FSC Developer version.
Please regenerate FSC content."""
    else:
        # Fall back to legacy
        log.info("Using legacy format")
        goals_data = cat.working_memory.get("fsc_safety_goals", [])
        fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
        system_name = cat.working_memory.get("system_name", "System")
        
        if not goals_data or not fsrs_data:
            return """❌ No FSC data found.

Generate FSC content first:
- New: `generate structured FSC content` (FSC Developer)
- Legacy: `derive FSRs for all goals` (FSC Developer)

Then retry document generation."""
    
    try:
        # Import generator (self-contained, no external deps)
        from ..generators.Functional_Safety_Concept.fsc_word_generator import FSCWordGenerator  
        
        generator = FSCWordGenerator()
        
        # Generate document
        if structured_content:
            log.info("Generating from structured content")
            doc = generator.generate(structured_content=structured_content)
            format_type = "Structured (Contract v1.0)"
        else:
            log.info("Generating from legacy format")
            doc = generator.generate(
                system_name=system_name,
                goals_data=goals_data,
                fsrs_data=fsrs_data
            )
            format_type = "Legacy"
        
        # Save
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(plugin_folder, "generated_documents", "06_FSC")
        os.makedirs(output_dir, exist_ok=True)
        
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in system_name).replace(" ", "_")
        
        filename = f"FSC_{safe_name}_{timestamp}.docx"
        filepath = os.path.join(output_dir, filename)
        
        doc.save(filepath)
        
        # Get stats if structured content
        stats_info = ""
        if structured_content:
            num_fsrs = len(structured_content.get('functional_safety_requirements', []))
            num_sms = len(structured_content.get('safety_mechanisms', []))
            stats_info = f"""
📊 **Content:**
- {num_fsrs} Functional Safety Requirements
- {num_sms} Safety Mechanisms
- Traceability matrix included
"""
        
        return f"""✅ **FSC Word Document Generated!**

**File:** `{filename}`
**Location:** `generated_documents/06_FSC/`
**Format:** {format_type}
{stats_info}
📄 **Document includes:**
- Title page with ISO 26262 reference
- Introduction and safety goals
- Detailed FSR specifications
- Safety mechanisms by type
- FSR-SM traceability matrix
- Architectural allocation
- Verification strategy

**Next Steps:**
1. 📖 Review in Microsoft Word
2. 👥 Share with safety team
3. ✍️ Complete approvals section"""
        
    except Exception as e:
        log.error(f"Document generation failed: {e}")
        import traceback
        log.error(traceback.format_exc())
        return f"❌ Error: {str(e)}"
 

@tool(
    return_direct=True,
    examples=[
        "create fsc excel",
        "generate fsc spreadsheet",
        "export fsc to excel"
    ]
)
def create_fsc_excel(tool_input, cat):
    """
    Generate FSC Excel workbook.
    
    Creates Excel workbook with multiple sheets:
    - Safety Goals
    - FSRs
    - Allocation Matrix
    - Traceability
    - Statistics
    
    Examples:
    - "create fsc excel"
    - "generate spreadsheet"
    """
    
    log.info("✅ TOOL CALLED: create_fsc_excel")
    
    # Get data
    goals_data = cat.working_memory.get("fsc_safety_goals", [])
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    strategies_data = cat.working_memory.get("fsc_safety_strategies", {})
    system_name = cat.working_memory.get("system_name", "System")
    allocation_data = cat.working_memory.get("fsc_allocation_matrix", {})
    
    # Validate
    if not goals_data or not fsrs_data:
        return """❌ Insufficient FSC data for Excel generation.

**Missing:**
- Safety Goals: {'✅' if goals_data else '❌'}
- FSRs: {'✅' if fsrs_data else '❌'}

Complete FSC development first."""
    
    try:
        # Import generator
        from fsc_excel_generator import FSCExcelGenerator
        
        # Create generator
        generator = FSCExcelGenerator()
        
        # Validate
        is_valid, warnings, errors = generator.validate_data(goals_data, fsrs_data)
        
        if not is_valid:
            return "❌ Data validation failed: " + "; ".join(errors)
        
        # Calculate statistics
        stats = generator.calculate_statistics(goals_data, fsrs_data)
        
        # Generate workbook
        log.info(f"📊 Generating Excel for {system_name}")
        
        wb = generator.generate(
            system_name=system_name,
            goals_data=goals_data,
            fsrs_data=fsrs_data,
            strategies_data=strategies_data,
            allocation_data=allocation_data
        )
        
        # Save workbook
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(plugin_folder, "generated_documents", "06_FSC")
        os.makedirs(output_dir, exist_ok=True)
        
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in system_name).replace(" ", "_")
        
        filename = f"FSC_{safe_name}_{timestamp}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        wb.save(filepath)
        
        return f"""✅ **FSC Excel Workbook Generated!**

**File:** `{filename}`
**Location:** `generated_documents/06_FSC/`

**Sheets:**
1. 📋 Safety Goals ({stats['total_goals']} goals)
2. 📌 FSRs ({stats['total_fsrs']} requirements)
3. 🏗️ Allocation Matrix
4. 🔗 Traceability
5. 📊 Statistics

**Content:**
- Safety Goals: {stats['total_goals']}
- FSRs: {stats['total_fsrs']}
- Allocated: {stats['allocated_fsrs']} ({stats['allocated_fsrs']/stats['total_fsrs']*100:.0f}%)

**ASIL Distribution:**
{chr(10).join([f"  - ASIL {asil}: {count}" for asil, count in sorted(stats['asil_distribution'].items())])}

Ready for filtering, sorting, and analysis in Excel!"""
        
    except ImportError as e:
        log.error(f"❌ Import error: {e}")
        return f"""❌ Excel generator not available.

**Error:** {str(e)}

**Solution:**
Install openpyxl: `pip install openpyxl`
"""
        
    except Exception as e:
        log.error(f"❌ Excel generation failed: {e}")
        import traceback
        log.error(traceback.format_exc())
        return f"❌ Failed to generate Excel: {str(e)}"


@tool(
    return_direct=True,
    examples=[
        "generate complete fsc package",
        "create all fsc files",
        "export fsc documentation"
    ]
)
def generate_complete_fsc_package(tool_input, cat):
    """
    Generate complete FSC documentation package (Word + Excel).
    
    Creates both Word document and Excel spreadsheet in one command.
    Includes completeness warnings for any missing sections.
    
    Examples:
    - "generate complete fsc package"
    - "create all fsc files"
    """
    
    log.info("✅ TOOL CALLED: generate_complete_fsc_package")
    
    # Check data
    goals_data = cat.working_memory.get("fsc_safety_goals", [])
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    
    if not goals_data or not fsrs_data:
        return """❌ Insufficient FSC data.

Complete FSC development first:
1. Load HARA
2. Derive FSRs
3. Generate package
"""
    
    results = []
    
    # Generate Word
    log.info("📄 Generating Word document...")
    word_result = create_fsc_word_document("", cat)
    results.append(("📄 Word Document", word_result))
    
    # Generate Excel
    log.info("📊 Generating Excel workbook...")
    excel_result = create_fsc_excel("", cat)
    results.append(("📊 Excel Workbook", excel_result))
    
    # Combine
    output = "✅ **Complete FSC Documentation Package Generated!**\n\n"
    output += "="*70 + "\n\n"
    
    for doc_type, result in results:
        output += f"### {doc_type}\n\n{result}\n\n"
        output += "="*70 + "\n\n"
    
    output += """
**📁 Documentation Package Complete!**

Both files are in `generated_documents/06_FSC/`

**Review Checklist:**
□ Check Word document for ⚠️ INCOMPLETE markers
□ Complete any missing workflow steps
□ Review all sections for accuracy
□ Share with safety team for peer review
□ Obtain formal approvals

**Next Steps:**
1. 📖 Review all documents
2. ✅ Complete missing sections (if any)
3. 👥 Safety team review
4. 📁 Include in safety case
5. ➡️ Proceed to Technical Safety Concept

Your FSC documentation is ready!"""
    
    return output