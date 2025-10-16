# ==============================================================================
# AI_Agent-OutputFormatter/code/tools/fsc_output_tools.py
# Simple tools that call self-contained generators with completeness warnings
# ==============================================================================

from cat.mad_hatter.decorators import tool
from cat.log import log
import os
import sys
from datetime import datetime

# Setup paths
current_file = os.path.abspath(__file__)
tools_folder = os.path.dirname(current_file)
code_folder = os.path.dirname(tools_folder)
plugin_folder = os.path.dirname(code_folder)

sys.path.insert(0, os.path.join(code_folder, 'generators'))

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
    
    Creates ISO 26262-3:2018 Clause 7 compliant Word document from FSC data
    in working memory. Includes completeness warnings for missing sections.
    
    Examples:
    - "create fsc word document"
    - "generate word file"
    - "export fsc to word"
    """
    
    log.info("✅ TOOL CALLED: create_fsc_word_document")
    
    # Get data from working memory
    goals_data = cat.working_memory.get("fsc_safety_goals", [])
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    strategies_data = cat.working_memory.get("fsc_safety_strategies", {})
    system_name = cat.working_memory.get("system_name", "System")
    
    # Quick validation
    if not goals_data or not fsrs_data:
        return """❌ Insufficient FSC data for Word document generation.

**Missing:**
- Safety Goals: {'✅' if goals_data else '❌ Missing'}
- FSRs: {'✅' if fsrs_data else '❌ Missing'}

**Required Workflow:**
1. Load HARA for [system]
2. Derive FSRs for all goals
3. Generate Word document

Please complete FSC development in FSC Developer plugin first."""
    
    try:
        # Import generator
        from generators.Functional_Safety_Concept.fsc_word_generator import FSCWordGenerator        # Create generator
        generator = FSCWordGenerator(plugin_folder)
        
        # Validate data
        is_valid, validation_warnings, errors = generator.validate_data(goals_data, fsrs_data)
        
        if not is_valid:
            error_msg = "❌ **FSC Data Validation Failed**\n\n"
            for error in errors:
                error_msg += f"- {error}\n"
            return error_msg
        
        # Calculate statistics
        stats = generator.calculate_statistics(goals_data, fsrs_data, strategies_data)
        
        # Prepare additional data
        fsc_data = {
            'allocation': cat.working_memory.get("fsc_allocation_matrix", {}),
            'mechanisms': cat.working_memory.get("fsc_safety_mechanisms", []),
            'validation': cat.working_memory.get("validation_criteria", []),
            'decomposition': cat.working_memory.get("fsc_asil_decompositions", [])
        }
        
        # Check completeness BEFORE generation
        completeness_warnings = generator.check_completeness(
            goals_data, fsrs_data, strategies_data, fsc_data
        )
        
        # Generate document
        log.info(f"📄 Generating Word document for {system_name}")
        
        doc = generator.generate(
            system_name=system_name,
            goals_data=goals_data,
            fsrs_data=fsrs_data,
            strategies_data=strategies_data,
            fsc_data=fsc_data
        )
        
        # Save document
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(plugin_folder, "generated_documents", "06_FSC")
        os.makedirs(output_dir, exist_ok=True)
        
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in system_name).replace(" ", "_")
        
        filename = f"FSC_{safe_name}_{timestamp}.docx"
        filepath = os.path.join(output_dir, filename)
        
        doc.save(filepath)
        
        # Build response
        response = f"""✅ **FSC Word Document Generated!**

**File:** `{filename}`
**Location:** `generated_documents/06_FSC/`
**Size:** ~{stats['estimated_pages']} pages (estimated)

**Content:**
📋 Safety Goals: {stats['total_goals']}
📌 FSRs: {stats['total_fsrs']}
  - Detection: {stats['fsr_by_type']['detection']}
  - Reaction: {stats['fsr_by_type']['reaction']}
  - Indication: {stats['fsr_by_type']['indication']}

**ASIL Distribution:**
{chr(10).join([f"  - ASIL {asil}: {count}" for asil, count in sorted(stats['asil_distribution'].items())])}

**Quality Metrics:**
- Coverage: {stats['coverage_pct']:.0f}% ({stats['goals_with_fsrs']}/{stats['total_goals']} goals)
- Allocation: {stats['allocation_pct']:.0f}% ({stats['allocated_fsrs']}/{stats['total_fsrs']} FSRs)
"""
        
        # Add completeness warnings prominently
        if completeness_warnings:
            response += f"\n{'='*60}\n"
            response += "\n⚠️ **DOCUMENT COMPLETENESS WARNINGS** ⚠️\n\n"
            response += "The following sections are incomplete or missing:\n\n"
            
            for warning in completeness_warnings:
                response += f"{warning}\n"
            
            response += f"\n{'='*60}\n"
            response += "\n**📋 Note:** These warnings are also included in the document on the title page.\n"
            response += "Incomplete sections are marked with ⚠️ INCOMPLETE markers in the document.\n"
        
        # Add validation warnings if any
        if validation_warnings:
            response += "\n**Data Quality Notices:**\n"
            for warning in validation_warnings[:3]:
                response += f"  ℹ️ {warning}\n"
        
        response += """
**ISO 26262-3:2018 Sections:**
1. ✅ Introduction
2. ✅ Safety Goals Overview
3. ✅ Functional Safety Requirements
4. ✅ FSR Allocation
5. ✅ Safety Mechanisms
6. ✅ ASIL Decomposition
7. ✅ Verification & Validation
8. ✅ Traceability
9. ✅ Approvals
"""
        
        if completeness_warnings:
            response += """
**⚠️ Recommended Actions:**
1. Complete the missing workflow steps listed above
2. Regenerate the document after completion
3. Review incomplete sections marked with ⚠️ in the document
"""
        else:
            response += """
**✅ Document Complete!**
All sections have been filled with available data.
"""
        
        response += """
**Next Steps:**
1. 📖 Review document in Microsoft Word
2. 👥 Share with safety team for technical review
3. ✍️ Complete Section 9 (Approvals)
4. ➡️ Proceed to Technical Safety Concept
"""
        
        return response
        
    except ImportError as e:
        log.error(f"❌ Import error: {e}")
        return f"""❌ Word generator not available.

**Error:** {str(e)}

**Solution:**
1. Ensure `fsc_word_generator.py` exists in `code/generators/`
2. Install python-docx: `pip install python-docx`
"""
        
    except Exception as e:
        log.error(f"❌ Word generation failed: {e}")
        import traceback
        log.error(traceback.format_exc())
        return f"""❌ Failed to generate Word document.

**Error:** {str(e)}

Check plugin logs for details."""


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