# ==============================================================================
# code/tools/fsc_file_tools.py
# Tools for FSC file generation - Fixed imports for your folder structure
# ==============================================================================

from cat.mad_hatter.decorators import tool
from cat.log import log
import os
import sys
from datetime import datetime

# Get current file location
current_file = os.path.abspath(__file__)

# Go up the folder tree
tools_folder = os.path.dirname(current_file)           # code/tools/
code_folder = os.path.dirname(tools_folder)            # code/
plugin_folder = os.path.dirname(code_folder)           # Plugin folder/

# Add generators folder to Python path
generators_folder = os.path.join(code_folder, 'generators')
if generators_folder not in sys.path:
    sys.path.insert(0, generators_folder)
    log.info(f"Added to sys.path: {generators_folder}")


@tool(return_direct=True)
def create_word_document(tool_input, cat):
    """
    Generate Word document with complete FSC report.
    
    Creates ISO 26262-3:2018 Clause 7 compliant document with:
    - Safety goals and strategies
    - Functional Safety Requirements
    - Allocation matrix
    - Completeness warnings for missing sections
    
    Examples:
    - "create word document"
    - "generate fsc document"
    """
    
    log.info("✅ TOOL CALLED: create_word_document")
    
    # Get data from working memory
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    goals_data = cat.working_memory.get("fsc_safety_goals", [])
    strategies_data = cat.working_memory.get("fsc_safety_strategies", {})
    system_name = cat.working_memory.get("system_name", "System")
    
    if not fsrs_data and not goals_data:
        return """❌ No FSC data available to export.

**Required:** Complete the FSC development workflow:
1. Load HARA for [system name]
2. Develop safety strategies for all goals
3. Derive FSRs for all goals

Then generate the Word document."""
    
    try:
        # Import from Functional_Safety_Concept subfolder
        from generators.Functional_Safety_Concept import FSCWordGenerator
        
        log.info("✅ Successfully imported FSCWordGenerator")
        
        # Setup output directory
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(plugin_folder, "generated_documents", "06_FSC")
        os.makedirs(output_dir, exist_ok=True)
        
        # Sanitize system name for filename
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in system_name).replace(" ", "_")
        
        filename = f"FSC_{safe_name}_{timestamp}.docx"
        filepath = os.path.join(output_dir, filename)
        
        # Create generator instance
        log.info(f"📄 Creating FSC Word generator for: {system_name}")
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
        
        # Prepare additional FSC data
        fsc_data = {
            'allocation': cat.working_memory.get("fsc_allocation_matrix", {}),
            'mechanisms': cat.working_memory.get("fsc_safety_mechanisms", []),
            'validation': cat.working_memory.get("validation_criteria", []),
            'decomposition': cat.working_memory.get("fsc_asil_decompositions", [])
        }
        
        # Check completeness
        completeness_warnings = generator.check_completeness(
            goals_data, fsrs_data, strategies_data, fsc_data
        )
        
        # Generate document
        log.info(f"📄 Generating Word document: {filename}")
        doc = generator.generate(
            system_name=system_name,
            goals_data=goals_data,
            fsrs_data=fsrs_data,
            strategies_data=strategies_data,
            fsc_data=fsc_data
        )
        
        # Save document
        doc.save(filepath)
        log.info(f"✅ Document saved: {filepath}")
        
        # Build response
        response = f"""✅ **FSC Word Document Generated Successfully!**

**File:** `{filename}`
**Location:** `generated_documents/06_FSC/`
**Size:** ~{stats['estimated_pages']} pages (estimated)

**Content Summary:**
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
        
        # Add completeness warnings
        if completeness_warnings:
            response += f"\n{'='*60}\n"
            response += "\n⚠️ **DOCUMENT COMPLETENESS WARNINGS** ⚠️\n\n"
            response += "The following sections are incomplete:\n\n"
            
            for warning in completeness_warnings:
                response += f"{warning}\n"
            
            response += f"\n{'='*60}\n"
            response += "\n**📋 Note:** These warnings are on the title page of the document.\n"
            response += "Incomplete sections are marked with ⚠️ INCOMPLETE.\n"
            
            response += """
**⚠️ Recommended Actions:**
1. Complete the missing workflow steps listed above
2. Regenerate the document
3. Review incomplete sections marked with ⚠️
"""
        else:
            response += """
**✅ Document Complete!**
All sections have been filled with available data.
"""
        
        # Add validation warnings
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

**Next Steps:**
1. 📖 Review document in Microsoft Word
2. 👥 Share with safety team
3. ✍️ Complete approvals (Section 9)
4. ➡️ Proceed to Technical Safety Concept
"""
        
        return response
        
    except ImportError as e:
        log.error(f"❌ Import error: {e}")
        import traceback
        log.error(traceback.format_exc())
        
        return f"""❌ Word generator not available.

**Error:** {str(e)}

**Debug Info:**
- Current file: {current_file}
- Tools folder: {tools_folder}
- Code folder: {code_folder}
- Plugin folder: {plugin_folder}
- Generators folder: {generators_folder}
- Generators in sys.path: {generators_folder in sys.path}

**Expected file location:**
`{os.path.join(generators_folder, 'Functional_Safety_Concept', 'fsc_word_generator.py')}`

**File exists:** {os.path.exists(os.path.join(generators_folder, 'Functional_Safety_Concept', 'fsc_word_generator.py'))}

**Solution:**
1. Verify file exists at expected location
2. Check __init__.py files exist in:
   - code/generators/__init__.py
   - code/generators/Functional_Safety_Concept/__init__.py
3. Install python-docx: pip install python-docx
"""
        
    except Exception as e:
        log.error(f"❌ Word generation failed: {e}")
        import traceback
        log.error(traceback.format_exc())
        
        return f"""❌ Failed to generate Word document.

**Error:** {str(e)}

**Troubleshooting:**
1. Check plugin logs for details
2. Verify FSC data in working memory
3. Ensure write permissions for generated_documents/
"""


@tool(return_direct=True)
def create_excel_file(tool_input, cat):
    """
    Generate Excel file with FSC data (FSRs, allocation, traceability).
    
    Creates Excel workbook with multiple sheets:
    - FSRs listing
    - Allocation matrix
    - Traceability
    - Statistics
    
    Examples: 
    - "create excel file"
    - "generate excel spreadsheet"
    """
    
    log.info("✅ TOOL CALLED: create_excel_file")
    
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    goals_data = cat.working_memory.get("fsc_safety_goals", [])
    system_name = cat.working_memory.get("system_name", "System")
    
    if not fsrs_data:
        return """❌ No FSC data available to export.

**Required:** Derive FSRs first using: `derive FSRs for all goals`

Then generate Excel file."""
    
    try:
        # Import from Functional_Safety_Concept subfolder
        from Functional_Safety_Concept.fsc_excel_generator import FSCExcelGenerator
        
        log.info("✅ Successfully imported FSCExcelGenerator")
        
        # Setup output directory
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(plugin_folder, "generated_documents", "06_FSC")
        os.makedirs(output_dir, exist_ok=True)
        
        # Sanitize system name
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in system_name).replace(" ", "_")
        
        filename = f"FSC_{safe_name}_{timestamp}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        # Create generator
        log.info(f"📊 Creating FSC Excel generator for: {system_name}")
        generator = FSCExcelGenerator()
        
        # Validate
        is_valid, warnings, errors = generator.validate_data(goals_data, fsrs_data)
        
        if not is_valid:
            return "❌ Data validation failed: " + "; ".join(errors)
        
        # Calculate statistics
        stats = generator.calculate_statistics(goals_data, fsrs_data)
        
        # Generate workbook
        log.info(f"📊 Generating Excel file: {filename}")
        
        wb = generator.generate(
            system_name=system_name,
            goals_data=goals_data,
            fsrs_data=fsrs_data,
            strategies_data=cat.working_memory.get("fsc_safety_strategies", {}),
            allocation_data=cat.working_memory.get("fsc_allocation_matrix", {})
        )
        
        # Save workbook
        wb.save(filepath)
        log.info(f"✅ Excel saved: {filepath}")
        
        # Build response
        asil_summary = ', '.join([f"ASIL {asil}: {count}" for asil, count in sorted(stats['asil_distribution'].items())])
        
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

**ASIL Distribution:** {asil_summary}

Ready for filtering, sorting, and analysis in Excel!"""
        
    except ImportError as e:
        log.error(f"❌ Import error: {e}")
        import traceback
        log.error(traceback.format_exc())
        
        return f"""❌ Excel generator not available.

**Error:** {str(e)}

**Debug Info:**
- Generators folder: {generators_folder}
- Expected file: {os.path.join(generators_folder, 'Functional_Safety_Concept', 'fsc_excel_generator.py')}
- File exists: {os.path.exists(os.path.join(generators_folder, 'Functional_Safety_Concept', 'fsc_excel_generator.py'))}

**Solution:**
1. Verify fsc_excel_generator.py exists
2. Check __init__.py files
3. Install openpyxl: pip install openpyxl
"""
        
    except Exception as e:
        log.error(f"❌ Excel generation failed: {e}")
        import traceback
        log.error(traceback.format_exc())
        
        return f"❌ Failed to generate Excel: {str(e)}"


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
    """
    
    log.info("✅ TOOL CALLED: generate_fsc_files")
    
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    
    if not fsrs_data:
        return """❌ No FSC data available.

**Required:** Derive FSRs first: `derive FSRs for all goals`"""
    
    # Generate Word
    word_result = create_word_document("", cat)
    
    # Generate Excel
    excel_result = create_excel_file("", cat)
    
    # Combine
    return f"""✅ **Complete FSC Documentation Package Generated!**

{'='*70}

### 📄 Word Document

{word_result}

{'='*70}

### 📊 Excel Workbook

{excel_result}

{'='*70}

**Next Steps:**
1. Review files in `generated_documents/06_FSC/`
2. Share with safety team
3. Include in ISO 26262 documentation
"""


@tool(return_direct=True)
def check_fsc_generators(tool_input, cat):
    """
    Diagnostic tool to verify FSC generators are accessible.
    
    Checks:
    - Folder structure
    - File existence
    - Import capability
    
    Use this to troubleshoot import errors.
    
    Example: "check fsc generators"
    """
    
    log.info("✅ TOOL CALLED: check_fsc_generators")
    
    report = "🔍 **FSC Generator Diagnostic Report**\n\n"
    
    # Check folder structure
    report += "**Folder Structure:**\n"
    report += f"✅ Plugin folder: `{plugin_folder}`\n" if os.path.exists(plugin_folder) else "❌ Plugin folder not found\n"
    report += f"✅ Code folder: `{code_folder}`\n" if os.path.exists(code_folder) else "❌ Code folder not found\n"
    report += f"✅ Tools folder: `{tools_folder}`\n" if os.path.exists(tools_folder) else "❌ Tools folder not found\n"
    report += f"✅ Generators folder: `{generators_folder}`\n" if os.path.exists(generators_folder) else "❌ Generators folder not found\n"
    
    fsc_folder = os.path.join(generators_folder, 'Functional_Safety_Concept')
    report += f"✅ FSC subfolder: `{fsc_folder}`\n" if os.path.exists(fsc_folder) else "❌ FSC subfolder not found\n"
    
    # Check files
    report += "\n**Required Files:**\n"
    
    files = {
        'generators __init__': os.path.join(generators_folder, '__init__.py'),
        'FSC __init__': os.path.join(fsc_folder, '__init__.py'),
        'fsc_word_generator.py': os.path.join(fsc_folder, 'fsc_word_generator.py'),
        'fsc_excel_generator.py': os.path.join(fsc_folder, 'fsc_excel_generator.py'),
        'fsr_excel_generator.py': os.path.join(fsc_folder, 'fsr_excel_generator.py'),
    }
    
    all_exist = True
    for name, path in files.items():
        exists = os.path.exists(path)
        report += f"{'✅' if exists else '❌'} {name}\n"
        if not exists:
            all_exist = False
    
    # Check sys.path
    report += f"\n**Python Path:**\n"
    report += f"{'✅' if generators_folder in sys.path else '⚠️'} Generators in sys.path\n"
    
    # Try imports
    report += "\n**Import Tests:**\n"
    
    try:
        from generators.Functional_Safety_Concept import FSCWordGenerator
        report += "✅ FSCWordGenerator imported successfully\n"
    except Exception as e:
        report += f"❌ FSCWordGenerator import failed: {e}\n"
    
    try:
        from generators.Functional_Safety_Concept import FSCExcelGenerator
        report += "✅ FSCExcelGenerator imported successfully\n"
    except Exception as e:
        report += f"❌ FSCExcelGenerator import failed: {e}\n"
    
    # Overall status
    report += "\n**Status:**\n"
    if all_exist:
        report += "✅ All files exist - Ready to generate documents!\n"
    else:
        report += "❌ Some files missing - Check file locations\n"
    
    return report