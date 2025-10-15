# ==============================================================================
# AI_Agent-OutputFormatter/hook_formatter.py (UPDATED)
# Minimal OutputFormatter with FSR and Allocation file generation
# ==============================================================================

from cat.mad_hatter.decorators import hook, tool
from cat.log import log
import os
from datetime import datetime

@hook(priority=1)
def before_cat_sends_message(message, cat):
    """
    Minimal output formatter.
    
    Agent already formatted the response well via agent_prompt_hook.
    This hook only adds:
    - File generation offers
    - User format preferences
    """
    
    content = message.get("content", "")
    
    if not content or len(content.strip()) < 20:
        return message
    
    # Check user format preference
    user_format = cat.working_memory.get("output_format")
    
    if user_format == "minimal":
        # User wants plain text - strip formatting
        message = strip_markdown_formatting(message)
        log.info("📝 Applied minimal formatting")
    
    # Add file generation offers if FSC data available
    message = add_file_generation_offers(message, cat)
    
    return message


def add_file_generation_offers(message, cat):
    """
    Add file generation offers at the end of response.
    """
    
    # Check what data is available
    fsrs = cat.working_memory.get("fsc_functional_requirements", [])
    goals = cat.working_memory.get("fsc_safety_goals", [])
    last_operation = cat.working_memory.get("last_operation")
    
    # Only offer files for relevant operations
    offer_files = False
    file_types = []
    
    if last_operation == "fsr_derivation" and fsrs:
        offer_files = True
        file_types = ["Excel spreadsheet with FSR table"]
    
    elif last_operation == "fsr_allocation" and fsrs:
        offer_files = True
        file_types = [
            "Excel allocation matrix",
            "Allocation analysis report"
        ]
    
    elif last_operation == "fsc_verification":
        offer_files = True
        file_types = ["Word verification report", "Excel compliance checklist"]
    
    # Add offer if applicable
    if offer_files:
        offer = "\n\n---\n\n"
        offer += "💾 **Generate Documents:**\n"
        
        for file_type in file_types:
            offer += f"- {file_type}\n"
        
        offer += "\n**Commands:**"
        offer += "\n- `create fsr excel` - FSR listing"
        offer += "\n- `create allocation excel` - Allocation matrix"
        offer += "\n- `generate word document` - Full FSC report"
        
        message["content"] += offer
        log.info("💾 Added file generation offer")
    
    return message


def strip_markdown_formatting(message):
    """
    Strip markdown formatting for minimal output preference.
    """
    import re
    
    content = message["content"]
    
    # Remove tables (keep content as plain list)
    table_pattern = r'\|[^\n]+\|\n\|[-:\s]+\|(\n\|[^\n]+\|)+'
    
    def table_to_list(match):
        """Convert table to plain list."""
        table_text = match.group(0)
        rows = [r for r in table_text.split('\n') if r.strip() and not r.strip().startswith('|--')]
        
        # Extract cell content
        plain_items = []
        for row in rows[1:]:  # Skip header
            cells = [c.strip() for c in row.split('|') if c.strip()]
            if cells:
                plain_items.append(' '.join(cells))
        
        return '\n'.join(f"- {item}" for item in plain_items)
    
    content = re.sub(table_pattern, table_to_list, content)
    
    # Remove markdown headers (## Header → Header)
    content = re.sub(r'^#{1,6}\s+', '', content, flags=re.MULTILINE)
    
    # Remove bold (**text** → text)
    content = re.sub(r'\*\*([^*]+)\*\*', r'\1', content)
    
    # Remove horizontal rules
    content = re.sub(r'^---+$', '', content, flags=re.MULTILINE)
    
    # Clean up extra whitespace
    content = re.sub(r'\n{3,}', '\n\n', content)
    
    message["content"] = content.strip()
    
    return message


# ==============================================================================
# FILE GENERATION TOOLS - FSR
# ==============================================================================

@tool(return_direct=True)
def create_fsr_excel(tool_input, cat):
    """
    Generate Excel file with FSR listing.
    
    Includes:
    - FSR summary sheet
    - FSR details with all fields
    - Statistics
    
    Examples: "create fsr excel", "generate fsr spreadsheet"
    """
    
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    system_name = cat.working_memory.get("system_name", "System")
    
    if not fsrs_data:
        return "No FSRs available to export. Please derive FSRs first using: derive FSRs for all goals"
    
    try:
        # Import formatter
        from .fsr_formatter_xls import create_fsr_excel
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        plugin_folder = os.path.dirname(os.path.dirname(__file__))
        output_dir = os.path.join(plugin_folder, "AI_Agent-OutputFormatter", "generated_documents", "fsr")
        os.makedirs(output_dir, exist_ok=True)
        
        filename = f"FSR_{system_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        log.info(f"📊 Generating FSR Excel: {filename}")
        
        # Create Excel workbook
        wb = create_fsr_excel(fsrs_data, system_name, timestamp)
        
        if not wb:
            return "❌ Failed to generate FSR Excel - openpyxl not available"
        
        wb.save(filepath)
        
        log.info(f"✅ FSR Excel saved: {filepath}")
        
        return f"""✅ FSR Excel file generated successfully!

**System:** {system_name}
**File:** {filename}
**Location:** `generated_documents/fsr/`
**FSRs:** {len(fsrs_data)}

**Contents:**
- FSR Summary with statistics
- FSR Details (all fields per ISO 26262-3:2018, Clause 7.4.2)
- Columns: FSR ID, Description, Allocation, ASIL, Safety Goal, Verification, FTTI

**File Path:** {filepath}
"""
        
    except Exception as e:
        log.error(f"❌ FSR Excel generation failed: {e}")
        import traceback
        log.error(traceback.format_exc())
        return f"❌ Failed to generate FSR Excel file: {str(e)}\n\nPlease check the logs for details."


# ==============================================================================
# FILE GENERATION TOOLS - ALLOCATION
# ==============================================================================

@tool(return_direct=True)
def create_allocation_excel(tool_input, cat):
    """
    Generate Excel file with FSR Allocation Matrix.
    
    Per ISO 26262-3:2018, Clause 7.4.2.8
    
    Includes:
    - Allocation Matrix (FSR → Component traceability)
    - By Component view
    - By ASIL view
    - Freedom from Interference analysis
    - Validation checklist
    
    Examples: 
    - "create allocation excel"
    - "generate allocation matrix"
    - "export allocation to excel"
    """
    
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    goals_data = cat.working_memory.get("fsc_safety_goals", [])
    system_name = cat.working_memory.get("system_name", "System")
    
    if not fsrs_data:
        return "No FSRs available. Please derive and allocate FSRs first."
    
    # Check if FSRs are allocated
    allocated_count = sum(1 for f in fsrs_data if f.get('allocated_to') and f.get('allocated_to') not in ['TBD', 'NOT ALLOCATED', 'N/A'])
    
    if allocated_count == 0:
        return """⚠️ No FSRs have been allocated yet.

**Required Steps:**
1. Derive FSRs: `derive FSRs for all goals`
2. Allocate FSRs: `allocate all FSRs`
3. Then generate allocation Excel: `create allocation excel`
"""
    
    try:
        # Import allocation formatter
        from .allocation_formatter import create_allocation_excel
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        plugin_folder = os.path.dirname(os.path.dirname(__file__))
        output_dir = os.path.join(plugin_folder, "AI_Agent-OutputFormatter", "generated_documents", "allocation")
        os.makedirs(output_dir, exist_ok=True)
        
        filename = f"Allocation_{system_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        log.info(f"📊 Generating Allocation Excel: {filename}")
        
        # Create Excel workbook
        wb = create_allocation_excel(fsrs_data, goals_data, system_name, timestamp)
        
        if not wb:
            return "❌ Failed to generate Allocation Excel - openpyxl not available"
        
        wb.save(filepath)
        
        log.info(f"✅ Allocation Excel saved: {filepath}")
        
        return f"""✅ Allocation Excel file generated successfully!

**System:** {system_name}
**File:** {filename}
**Location:** `generated_documents/allocation/`
**FSRs:** {len(fsrs_data)} total, {allocated_count} allocated

**Contents per ISO 26262-3:2018, Clause 7.4.2.8:**

📋 **Sheet 1: Allocation Matrix**
   - Complete FSR → Component traceability
   - Columns: FSR ID, Description, Type, ASIL, Safety Goal, Allocated To, Component Type, Rationale, Interface

📊 **Sheet 2: By Component**
   - Component-centric view
   - Shows FSR distribution across components
   - ASIL levels per component

🎯 **Sheet 3: By ASIL**
   - ASIL integrity verification (7.4.2.8.a)
   - Allocation completeness by ASIL level

🛡️ **Sheet 4: Freedom from Interference**
   - Mixed ASIL analysis (7.4.2.8.b)
   - Risk level assessment
   - Interference considerations

✅ **Sheet 5: Validation**
   - Allocation completeness checklist
   - ISO 26262 compliance verification

**File Path:** {filepath}

**Next Steps:**
- Review allocation matrix for completeness
- Verify freedom from interference measures
- Check interface specifications (7.4.2.8.c)
"""
        
    except Exception as e:
        log.error(f"❌ Allocation Excel generation failed: {e}")
        import traceback
        log.error(traceback.format_exc())
        return f"❌ Failed to generate Allocation Excel file: {str(e)}\n\nPlease check the logs for details."


@tool(return_direct=True)
def create_allocation_report(tool_input, cat):
    """
    Generate comprehensive allocation analysis report.
    
    Text-based report with:
    - Allocation summary
    - Component analysis
    - ASIL distribution
    - Freedom from interference analysis
    - Recommendations
    
    Examples: "create allocation report", "generate allocation analysis"
    """
    
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    system_name = cat.working_memory.get("system_name", "System")
    
    if not fsrs_data:
        return "No FSRs available to analyze."
    
    try:
        # Generate analysis
        report = generate_allocation_analysis_report(fsrs_data, system_name)
        
        # Save to file
        plugin_folder = os.path.dirname(os.path.dirname(__file__))
        output_dir = os.path.join(plugin_folder, "AI_Agent-OutputFormatter", "generated_documents", "allocation")
        os.makedirs(output_dir, exist_ok=True)
        
        filename = f"Allocation_Report_{system_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        filepath = os.path.join(output_dir, filename)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(report)
        
        log.info(f"✅ Allocation report saved: {filepath}")
        
        # Return summary
        return f"""✅ Allocation Analysis Report generated!

**File:** {filename}
**Location:** `generated_documents/allocation/`

{report[:1000]}...

*[Full report saved to file]*
"""
        
    except Exception as e:
        log.error(f"❌ Report generation failed: {e}")
        return f"❌ Failed to generate report: {str(e)}"


def generate_allocation_analysis_report(fsrs_data, system_name):
    """Generate text-based allocation analysis report."""
    
    from datetime import datetime
    
    report = f"""
================================================================================
FSR ALLOCATION ANALYSIS REPORT
================================================================================

System: {system_name}
Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
ISO 26262-3:2018, Clause 7.4.2.8 - FSR Allocation to Architectural Elements

================================================================================
1. EXECUTIVE SUMMARY
================================================================================

"""
    
    # Calculate statistics
    total = len(fsrs_data)
    allocated = [f for f in fsrs_data if f.get('allocated_to') and f.get('allocated_to') not in ['TBD', 'NOT ALLOCATED', 'N/A']]
    allocated_count = len(allocated)
    unallocated_count = total - allocated_count
    
    report += f"""Total FSRs: {total}
Allocated: {allocated_count} ({allocated_count*100//total if total > 0 else 0}%)
Unallocated: {unallocated_count}

"""
    
    # Component summary
    components = set(f.get('allocated_to') for f in allocated)
    report += f"Unique Components: {len(components)}\n"
    
    # ASIL distribution
    by_asil = {}
    for fsr in fsrs_data:
        asil = fsr.get('asil', 'QM')
        by_asil[asil] = by_asil.get(asil, 0) + 1
    
    report += "\nASIL Distribution:\n"
    for asil in ['D', 'C', 'B', 'A', 'QM']:
        if asil in by_asil:
            report += f"  - ASIL {asil}: {by_asil[asil]} FSRs\n"
    
    report += """

================================================================================
2. ALLOCATION BY COMPONENT
================================================================================

"""
    
    # Group by component
    by_component = {}
    for fsr in allocated:
        comp = fsr.get('allocated_to')
        if comp not in by_component:
            by_component[comp] = []
        by_component[comp].append(fsr)
    
    for component, comp_fsrs in sorted(by_component.items(), key=lambda x: len(x[1]), reverse=True):
        asil_levels = sorted(set(f.get('asil', 'QM') for f in comp_fsrs), reverse=True)
        comp_type = comp_fsrs[0].get('allocation_type', 'Unknown')
        
        report += f"\n{component}\n"
        report += f"{'='*len(component)}\n"
        report += f"Type: {comp_type}\n"
        report += f"FSR Count: {len(comp_fsrs)}\n"
        report += f"ASIL Levels: {', '.join(asil_levels)}\n"
        report += f"FSRs: {', '.join(f.get('id', '') for f in comp_fsrs[:5])}"
        if len(comp_fsrs) > 5:
            report += f" ... (+{len(comp_fsrs)-5} more)"
        report += "\n"
    
    report += """

================================================================================
3. FREEDOM FROM INTERFERENCE ANALYSIS
================================================================================

Per ISO 26262-3:2018, Clause 7.4.2.8.b

"""
    
    # Check for mixed ASIL
    high_risk_components = []
    for component, comp_fsrs in by_component.items():
        asil_levels = set(f.get('asil', 'QM') for f in comp_fsrs)
        if len(asil_levels) > 1 and ('D' in asil_levels or 'C' in asil_levels):
            high_risk_components.append((component, asil_levels))
    
    if high_risk_components:
        report += "⚠️ HIGH RISK COMPONENTS (Mixed ASIL including C/D):\n\n"
        for comp, asils in high_risk_components:
            report += f"  - {comp}: {', '.join(sorted(asils, reverse=True))}\n"
            report += f"    → Requires spatial/temporal independence and partitioning\n\n"
    else:
        report += "✅ No high-risk ASIL mixing detected\n"
    
    report += """

================================================================================
4. RECOMMENDATIONS
================================================================================

"""
    
    if unallocated_count > 0:
        report += f"1. Complete allocation for {unallocated_count} unallocated FSRs\n"
    
    if high_risk_components:
        report += f"2. Implement freedom from interference measures for {len(high_risk_components)} high-risk components\n"
    
    report += """3. Define interface specifications per ISO 26262-3:2018, Clause 7.4.2.8.c
4. Verify ASIL integrity per ISO 26262-3:2018, Clause 7.4.2.8.a
5. Document allocation rationale for all FSRs

================================================================================
END OF REPORT
================================================================================
"""
    
    return report


# ==============================================================================
# OTHER TOOLS
# ==============================================================================

@tool(return_direct=True)
def create_word_document(tool_input, cat):
    """
    Generate Word document with complete FSC report.
    
    Examples: "create word document", "generate word file", "create FSC document"
    """
    
    fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
    goals_data = cat.working_memory.get("fsc_safety_goals", [])
    system_name = cat.working_memory.get("system_name", "System")
    
    if not fsrs_data and not goals_data:
        return "No FSC data available to export."
    
    return """⚠️ Word document generation not yet implemented.

**Available formats:**
- Excel FSR listing: `create fsr excel`
- Excel allocation matrix: `create allocation excel`
- Text allocation report: `create allocation report`

**Coming soon:**
- Complete Word FSC document per ISO 26262-3:2018, Clause 7.5
"""


@tool(return_direct=False)
def set_output_format(tool_input, cat):
    """
    Set output format preference.
    
    Options: "minimal" (plain text), "standard" (normal), "detailed" (extra info)
    
    Examples: "set output format to minimal", "format as detailed"
    """
    
    input_lower = str(tool_input).lower()
    
    if "minimal" in input_lower:
        format_type = "minimal"
    elif "detailed" in input_lower:
        format_type = "detailed"
    else:
        format_type = "standard"
    
    cat.working_memory["output_format"] = format_type
    
    return f"Output format set to: {format_type}. This will apply to future responses."