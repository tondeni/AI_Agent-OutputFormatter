# fsr_formatter_xls.py - Functional Safety Requirements Excel Formatter
# Generates Excel file with FSRs per ISO 26262-3:2018, Clause 7.4.2

from cat.log import log

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False
    log.warning("openpyxl not available - FSR Excel generation disabled")


def create_fsr_excel(fsrs, system_name, timestamp):
    """
    Create Excel workbook with Functional Safety Requirements.
    
    Args:
        fsrs: List of FSR dictionaries
        system_name: Name of the system
        timestamp: Timestamp for the document
        
    Returns:
        openpyxl.Workbook or None
    """
    
    if not EXCEL_AVAILABLE:
        log.warning("Cannot create FSR Excel - openpyxl not available")
        return None
    
    if not fsrs:
        log.warning("No FSRs provided to create Excel")
        return None
    
    log.info(f"Creating FSR Excel with {len(fsrs)} requirements")
    
    # Create workbook
    wb = openpyxl.Workbook()
    
    # Create sheets
    ws_summary = wb.active
    ws_summary.title = "FSR Summary"
    
    ws_details = wb.create_sheet("FSR Details")
    
    # Fill Summary Sheet
    create_summary_sheet(ws_summary, fsrs, system_name, timestamp)
    
    # Fill Details Sheet
    create_details_sheet(ws_details, fsrs, system_name, timestamp)
    
    log.info("✅ FSR Excel workbook created successfully")
    
    return wb


def create_summary_sheet(ws, fsrs, system_name, timestamp):
    """
    Create summary sheet with FSR overview.
    """
    
    # Title
    ws['A1'] = f"Functional Safety Requirements - {system_name}"
    ws['A1'].font = Font(size=16, bold=True, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws.merge_cells('A1:G1')
    
    # Metadata
    ws['A2'] = f"Generated: {timestamp}"
    ws['A3'] = f"Total FSRs: {len(fsrs)}"
    ws['A4'] = "ISO 26262-3:2018, Clause 7.4.2"
    
    # Statistics
    ws['A6'] = "FSR Statistics by ASIL:"
    ws['A6'].font = Font(bold=True)
    
    asil_counts = {}
    for fsr in fsrs:
        asil = fsr.get('asil', 'Unknown')
        asil_counts[asil] = asil_counts.get(asil, 0) + 1
    
    row = 7
    for asil in ['D', 'C', 'B', 'A', 'QM']:
        if asil in asil_counts:
            ws[f'A{row}'] = f"ASIL {asil}:"
            ws[f'B{row}'] = asil_counts[asil]
            row += 1
    
    # Type statistics
    ws[f'A{row + 1}'] = "FSR Statistics by Type:"
    ws[f'A{row + 1}'].font = Font(bold=True)
    
    type_counts = {}
    for fsr in fsrs:
        ftype = fsr.get('type', 'Unknown')
        type_counts[ftype] = type_counts.get(ftype, 0) + 1
    
    row += 2
    for ftype, count in sorted(type_counts.items()):
        ws[f'A{row}'] = f"{ftype}:"
        ws[f'B{row}'] = count
        row += 1
    
    # Set column widths
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 15


def create_details_sheet(ws, fsrs, system_name, timestamp):
    """
    Create details sheet with all FSR information.
    Columns: FSR ID, Description, ASIL, Linked-SG, Operating Modes, Preliminary Allocation, Verification Criteria
    """
    
    # Define header style
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    border_thin = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Headers
    headers = [
        "FSR ID",
        "Description",
        "ASIL",
        "Linked-SG",
        "Operating Modes",
        "Preliminary Allocation",
        "Verification Criteria"
    ]
    
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border_thin
    
    # Set column widths
    ws.column_dimensions['A'].width = 20  # FSR ID
    ws.column_dimensions['B'].width = 60  # Description
    ws.column_dimensions['C'].width = 8   # ASIL
    ws.column_dimensions['D'].width = 12  # Linked-SG
    ws.column_dimensions['E'].width = 25  # Operating Modes
    ws.column_dimensions['F'].width = 25  # Preliminary Allocation
    ws.column_dimensions['G'].width = 40  # Verification Criteria
    
    # Freeze header row
    ws.freeze_panes = 'A2'
    
    # Data rows
    row_idx = 2
    for fsr in fsrs:
        fsr_id = fsr.get('id', 'Unknown')
        description = fsr.get('description', 'N/A')
        asil = fsr.get('asil', 'N/A')
        linked_sg = fsr.get('safety_goal_id', 'N/A')
        operating_modes = fsr.get('operating_modes', 'N/A')
        allocation = fsr.get('allocated_to', 'N/A')
        verification = fsr.get('verification_criteria', 'N/A')
        
        # Write data
        ws.cell(row=row_idx, column=1).value = fsr_id
        ws.cell(row=row_idx, column=2).value = description
        ws.cell(row=row_idx, column=3).value = asil
        ws.cell(row=row_idx, column=4).value = linked_sg
        ws.cell(row=row_idx, column=5).value = operating_modes
        ws.cell(row=row_idx, column=6).value = allocation
        ws.cell(row=row_idx, column=7).value = verification
        
        # Style data cells
        for col_idx in range(1, 8):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border_thin
            
            # Color code by ASIL
            if col_idx == 3:  # ASIL column
                if asil == 'D':
                    cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                elif asil == 'C':
                    cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                elif asil == 'B':
                    cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                elif asil == 'A':
                    cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        
        row_idx += 1
    
    # Auto-filter
    ws.auto_filter.ref = f"A1:G{row_idx - 1}"
    
    log.info(f"✅ Created FSR details sheet with {row_idx - 1} requirements")


def parse_fsrs(llm_response, safety_goals):
    """
    Parse FSRs from LLM response with markdown-aware field extraction.
    Handles format: - **Field:** value
    """
    fsrs = []
    current_sg = None
    current_fsr = None
    
    lines = llm_response.split('\n')
    
    type_mapping = {
        'AVD': 'Fault Avoidance',
        'DET': 'Fault Detection',
        'CTL': 'Fault Control',
        'SST': 'Safe State Transition',
        'TOL': 'Fault Tolerance',
        'WRN': 'Warning/Indication',
        'TIM': 'Timing',
        'ARB': 'Arbitration'
    }
    
    log.info("🔍 Starting FSR parsing...")
    
    for idx, line in enumerate(lines):
        line_stripped = line.strip()
        
        # Detect safety goal section
        if '## FSRs for Safety Goal:' in line_stripped or 'FSRs for Safety Goal:' in line_stripped:
            for sg in safety_goals:
                if sg['id'] in line_stripped:
                    current_sg = sg
                    log.info(f"📍 Found section for {current_sg['id']}")
                    break
        
        # Detect FSR ID line - handle various formats
        if current_sg and ('FSR-' in line_stripped and ('**FSR-' in line_stripped or line_stripped.startswith('FSR-'))):
            # Save previous FSR
            if current_fsr:
                fsrs.append(current_fsr)
            
            # Extract FSR ID - remove markdown and extra formatting
            fsr_id = line_stripped.replace('**', '').replace('*', '').strip()
            # If line has additional text after ID, split it
            if '\n' in fsr_id or any(x in fsr_id for x in ['Description', 'ASIL', 'Operating']):
                fsr_id = fsr_id.split()[0]  # Take just first word (the ID)
            
            # Ensure proper SG prefix
            if not fsr_id.startswith('FSR-SG-'):
                # Fix FSR-001-AVD-1 to FSR-SG-001-AVD-1
                import re
                match = re.search(r'FSR-(\d+)', fsr_id)
                if match and current_sg:
                    sg_num = match.group(1)
                    fsr_id = fsr_id.replace(f'FSR-{sg_num}', f'FSR-{current_sg["id"]}')
            
            # Determine type
            fsr_type = 'General'
            for type_code, type_name in type_mapping.items():
                if f'-{type_code}-' in fsr_id:
                    fsr_type = type_name
                    break
            
            current_fsr = {
                'id': fsr_id,
                'safety_goal_id': current_sg['id'],
                'safety_goal': current_sg['description'],
                'asil': current_sg['asil'],
                'type': fsr_type,
                'description': '',
                'operating_modes': '',
                'allocated_to': '',
                'verification_criteria': '',
                'timing': current_sg.get('ftti', ''),
                'safe_state': current_sg.get('safe_state', ''),
                'emergency_operation': '',
                'functional_redundancy': ''
            }
            log.debug(f"🆕 Created FSR: {fsr_id}")
        
        # Extract FSR fields - HANDLE MARKDOWN FORMAT: - **Field:** value
        if current_fsr and line_stripped:
            # Remove leading dash/asterisk and spaces
            clean_line = line_stripped.lstrip('- ').lstrip('* ').strip()
            
            # Check if line has the **Field:** pattern
            if '**' in clean_line and ':' in clean_line:
                # Extract field name and value
                # Format: **Field Name:** value
                parts = clean_line.split('**')
                if len(parts) >= 2:
                    field_and_value = parts[1]  # Get text after first **
                    if ':' in field_and_value:
                        field_name = field_and_value.split(':')[0].strip().lower()
                        field_value = ':'.join(field_and_value.split(':')[1:]).strip()
                        
                        # Map field names to FSR properties
                        if 'description' in field_name:
                            current_fsr['description'] = field_value
                            log.debug(f"  📝 Description: {field_value[:50]}...")
                        
                        elif 'asil' in field_name and 'linked' not in field_name:
                            current_fsr['asil'] = field_value
                            log.debug(f"  🏷️ ASIL: {field_value}")
                        
                        elif 'operating mode' in field_name:
                            current_fsr['operating_modes'] = field_value
                            log.debug(f"  ⚙️ Modes: {field_value}")
                        
                        elif 'preliminary allocation' in field_name or 'allocation' in field_name:
                            current_fsr['allocated_to'] = field_value
                            log.debug(f"  📍 Allocation: {field_value}")
                        
                        elif 'verification' in field_name:
                            current_fsr['verification_criteria'] = field_value
                            log.debug(f"  ✓ Verification: {field_value[:50]}...")
    
    # Save last FSR
    if current_fsr:
        fsrs.append(current_fsr)
    
    log.info(f"✅ Parsed {len(fsrs)} FSRs from LLM response")
    
    # Debug: show sample
    if fsrs:
        log.info(f"📊 Sample FSR: {fsrs[0]['id']}")
        log.info(f"  Description: {fsrs[0]['description'][:80] if fsrs[0]['description'] else 'EMPTY!'}")
    
    return fsrs