# hooks/output_formatter.py
from cat.mad_hatter.decorators import hook
from cat.log import log

@hook(priority=5)
def before_cat_sends_message(message, cat):
    """
    Format agent output before sending to user.
    
    Adds:
    - ISO-compliant formatting
    - File generation offers
    - Next step suggestions
    - Traceability information
    """
    
    content = message.get("content", "")
    
    # Skip empty or very short messages
    if not content or len(content.strip()) < 20:
        return message
    
    # Apply formatting based on context
    message = apply_iso_formatting(message, cat)
    message = add_file_generation_offers(message, cat)
    message = add_traceability_info(message, cat)
    message = add_next_steps(message, cat)
    
    log.info("✅ Applied output formatting")
    
    return message


def apply_iso_formatting(message, cat):
    """
    Apply ISO 26262 standard formatting conventions.
    
    - Add clause references
    - Structure with clear sections
    - Highlight requirements (shall/should)
    """
    
    content = message.get("content", "")
    
    # Check if response contains requirements
    if "shall" in content.lower() or "requirement" in content.lower():
        # Add ISO formatting markers
        content = content.replace("shall", "**shall**")
        content = content.replace("Shall", "**Shall**")
        
        # Add requirement prefix if not present
        if "**Requirement:**" not in content and "FSR-" in content:
            lines = content.split('\n')
            formatted_lines = []
            
            for line in lines:
                if line.strip().startswith("FSR-"):
                    formatted_lines.append(f"**Requirement:** {line}")
                else:
                    formatted_lines.append(line)
            
            content = '\n'.join(formatted_lines)
        
        message.text = content
    
    return message


def add_file_generation_offers(message, cat):
    """
    Offer to generate documents when FSRs or other artifacts are available.
    """
    
    last_operation = cat.working_memory.get("last_operation")
    fsrs = cat.working_memory.get("fsc_functional_requirements", [])
    goals = cat.working_memory.get("fsc_safety_goals", [])
    
    # Only add offers for relevant operations
    if not last_operation:
        return message
    
    offers = []
    
    if last_operation == "fsr_derivation" and fsrs:
        offers = [
            "📊 Excel FSR table with traceability",
            "📄 Word FSC document with all requirements",
            "📋 FSR allocation matrix"
        ]
    
    elif last_operation == "fsr_allocation" and fsrs:
        offers = [
            "📊 Excel allocation matrix",
            "📄 Word allocation report with analysis",
            "📋 Traceability matrix (SG → FSR → Component)"
        ]
    
    elif last_operation == "fsc_verification":
        offers = [
            "📄 Word FSC verification report",
            "📊 Excel compliance checklist",
            "📋 Complete FSC document package"
        ]
    
    # Add offers to message
    if offers:
        offer_text = "\n\n---\n\n💾 **Generate Documents:**\n\n"
        for offer in offers:
            offer_text += f"- {offer}\n"
        
        offer_text += "\n**Commands:**\n"
        offer_text += "- `generate FSR spreadsheet` - Generate FSR excel file\n"
        offer_text += "- `create allocation excel` - Generate allocation matrix\n"
        offer_text += "- `generate FSC document` - Generate complete Word document\n"
        
        message["content"] += offer_text
        log.info("💾 Added file generation offers")
    
    return message


def add_traceability_info(message, cat):
    """
    Add traceability information when relevant.
    """
    
    content = message.get("content", "")
    
    # Check if FSRs are mentioned
    if "FSR-" in content:
        fsrs_data = cat.working_memory.get("fsc_functional_requirements", [])
        
        if fsrs_data:
            # Count FSRs by safety goal
            goal_fsr_count = {}
            for fsr in fsrs_data:
                sg_id = fsr.get('safety_goal_id', 'Unknown')
                goal_fsr_count[sg_id] = goal_fsr_count.get(sg_id, 0) + 1
            
            # Add traceability summary
            trace_info = "\n\n---\n\n🔗 **Traceability:**\n\n"
            trace_info += f"- **Total FSRs:** {len(fsrs_data)}\n"
            trace_info += f"- **Safety Goals:** {len(goal_fsr_count)}\n"
            trace_info += f"- **Avg FSRs per Goal:** {len(fsrs_data) / len(goal_fsr_count):.1f}\n"
            
            message["content"] += trace_info
            log.info("🔗 Added traceability information")
    
    return message


def add_next_steps(message, cat):
    """
    Suggest next steps based on workflow stage.
    """
    
    current_stage = cat.working_memory.get("fsc_stage", "not_started")
    
    # Skip if response already has next steps
    if "Next Steps" in message.get("content", ""):
        return message
    
    next_steps_map = {
        "hara_loaded": [
            ("Develop Safety Strategies", "develop safety strategy for all goals"),
            ("View Safety Goals", "show HARA statistics")
        ],
        "strategies_developed": [
            ("Derive FSRs", "derive FSRs for all goals"),
            ("View Strategies", "show strategy summary")
        ],
        "fsrs_derived": [
            ("Allocate FSRs", "allocate all FSRs"),
            ("View FSRs", "show FSR summary")
        ],
        "fsrs_allocated": [
            ("Specify Validation Criteria", "specify validation criteria"),
            ("View Allocation", "show allocation summary")
        ],
        "criteria_specified": [
            ("Verify FSC", "verify FSC"),
            ("Generate Document", "generate FSC document")
        ]
    }
    
    if current_stage in next_steps_map and current_stage != "not_started":
        steps_text = "\n\n---\n\n**➡️ Next Steps:**\n\n"
        
        for step_name, command in next_steps_map[current_stage]:
            steps_text += f"- **{step_name}:** `{command}`\n"
        
        message.text += steps_text
        log.info(f"➡️ Added next steps for stage: {current_stage}")
    
    return message