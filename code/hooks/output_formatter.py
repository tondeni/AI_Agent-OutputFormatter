# hooks/output_formatter.py
from cat.mad_hatter.decorators import hook
from cat.log import log

@hook(priority=5)
def before_cat_sends_message(message, cat):
    """
    Format agent output before sending to user.
    
    Adds:
    - ISO-compliant formatting
    - Next step suggestions
    - File generation offers
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
    log.warning(f"Content is: {message.text}")
    
    return message


def add_file_generation_offers(message, cat):
    """
    Offer to generate documents when FSRs or other artifacts are available.
    """
    
    last_operation = cat.working_memory.get("last_operation")
    fsrs = cat.working_memory.get("fsc_functional_requirements", [])
    goals = cat.working_memory.get("fsc_safety_goals", [])
    strategies = cat.working_memory.get("fsc_safety_strategies", [])
    validation_criteria = cat.working_memory.get("fsc_validation_criteria", [])
    
    # Only add offers for relevant operations
    if not last_operation:
        return message
    
    offer_text = ""
    
    # Strategy development offers
    if last_operation == "strategy_development" and strategies:
        offer_text = "\n\n📊 **Export Options:**\n"
        offer_text += "- `export strategies to excel` - Generate strategy matrix\n"
    
    # FSR derivation offers
    elif last_operation == "fsr_derivation" and fsrs:
        offer_text = "\n\nTraceability Matrix Ready (SG → FSR)!\n"
        offer_text += "\n\n📊 **Export Options:**\n"
        offer_text += "- `export FSRs to excel` - Generate traceability matrix (SG → FSR)\n"
    
    # FSR allocation offers
    elif last_operation == "fsr_allocation" and fsrs:
        offer_text = "\n\nAllocation Matrix Ready (SG → FSR → Component)!\n"
        offer_text += "📊 **Export Options:**\n"
        offer_text += "- `export FSRs allocated to excel` - Generate allocation matrix\n"
    
    # Safety mechanisms offers
    elif last_operation == "fsr_mechanisms" and fsrs:
        offer_text = "\n\nSafety Mechanisms Ready (SG → FSR → Component → Safety Mechanisms)!\n"
        offer_text += "📊 **Export Options:**\n"
        offer_text += "- `export FSRs Safety Mechanisms` - Generate safety mechanism matrix\n"
    
    # Validation criteria offers
    elif last_operation == "validation_criteria_specification" and validation_criteria:
        offer_text = "\n\nValidation Criteria Ready (SG → FSR → Component → Safety Mechanisms → Validation Criteria)!\n"
        offer_text += "\n\n📊 **Export Options:**\n"
        offer_text += "- `export validation criteria to excel` - Generate validation criteria matrix\n"
    
    # FSC verification offers
    elif last_operation == "fsc_verification":
        offer_text = "\n\n📄 **Document Generation:**\n"
        offer_text += "- `generate FSC document` - Generate complete Word document\n"
        offer_text += "- `export verification report to PDF` - Export verification report\n"
    
    if offer_text:
        message.text += offer_text
        log.info("💾 Added file generation offers")
    
    return message


def add_traceability_info(message, cat):
    """
    Add traceability information when relevant.
    """
    
    content = message.get("content", "")
    last_operation = cat.working_memory.get("last_operation")

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
            if last_operation == "fsr_derivation":
                trace_info = "\n\n---\n\n🔗 **Traceability:**\n\n"
                trace_info += f"- **Total FSRs:** {len(fsrs_data)}\n"
                trace_info += f"- **Safety Goals:** {len(goal_fsr_count)}\n"
                trace_info += f"- **Avg FSRs per Goal:** {len(fsrs_data) / len(goal_fsr_count):.1f}\n"
            
            message.text = trace_info
            log.info("🔗 Added traceability information")
    
    # Check if strategies are mentioned
    elif "strategy" in content.lower() or "strategies" in content.lower():
        strategies_data = cat.working_memory.get("fsc_safety_strategies", [])
        
        if strategies_data and "Strategy Summary" in content:
            # Already has summary, skip
            pass
    
    # Check if validation criteria are mentioned
    elif "validation criteria" in content.lower() or "acceptance criteria" in content.lower():
        validation_data = cat.working_memory.get("fsc_validation_criteria", [])
        
        if validation_data and "Validation Criteria Summary" not in content:
            # Could add brief stats if needed
            pass
    
    return message


def add_next_steps(message, cat):
    """
    Suggest next steps based on workflow stage.
    """
    
    current_stage = cat.working_memory.get("fsc_stage", "not_started")
    
    # Skip if response already has next steps
    if "Next Step" in message.get("content", "") or "next step" in message.get("content", "").lower():
        return message
    
    next_steps_map = {
        "hara_loaded": [
            ("Develop Safety Strategies for Safety Goals", "create safety strategies"),
            ("View Safety Goals", "Show HARA statistics")
        ],
        "strategies_developed": [
            ("Derive Functional Safety Requirements", "derive FSRs"),
            ("View Safety Strategies", "Show Safety Strategies summary")
        ],
        "fsrs_derived": [
            ("Allocate FSRs", "allocate FSRs to architecture"),
            ("View FSRs", "Show FSR summary")
        ],
        "fsrs_allocated": [
            ("Define Safety Mechanisms", "identify safety mechanisms"),
            ("View Allocation Matrix", "show allocation matrix")
        ],
        "mechanisms_identified": [
            ("Specify Validation Criteria", "specify validation criteria"),
            ("View Validation Criteria", "show mechanism summary")
        ],
        "validation_criteria_specified": [
            ("Verify FSC", "verify FSC"),
            ("View verification Report", "show verification report" )
        ],
        "fsc_verified": [
            ("Generate FSC Document", "generate structured FSC content"),
            ("View Generated Functiona Safety Concept", "show verification report")
        ]
    }
    
    if current_stage in next_steps_map and current_stage != "not_started":
        steps_text = "\n\n---\n\n**➡️ Next Steps:**\n\n"
        
        for step_name, command in next_steps_map[current_stage]:
            steps_text += f"- **{step_name}:** `{command}`\n"
        
        message.text += steps_text
        log.info(f"➡️ Added next steps for stage: {current_stage}")
    
    return message