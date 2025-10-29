"""
Smart Output Formatter for ISO 26262 Plugins
Routes to formatters based on working memory stage with automatic cleanup
"""

from cat.mad_hatter.decorators import hook
from cat.log import log
from typing import Dict, Optional


def is_already_formatted(text: str) -> bool:
    """
    Check if response is already well-formatted.
    Skip LLM if it contains markdown tables or structured output.
    """
    indicators = [
        '|---|',  # Markdown table separator
        '## 📋',  # Section headers
        '### 📊',  # Subsection headers
        '*ISO 26262',  # ISO references
        '✅ **Successfully',  # Success messages
        'FSR-ID | Type | ASIL',  # FSR table header
    ]
    
    return any(indicator in text for indicator in indicators)


# ============================================================================
# STAGE-BASED ROUTING WITH CLEANUP
# ============================================================================

class StageRouter:
    """Routes to appropriate formatter based on working memory stage"""
    
    # Map last_operation → formatter type
    OPERATION_TO_FORMATTER = {
        # Strategy operations
        'strategy_development': 'safety_strategies',
        'strategy_generation': 'safety_strategies',
        'strategies_developed': 'safety_strategies',
        
        # FSR operations
        'fsr_derivation': 'fsrs',
        'fsr_generation': 'fsrs',
        'fsrs_derived': 'fsrs',
        
        # Safety goal operations
        'hara_loaded': 'safety_goals',
        'safety_goals_loaded': 'safety_goals',
        'goals_derived': 'safety_goals',
        
        # Allocation operations
        'fsr_allocation': 'allocation',
        'allocation_complete': 'allocation',
        'fsrs_allocated': 'allocation',
        
        # Safety mechanism operations
        'mechanism_identification': 'mechanisms',
        'mechanisms_identified': 'mechanisms',
        
        # Validation operations
        'validation_criteria_specification': 'validation',
        'validation_criteria_specified': 'validation',
        
        # Verification operations
        'fsc_verification': 'verification',
        'fsc_verified': 'verification',
    }
    
    # Operations that already return formatted output (skip formatter)
    SKIP_FORMATTING_OPS = [
        'fsr_derivation',
        'fsr_generation',
    ]
    
    @staticmethod
    def get_formatter_type(cat) -> Optional[str]:
        """
        Determine which formatter to use based on working memory.
        
        Returns:
            formatter_type or None
        """
        
        # Check if formatting is explicitly requested
        needs_formatting = cat.working_memory.get('needs_formatting', False)
        
        if not needs_formatting:
            log.info("📍 No formatting needed (needs_formatting=False)")
            return None
        
        # Get last operation
        last_operation = cat.working_memory.get('last_operation')
        
        if not last_operation:
            log.info("📍 No formatter routing (no last_operation)")
            return None
        
        # Check if should skip
        if last_operation in StageRouter.SKIP_FORMATTING_OPS:
            log.info(f"⏭️ Skipping formatter - {last_operation} already formatted")
            return None
        
        # Get formatter type
        formatter_type = StageRouter.OPERATION_TO_FORMATTER.get(last_operation)
        
        if formatter_type:
            log.info(f"📍 Routing: {last_operation} → {formatter_type}")
        else:
            log.info(f"⚠️ No formatter for operation: {last_operation}")
        
        return formatter_type
    
    @staticmethod
    def cleanup_formatting_state(cat):
        """
        Clean up working memory after formatting to prevent re-formatting.
        
        This is CRITICAL to prevent the formatter from running on every message.
        """
        
        # Clear the formatting flag
        if 'needs_formatting' in cat.working_memory:
            del cat.working_memory['needs_formatting']
            log.info("🧹 Cleared needs_formatting flag")
        
        # Optionally clear last_operation (uncomment if you want to clear it)
        # if 'last_operation' in cat.working_memory:
        #     del cat.working_memory['last_operation']
        #     log.info("🧹 Cleared last_operation")
        
        # NOTE: We keep fsc_stage for workflow guidance


# ============================================================================
# LLM-BASED SMART FORMATTER
# ============================================================================

class SmartFormatter:
    """Uses LLM to intelligently format content into clean structures"""
    
    def __init__(self, llm_function):
        self.llm = llm_function
    
    def format_content(self, content: str, formatter_type: str, cat) -> str:
        """
        Main formatting function - uses LLM to create clean output
        
        Args:
            content: Text to format
            formatter_type: Which formatter to use (from StageRouter)
            cat: Cat instance for context
        """
        
        # Get system context
        system_name = cat.working_memory.get('system_name', 'System')
        
        # Route to appropriate formatter
        if formatter_type == 'safety_strategies':
            return self._format_strategies(content, system_name, cat)
        elif formatter_type == 'safety_goals':
            return self._format_safety_goals(content, system_name, cat)
        elif formatter_type == 'fsrs':
            return self._format_fsrs(content, system_name, cat)
        elif formatter_type == 'allocation':
            return self._format_allocation(content, system_name, cat)
        elif formatter_type == 'mechanisms':
            return self._format_mechanisms(content, system_name, cat)
        elif formatter_type == 'validation':
            return self._format_validation(content, system_name, cat)
        elif formatter_type == 'verification':
            return self._format_verification(content, system_name, cat)
        else:
            log.warning(f"⚠️ Unknown formatter type: {formatter_type}")
            return content
    
    # ========================================================================
    # FORMATTERS (same as before - keeping them for reference)
    # ========================================================================
    
    def _format_strategies(self, content: str, system_name: str, cat) -> str:
        """Format safety strategies into clean table"""
        
        prompt = f"""You are formatting ISO 26262 Safety Strategies into a professional table.

CONTENT TO FORMAT:
{content}

TASK:
Extract all safety strategies and format them as a clean markdown table with these columns:
- Safety Goal ID (e.g., SG-001)
- Strategy Type (e.g., Fault Avoidance, Fault Detection, etc.)
- Description (concise, 1-2 sentences)

OUTPUT FORMAT:
## 🎯 Safety Strategies for {system_name}

*ISO 26262-3:2018, Clause 7.4.2.3 - Safety Strategies*

| Safety Goal | Strategy Type | Description |
|-------------|---------------|-------------|
| SG-001 | Fault Avoidance | ... |
...

### Summary
- Total Safety Goals: X
- Total Strategies: Y

Only output the formatted table and summary."""

        try:
            formatted = self.llm(prompt)
            log.info("✅ LLM formatted strategies successfully")
            return "\n\n" + formatted
        except Exception as e:
            log.error(f"❌ LLM formatting failed: {e}")
            return content
    
    def _format_safety_goals(self, content: str, system_name: str, cat) -> str:
        """Format safety goals into clean table"""
        
        prompt = f"""You are formatting ISO 26262 Safety Goals into a professional table.

CONTENT TO FORMAT:
{content}

OUTPUT FORMAT:
## 🎯 Safety Goals for {system_name}

*ISO 26262-3:2018, Clause 6.4.6 - Safety Goals*

| SG-ID | Safety Goal | ASIL | Safe State | FTTI |
|-------|-------------|------|------------|------|
| SG-001 | ... | B | ... | 100ms |
...

### ASIL Distribution
- ASIL D: X goals
- ASIL C: X goals
...

Only output the table and distribution."""

        try:
            formatted = self.llm(prompt)
            log.info("✅ LLM formatted safety goals successfully")
            return "\n\n" + formatted
        except Exception as e:
            log.error(f"❌ LLM formatting failed: {e}")
            return content
    
    def _format_fsrs(self, content: str, system_name: str, cat) -> str:
        """FSR formatting (should be skipped - tool formats directly)"""
        log.warning("⚠️ FSR formatter called - this shouldn't happen!")
        return content
    
    def _format_allocation(self, content: str, system_name: str, cat) -> str:
        """Format FSR allocation matrix"""
        
        prompt = f"""Format ISO 26262 FSR Allocation into a professional matrix.

CONTENT:
{content}

OUTPUT FORMAT:
## 🗺️ FSR Allocation Matrix for {system_name}

| FSR-ID | Description | ASIL | Allocated To | Type |
|--------|-------------|------|--------------|------|
...

Only output the formatted matrix."""

        try:
            formatted = self.llm(prompt)
            log.info("✅ LLM formatted allocation successfully")
            return "\n\n" + formatted
        except Exception as e:
            log.error(f"❌ LLM formatting failed: {e}")
            return content
    
    def _format_mechanisms(self, content: str, system_name: str, cat) -> str:
        """Format safety mechanisms catalog"""
        
        prompt = f"""Format ISO 26262 Safety Mechanisms into a professional catalog.

CONTENT:
{content}

OUTPUT FORMAT:
## 🛡️ Safety Mechanisms for {system_name}

| SM-ID | Mechanism | Type | FSR Coverage | ASIL |
|-------|-----------|------|--------------|------|
...

Only output the formatted catalog."""

        try:
            formatted = self.llm(prompt)
            log.info("✅ LLM formatted mechanisms successfully")
            return "\n\n" + formatted
        except Exception as e:
            log.error(f"❌ LLM formatting failed: {e}")
            return content
    
    def _format_validation(self, content: str, system_name: str, cat) -> str:
        """Format validation criteria"""
        
        prompt = f"""Format ISO 26262 Validation Criteria.

CONTENT:
{content}

OUTPUT FORMAT:
## ✓ Validation Criteria for {system_name}

| VC-ID | FSR | Acceptance Criteria | Test Method |
|-------|-----|---------------------|-------------|
...

Only output the formatted criteria."""

        try:
            formatted = self.llm(prompt)
            log.info("✅ LLM formatted validation successfully")
            return "\n\n" + formatted
        except Exception as e:
            log.error(f"❌ LLM formatting failed: {e}")
            return content
    
    def _format_verification(self, content: str, system_name: str, cat) -> str:
        """Format verification report"""
        
        prompt = f"""Format ISO 26262 Verification Report.

CONTENT:
{content}

OUTPUT FORMAT:
## ✓ Verification Report for {system_name}

### Overall Status
...

### Verification Checklist
| Check | Status | Notes |
|-------|--------|-------|
...

Only output the formatted report."""

        try:
            formatted = self.llm(prompt)
            log.info("✅ LLM formatted verification successfully")
            return "\n\n" + formatted
        except Exception as e:
            log.error(f"❌ LLM formatting failed: {e}")
            return content


# ============================================================================
# WORKFLOW GUIDANCE
# ============================================================================

class WorkflowGuide:
    """Adds next steps and export options based on workflow stage"""
    
    NEXT_STEPS = {
        'hara_loaded': {
            'next': "develop safety strategies for all goals",
            'alternative': "develop safety strategy for SG-001"
        },
        'strategies_developed': {
            'next': "derive FSRs for all goals",
            'export': "export strategies to excel"
        },
        'fsrs_derived': {
            'next': "allocate all FSRs",
            'export': "export FSRs to excel"
        },
        'fsrs_allocated': {
            'next': "identify safety mechanisms",
            'export': "export allocation matrix"
        },
        'mechanisms_identified': {
            'next': "specify validation criteria",
            'export': "export mechanisms to excel"
        },
        'validation_criteria_specified': {
            'next': "verify FSC",
            'export': "export validation criteria"
        },
        'fsc_verified': {
            'next': "generate FSC document",
            'alternative': "create FSC excel"
        }
    }
    
    @staticmethod
    def add_guidance(content: str, cat) -> str:
        """Add workflow guidance at the end of message"""
        
        stage = cat.working_memory.get('fsc_stage')
        
        if not stage or stage not in WorkflowGuide.NEXT_STEPS:
            return content
        
        # Don't add if already present
        if '### 🚀 Next Steps' in content or '**Next Steps:**' in content:
            return content
        
        guidance = WorkflowGuide.NEXT_STEPS[stage]
        
        footer = "\n\n---\n\n"
        footer += "### 🚀 Next Steps\n\n"
        footer += f"**Recommended:** `{guidance['next']}`\n"
        
        if 'alternative' in guidance:
            footer += f"**Alternative:** `{guidance['alternative']}`\n"
        
        if 'export' in guidance:
            footer += f"**Export:** `{guidance['export']}`\n"
        
        return content + footer


# ============================================================================
# MAIN HOOK
# ============================================================================

@hook(priority=5)
def before_cat_sends_message(message, cat):
    """
    Smart formatting hook with automatic cleanup.
    
    Flow:
    1. Check if already formatted → skip
    2. Check if needs_formatting flag → route to formatter
    3. Clean up working memory after formatting
    4. Add workflow guidance
    
    CRITICAL: Cleanup prevents formatter from running on every message!
    """
    
    content = message.get("content", "")
    
    # Skip if too short or empty
    if not content or len(content.strip()) < 50:
        return message
    
    # Skip if already formatted (has tables)
    if is_already_formatted(content):
        log.info("✅ Content already formatted, skipping")
        # Still add workflow guidance
        message['content'] = WorkflowGuide.add_guidance(content, cat)
        # ✅ CLEANUP: Clear formatting flag
        StageRouter.cleanup_formatting_state(cat)
        return message
    
    try:
        # Get formatter type from working memory
        formatter_type = StageRouter.get_formatter_type(cat)
        
        if formatter_type:
            log.info(f"🎨 Formatting with {formatter_type} formatter")
            
            # Format with LLM
            formatter = SmartFormatter(cat.llm)
            formatted_content = formatter.format_content(content, formatter_type, cat)
            
            # Add workflow guidance
            formatted_content = WorkflowGuide.add_guidance(formatted_content, cat)
            
            message['content'] = formatted_content
            log.info("✅ Content formatted successfully")
        else:
            log.info("ℹ️ No formatter needed")
            # Still add workflow guidance
            message['content'] = WorkflowGuide.add_guidance(content, cat)
        
        # ✅ CLEANUP: Always clear formatting state after processing
        StageRouter.cleanup_formatting_state(cat)
        
    except Exception as e:
        log.error(f"❌ Formatting error: {e}")
        import traceback
        log.error(traceback.format_exc())
        # On error, cleanup and return original content
        StageRouter.cleanup_formatting_state(cat)
        message['content'] = WorkflowGuide.add_guidance(content, cat)
    
    return message