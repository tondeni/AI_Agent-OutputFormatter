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

    indicators = []

    # indicators = [
    #     '|---|',  # Markdown table separator
    #     '## 📋',  # Section headers
    #     '### 📊',  # Subsection headers
    #     '*ISO 26262',  # ISO references
    #     '✅ **Successfully',  # Success messages
    #     'FSR-ID | Type | ASIL',  # FSR table header
    # ]
    
    return any(indicator in text for indicator in indicators)


# ============================================================================
# STAGE-BASED ROUTING WITH CLEANUP
# ============================================================================

class StageRouter:
    """Routes to appropriate formatter based on working memory stage"""
    
    # Map last_operation → formatter type
    OPERATION_TO_FORMATTER = {
        
        # HARA OPERATIONS
        'function_extraction': 'hara_functions',
        'hazop_analysis': 'hara_hazop',
        'operational_situations_defined': 'hara_situations',
        'esc_assessment_complete': 'hara_esc',
        'asil_determination_complete': 'hara_asil',
        'safety_goals_derived': 'hara_safety_goals',
        
        # FSC OPERATIONS
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
        
        # Route to appropriate HARA formatter
        if formatter_type == 'hara_functions':
            return self._format_hara_functions(content, cat)
        elif formatter_type == 'hara_hazop':
            return self._format_hara_hazop(content, cat)
        elif formatter_type == 'hara_situations':
            return self._format_operational_situations(content, cat)
        elif formatter_type == 'hara_esc':
            return self._format_hara_esc(content, cat)
        elif formatter_type == 'hara_asil':
            return self._format_hara_asil(content, cat)
        elif formatter_type == 'hara_safety_goals':
            return self._format_hara_safety_goals(content, cat)
        elif formatter_type == 'safety_strategies':
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

    def _format_hara_functions(self, content: str, cat) -> str:
        """Format extracted HARA functions with criticality analysis"""
    
        log.warning( " ---------------------- FORMAT HARA FUNCTION --------------------------")
        # Get data from working memory (try both keys for compatibility)
        system_name = cat.working_memory.get('hara_item_name', 
                                        cat.working_memory.get('system_name', 'System'))
        functions = cat.working_memory.get('item_functions', [])
    
        if not functions:
            log.warning("No functions data in working memory")
            return content
    
        # Build systematic output
        output = f"## 🔧 Safety-Relevant Functions: {system_name}\n\n"
        output += "*ISO 26262-3:2018, Clause 5 - Item Definition*\n\n"
        
        # Parse functions (handle both string and list formats)
        if isinstance(functions, str):
            func_list = [l.strip() for l in functions.split('\n') if l.strip()]
        else:
            func_list = functions

        enhancement_prompt = f"""Format ISO 26262 Item: {system_name} functions into a professional matrix.

FUNCTIONS CONTENT:
{func_list}

OUTPUT FORMAT:
## 🗺️ Functions extracted from Item Definition of {system_name}

| Function-ID | Name | Description |
|--------|-------------|------|
...

Only output the formatted matrix."""
            
        output = cat.llm(enhancement_prompt)        
        return output

    def _format_hara_hazop(self, content: str, cat) -> str:
        """
        Format HAZOP analysis results with MORE details.

        Includes severity rationale and can show more hazards.
        """

        # Get complete output
        complete_output = cat.working_memory.get('hazop_complete_output', None)

        if not complete_output or complete_output.get('status') != 'success':
            log.warning("No complete HAZOP output available")
            return content

        # Extract data
        system_name = complete_output['system_name']
        hazop_results = complete_output['hazards']
        stats = complete_output['statistics']
        timestamp = complete_output.get('timestamp', '')

        # Build output
        output = f"## ⚠️ HAZOP Analysis Report: {system_name}\n\n"
        output += f"*{complete_output['iso_standard']}, Clause {complete_output['clause']}*\n\n"
        output += f"**Analysis Date:** {timestamp}\n\n"

        # Summary statistics
        output += "### 📊 Summary\n\n"
        output += f"- **Total Hazards:** {stats['total_hazards']}\n"
        output += f"- **Functions Analyzed:** {stats['functions_analyzed']}\n\n"

        # Severity distribution with visual indicators
        output += "### Severity Distribution\n\n"
        sev_dist = stats['severity_distribution']
        sev_icons = {
            'S3': '🔴',  # Critical
            'S2': '🟠',  # High
            'S1': '🟡',  # Medium
            'S0': '🟢'   # Low
        }

        for sev in ['S3', 'S2', 'S1', 'S0']:
            count = sev_dist.get(sev, 0)
            if count > 0:
                icon = sev_icons.get(sev, '⚪')
                percentage = (count / stats['total_hazards'] * 100)
                bar = '█' * int(percentage / 5)  # Simple bar chart
                output += f"{icon} **{sev}**: {count} hazards ({percentage:.1f}%) {bar}\n"

        output += "\n"

        # Group by function
        by_function = {}
        for hazard in hazop_results:
            func_key = f"{hazard['function_id']}: {hazard['function_name']}"
            if func_key not in by_function:
                by_function[func_key] = []
            by_function[func_key].append(hazard)

        # Detailed tables by function
        output += "### 🔍 Detailed Hazard Analysis\n\n"

        for func_key, hazards in by_function.items():
            # output += f"#### {func_key}\n\n"

            # Expanded table with rationale
            output += "| ID | Guide Word | Behavior | Event | Severity | Rationale |\n"
            output += "|:---|:-----------|:---------|:------|:--------:|:----------|\n"

            for h in hazards:
                output += f"| {h['hazard_id']} "
                output += f"| {h['guide_word']} "
                output += f"| {h['malfunctioning_behavior'][:40]}... "
                output += f"| {h['hazardous_event'][:40]}... "
                output += f"| **{h['severity']}** "
                output += f"| {h['severity_rationale'][:40]}... |\n"

            output += f"\n**Total hazards for this function:** {len(hazards)}\n\n"

        # Compliance notes
        if 'compliance_notes' in complete_output:
            output += "### ✅ Compliance Notes\n\n"
            for note in complete_output['compliance_notes']:
                output += f"- {note}\n"
            output += "\n"

        # Next steps
        if 'next_steps' in complete_output:
            output += "### 📋 Recommended Next Steps\n\n"
            for step in complete_output['next_steps']:
                output += f"- {step}\n"

        return output

    def _format_hara_hazop_old(self, content: str, cat) -> str:
        """Format HAZOP analysis results grouped by function"""
    
        system_name = cat.working_memory.get('hara_item_name', 
                                        cat.working_memory.get('system_name', 'System'))
        hazop_results = cat.working_memory.get('hazop_results', [])

        if not hazop_results:
            log.warning("No HAZOP results in working memory")
            return content
    
        output = f"## ⚠️ HAZOP Analysis Results: {system_name}\n\n"
        output += "*ISO 26262-3:2018, Clause 6.4.3 - Hazard Identification*\n\n"
        
        # # Parse functions (handle both string and list formats)
        # if isinstance(hazop_results, str):
        #     func_list = [l.strip() for l in functions.split('\n') if l.strip()]
        # else:
        #     func_list = functions

        enhancement_prompt = f"""Format Hazop analysis result : {hazop_results} into a professional matrix.


OUTPUT FORMAT:
## 🗺️ Functions extracted from Item Definition of {system_name}

| Hazard ID	| Function | Guide Word	| Malfunctioning Behavior | Hazardous Event	| Severity |
|-----------|----------|------      |-------------------------|-----------------|----------|
...

Only output the formatted matrix."""
            
        output = cat.llm(enhancement_prompt)
        return output

    def _format_operational_situations(self, content: str, cat) -> str:
        """
        Format Operational Situations analysis results into a markdown report.

        This function is modeled after _format_hara_hazop and expects to find
        a 'situations_complete_output' dictionary in cat.working_memory.
        """

        # Get complete output from working memory
        # We assume a structure similar to HAZOP, where all individual
        # components (scenarios, stats, etc.) are bundled.
        complete_output = cat.working_memory.get('operational_situations_complete_output', None)

        if not complete_output:
            log.warning("No 'situations_complete_output' dictionary found in working memory.")
            return content

        # Check for a status, if your workflow uses one
        if complete_output.get('status') != 'success':
            log.warning("Situations output is not ready or failed.")
            return content

        # --- Extract data using the keys you provided ---
        # (via the 'situations_complete_output' dictionary)
        system_name = complete_output.get('system_name', 'N/A')
        scenarios = complete_output.get('scenarios', []) # Renamed from 'operational_situations' for clarity
        stats = complete_output.get('statistics', {})
        timestamp = complete_output.get('timestamp', 'N/A')
        iso_standard = complete_output.get('iso_standard', 'N/A')
        clause = complete_output.get('clause', 'N/A')
        compliance_notes = complete_output.get('compliance_notes', [])
        next_steps = complete_output.get('next_steps', [])

        # --- Build Markdown Report ---
        output = f"## 🏙️ Operational Situations Analysis for {system_name}\n\n"
        output += f"*{iso_standard}, Clause {clause}*\n\n"
        output += f"**Analysis Date:** {timestamp}\n\n"

        # --- Summary statistics ---
        # We guess the key 'total_scenarios' based on 'total_hazards'
        # Use len(scenarios) as a fallback.
        total_scenarios = stats.get('total_scenarios', len(scenarios))
        output += "### 📊 Summary\n\n"
        output += f"- **Total Scenarios Identified:** {total_scenarios}\n"
        # Add any other stats you have, e.g.:
        # output += f"- **Operating Modes Analyzed:** {stats.get('modes_analyzed', 'N/A')}\n"
        output += "\n"

        # --- Detailed Scenarios Table ---
        output += "### 🔍 Detailed Operational Scenarios\n\n"

        if not scenarios:
            output += "No scenarios were identified.\n\n"
        else:
            # --- !! ASSUMPTION !! ---
            # You MUST update these headers and row keys to match your
            # 'scenario' object structure.
            # I've guessed some common keys for an operational scenario.
            output += "| ID | Scenario | Exposure | Duration % | description |\n"
            output += "|:---|:---------|:---------|:---------- |:------------|\n"

            for s in scenarios:
                # Use .get() for safety, as we are guessing keys
                s_id = s.get('scenario_id', 'N/A')
                s_nam = s.get('name', 'No description') # Truncate
                s_exp = s.get('exposure_class', 'N/A')
                s_dur = s.get('duration_percentage', 'N/A')
                s_desc = s.get('description', 'N/A')

                output += f"| {s_id} "
                output += f"| {s_nam} "
                output += f"| {s_exp} "
                output += f"| {s_dur} "
                output += f"| {s_desc} |\n"

            output += f"\n**Total scenarios:** {len(scenarios)}\n\n"

        # --- Compliance notes ---
        if compliance_notes:
            output += "### ✅ Compliance Notes\n\n"
            for note in compliance_notes:
                output += f"- {note}\n"
            output += "\n"

        # --- Next steps ---
        if next_steps:
            output += "### 📋 Recommended Next Steps\n\n"
            for step in next_steps:
                output += f"- {step}\n"
            output += "\n"

        return output

    def _format_hara_esc(self, content: str, cat) -> str:
        """
        Format E/S/C Assessment results into a markdown report.

        This function expects to find an 'esc_assessment_complete_output'
        """

        # Get complete output from working memory
        complete_output = cat.working_memory.get('esc_assessment_complete_output', None)

        if not complete_output:
            log.warning("No 'esc_assessment_complete_output' dictionary found in working memory.")
            return content

        # Check status
        if complete_output.get('status') != 'success':
            log.warning("ESC assessment output is not ready or failed.")
            return content

        # --- Extract data ---
        system_name = complete_output.get('system_name', 'N/A')
        hazards = complete_output.get('hazards', [])
        stats = complete_output.get('statistics', {})
        timestamp = complete_output.get('timestamp', 'N/A')
        iso_standard = complete_output.get('iso_standard', 'N/A')
        clause = complete_output.get('clause', 'N/A')
        compliance_notes = complete_output.get('compliance_notes', [])
        next_steps = complete_output.get('next_steps', [])

        # # --- Build Markdown Report ---
        # output = f"## 🎯 HARA - E/S/C Assessment for {system_name}\n\n"
        # output += f"*{iso_standard}, {clause}*\n\n"
        # output += f"**Assessment Date:** {timestamp}\n\n"

        # # --- Summary Statistics ---
        # total_hazards = stats.get('total_hazards', len(hazards))
        # severity_dist = stats.get('severity_distribution', {})
        # exposure_dist = stats.get('exposure_distribution', {})
        # controllability_dist = stats.get('controllability_distribution', {})

        # output += "### 📊 Summary Statistics\n\n"
        # output += f"- **Total Hazards Assessed:** {total_hazards}\n\n"

        # # Severity Distribution
        # output += "**Severity Distribution:**\n"
        # output += f"- S3 (Life-threatening/Fatal): {severity_dist.get('S3', 0)}\n"
        # output += f"- S2 (Severe injuries): {severity_dist.get('S2', 0)}\n"
        # output += f"- S1 (Light/Moderate injuries): {severity_dist.get('S1', 0)}\n"
        # output += f"- S0 (No injuries): {severity_dist.get('S0', 0)}\n\n"

        # # Exposure Distribution
        # output += "**Exposure Distribution:**\n"
        # output += f"- E4 (High probability ≥10%): {exposure_dist.get('E4', 0)}\n"
        # output += f"- E3 (Medium probability 1-10%): {exposure_dist.get('E3', 0)}\n"
        # output += f"- E2 (Low probability 0.1-1%): {exposure_dist.get('E2', 0)}\n"
        # output += f"- E1 (Very low probability): {exposure_dist.get('E1', 0)}\n"
        # output += f"- E0 (Incredibly unlikely): {exposure_dist.get('E0', 0)}\n\n"

        # # Controllability Distribution
        # output += "**Controllability Distribution:**\n"
        # output += f"- C3 (Difficult/Uncontrollable <90%): {controllability_dist.get('C3', 0)}\n"
        # output += f"- C2 (Normally controllable ≥90%): {controllability_dist.get('C2', 0)}\n"
        # output += f"- C1 (Simply controllable ≥99%): {controllability_dist.get('C1', 0)}\n"
        # output += f"- C0 (Controllable in general >99%): {controllability_dist.get('C0', 0)}\n\n"

        # --- Summary Table ---
        output = "### 📋 HARA Table Summary\n\n"

        if not hazards:
            output += "No hazards were assessed.\n\n"
        else:
            output += "| ID | Hazardous Event | E | S | C | Scenario |\n"
            output += "|:---|:----------------|:-:|:-:|:-:|:---------|\n"

            for h in hazards:
                h_id = h.get('hazard_id', 'N/A')
                h_event = h.get('hazardous_event', 'N/A')
                h_exp = h.get('exposure', {}).get('rating', 'E?')
                h_sev = h.get('severity', {}).get('rating', 'S?')
                h_con = h.get('controllability', {}).get('rating', 'C?')
                h_scenario = h.get('driving_scenario', 'N/A')

                output += f"| {h_id} "
                output += f"| {h_event} "
                output += f"| {h_exp} "
                output += f"| {h_sev} "
                output += f"| {h_con} "
                output += f"| {h_scenario} |\n"

            output += f"\n**Total hazards:** {len(hazards)}\n\n"

        # --- Compliance Notes ---
        if compliance_notes:
            output += "### ✅ ISO 26262 Compliance\n\n"
            for note in compliance_notes:
                output += f"- {note}\n"
            output += "\n"

        # --- Next Steps ---
        if next_steps:
            output += "### 📋 Recommended Next Steps\n\n"
            for step in next_steps:
                output += f"- {step}\n"
            output += "\n"

        return output

    def _format_hara_asil(self, content: str, cat) -> str:
        """Format ASIL determination with distribution and a complete summary table"""

        system_name = cat.working_memory.get('hara_item_name', 
                                             cat.working_memory.get('system_name', 'System'))
        hara_table = cat.working_memory.get('complete_hara_table', [])

        if not hara_table:
            log.warning("No HARA table found in working memory")
            return "No HARA data found. Please run HAZOP and E/S/C assessment first."

        # Check if ASIL determination has been run
        if 'asil' not in hara_table[0]:
            log.warning("HARA table found, but ASIL data is missing.")
            return "ASIL data is missing. Please run `determine asil` first."

        output = f"## 🎯 ASIL Determination: {system_name}\n\n"
        output += "*ISO 26262-3:2018, Clause 6.4.5 & Table 4 - ASIL Determination*\n\n"

        # --- ASIL Distribution ---
        asil_dist = {}
        for h in hara_table:
            asil = h.get('asil', 'QM')
            asil_dist[asil] = asil_dist.get(asil, 0) + 1

        output += "### 📊 ASIL Distribution\n\n"

        total = len(hara_table)
        # Add 'INVALID' in case of parsing errors
        for asil in ['D', 'C', 'B', 'A', 'QM', 'INVALID']: 
            count = asil_dist.get(asil, 0)
            if count > 0:
                pct = (count / total * 100) if total > 0 else 0
                bar = "█" * int(pct / 4) # Create a small bar chart
                output += f"- **ASIL {asil}**: {count} hazards ({pct:.1f}%) {bar}\n"

        output += f"\n**Total Hazards:** {len(hara_table)}\n"

        # --- Complete HARA Table with ASIL ---
        output += "\n### 📋 Complete HARA Table (Sorted by ASIL)\n\n"
        
        # Create table header
        output += "| ID | Hazardous Event | S | E | C | **ASIL** |\n"
        output += "|:---|:---|:---:|:---:|:---:|:---:|\n" # Table alignment

        # Sort table by ASIL to show most critical first
        asil_order = {'D': 0, 'C': 1, 'B': 2, 'A': 3, 'QM': 4, 'INVALID': 5}
        sorted_table = sorted(
            hara_table, 
            key=lambda h: asil_order.get(h.get('asil', 'QM'), 99)
        )

        # Populate table rows
        for h in sorted_table:
            haz_id = h.get('id', '???')
            event = h.get('hazardous_event', '???')
            
            s = h.get('severity', '?')
            e = h.get('exposure', '?')
            c = h.get('controllability', '?')
            asil = h.get('asil', '???')

            # Highlight the most critical hazards
            if asil == 'D':
                output += f"| **{haz_id}** | {event} | **{s}** | **{e}** | **{c}** | **ASIL {asil}** |\n"
            elif asil == 'C':
                output += f"| {haz_id} | {event} | {s} | {e} | {c} | **{asil}** |\n"
            else:
                output += f"| {haz_id} | {event} | {s} | {e} | {c} | {asil} |\n"
        
        output += "\n---\n"
        output += "### 🚀 Next Steps\n"
        output += "1. Review the full ASIL table above.\n"
        output += "2. Proceed to derive safety goals: `derive safety goals`\n"

        return output

    def _format_hara_safety_goals(self, content: str, cat) -> str:
        """
        Format derived safety goals into a complete markdown report.

        This function expects to find a 'safety_goals_complete_output'
        dictionary in cat.working_memory.
        """

        # Get complete output from working memory
        complete_output = cat.working_memory.get('safety_goals_complete_output', None)

        if not complete_output:
            log.warning("No 'safety_goals_complete_output' dictionary found in working memory.")
            # Check for a 'no ASIL hazards' warning
            if "No ASIL-rated hazards found" in content:
                log.info("Formatting 'No ASIL-rated hazards' warning.")
                return (
                    "## 🎯 Safety Goals\n\n"
                    "*{iso_standard}, Clause {clause}*\n\n"
                    "**Analysis complete: No safety goals are required.**\n\n"
                    "**Rationale:** {message}\n"
                    "- Total Hazards Analyzed: {details.get('total_hazards_analyzed', 'N/A')}\n"
                    "- QM-Rated Hazards: {details.get('qm_rated_hazards', 'N/A')}\n\n"
                    "Per ISO 26262, hazards rated QM (Quality Management) do not require ASIL decomposition or dedicated safety goals.\n\n"
                    "### 📋 Recommended Next Steps\n\n"
                    "- **Review** the HARA E/S/C and ASIL assessments to confirm all ratings are correct: `show asil ratings`\n"
                    "- If ratings are correct, the HARA is complete.\n"
                ).format(
                    iso_standard=complete_output.get('iso_standard', 'ISO 26262-3:2018'),
                    clause=complete_output.get('clause', '6.4.6'),
                    message=complete_output.get('message', 'No ASIL-rated (A, B, C, D) hazards were found.'),
                    details=complete_output.get('details', {})
                )
            log.error("Safety goals output missing and not a 'no goals' warning.")
            return content

        # Check status
        if complete_output.get('status') != 'success':
            # Handle the 'warning' case (no ASIL hazards)
            if complete_output.get('status') == 'warning':
                log.info("Formatting 'No ASIL-rated hazards' warning.")
                return (
                    "## 🎯 Safety Goals\n\n"
                    "*{iso_standard}, Clause {clause}*\n\n"
                    "**Analysis complete: No safety goals are required.**\n\n"
                    "**Rationale:** {message}\n"
                    "- Total Hazards Analyzed: {details.get('total_hazards_analyzed', 'N/A')}\n"
                    "- QM-Rated Hazards: {details.get('qm_rated_hazards', 'N/A')}\n\n"
                    "Per ISO 26262, hazards rated QM (Quality Management) do not require ASIL decomposition or dedicated safety goals.\n\n"
                    "### 📋 Recommended Next Steps\n\n"
                    "- **Review** the HARA E/S/C and ASIL assessments to confirm all ratings are correct: `show asil ratings`\n"
                    "- If ratings are correct, the HARA is complete.\n"
                ).format(
                    iso_standard=complete_output.get('iso_standard', 'ISO 26262-3:2018'),
                    clause=complete_output.get('clause', '6.4.6'),
                    message=complete_output.get('message', 'No ASIL-rated (A, B, C, D) hazards were found.'),
                    details=complete_output.get('details', {})
                )
            log.warning("Safety goals output is not ready or failed.")
            return content

        # --- Extract data ---
        system_name = complete_output.get('system_name', 'N/A')
        safety_goals = complete_output.get('safety_goals', [])
        stats = complete_output.get('statistics', {})
        timestamp = complete_output.get('timestamp', 'N/A')
        iso_standard = complete_output.get('iso_standard', 'N/A')
        clause = complete_output.get('clause', 'N/A')
        compliance_notes = complete_output.get('compliance_notes', [])
        next_steps = complete_output.get('next_steps', [])
        validation_issues = complete_output.get('validation_issues', [])

        # --- Build Markdown Report ---
        output = f"## 🎯 Safety Goals Derived for {system_name}\n\n"
        output += f"*{iso_standard}, {clause}*\n\n"
        output += f"**Derivation Date:** {timestamp}\n\n"

        # --- Summary Statistics ---
        total_goals = stats.get('total_goals', len(safety_goals))
        asil_dist = stats.get('asil_distribution', {})

        output += "### 📊 Summary Statistics\n\n"
        output += f"- **Total Safety Goals Derived:** {total_goals}\n\n"

        # ASIL Distribution
        output += "**ASIL Distribution:**\n"
        for asil in ['D', 'C', 'B', 'A']:
            count = asil_dist.get(asil, 0)
            if count > 0:
                output += f"- **ASIL {asil}**: {count} goals\n"
        
        if total_goals == 0:
             output += "- No ASIL-rated (A, B, C, D) hazards were found.\n"
        output += "\n"

        # --- Detailed Safety Goals Table ---
        output += "### 📋 Derived Safety Goals\n\n"

        if not safety_goals:
            output += "No safety goals were derived.\n\n"
        else:
            # --- THIS IS THE FIX ---
            output += "| ID   | Safety Goal | ASIL | Safe State | FTTI (ms) |\n"
            output += "|:-----|:------------|:----:|:-----------|:----------|\n"

            for goal in safety_goals:
                sg_id = goal.get('sg_id', 'N/A')
                statement = goal.get('statement', 'N/A')
                asil = goal.get('asil', '?')
                safe_state = goal.get('safe_state', 'TBD')
                ftti = goal.get('ftti_ms', 'TBD')

                output += f"| {sg_id} "
                output += f"| {statement} "
                output += f"| **{asil}** "
                output += f"| {safe_state} "
                output += f"| {ftti} |\n"
            
            output += "\n"

        # # --- Validation Issues ---
        # if validation_issues:
        #     output += "### ⚠️ Validation Issues for Review\n\n"
        #     output += "The following items require manual review and refinement:\n"
        #     for issue in validation_issues:
        #         output += f"- {issue}\n"
        #     output += "\n"

        # --- Compliance Notes ---
        if compliance_notes:
            output += "### ✅ ISO 26262 Compliance\n\n"
            for note in compliance_notes:
                output += f"- {note}\n"
            output += "\n"

        # --- Next Steps ---
        if next_steps:
            output += "### 🚀 Recommended Next Steps\n\n"
            for step in next_steps:
                output += f"- {step}\n"
            output += "\n"

        return output

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