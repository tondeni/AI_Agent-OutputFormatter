# hooks/output_formatter.py - OPTIMIZED FOR FSC DEVELOPER PLUGIN
# Rule-based for ISO 26262 compliance, LLM for insights only

from cat.mad_hatter.decorators import hook
from cat.log import log
from typing import Dict, List


# ============================================================================
# RECOMMENDED: Rule-Based Formatters for All Standard Operations
# ============================================================================

class FSCTableFormatter:
    """
    Deterministic formatting for ISO 26262 compliance.
    Use this for all traceability artifacts.
    """
    
    @staticmethod
    def format_fsr_table(fsrs: List[Dict]) -> str:
        """FSR table - ISO 26262-3:2018 Clause 7.4.2"""
        if not fsrs:
            return ""
        
        table = "\n\n## 📋 Functional Safety Requirements\n\n"
        table += "| FSR-ID | Type | ASIL | Safety Goal | Description | Allocated To |\n"
        table += "|--------|------|------|-------------|-------------|-------------|\n"
        
        for fsr in fsrs:
            desc = fsr.get('description', 'N/A')[:50]
            desc = desc.replace('\n', ' ').replace('|', '\\|')
            
            table += f"| {fsr.get('id', 'N/A')} | "
            table += f"{fsr.get('type', 'N/A')} | "
            table += f"{fsr.get('asil', 'QM')} | "
            table += f"{fsr.get('safety_goal_id', 'N/A')} | "
            table += f"{desc}... | "
            table += f"{fsr.get('allocated_to', 'TBD')} |\n"
        
        # Add compliance footer
        table += "\n*ISO 26262-3:2018, Clause 7.4.2 - Functional Safety Requirements*\n"
        
        return table
    
    @staticmethod
    def format_allocation_matrix(fsrs: List[Dict]) -> str:
        """Allocation Matrix - ISO 26262-3:2018 Clause 7.4.2.8"""
        if not fsrs:
            return ""
        
        table = "\n\n## 🗺️ FSR Allocation Matrix\n\n"
        table += "*Traceability: Safety Goal → FSR → Architectural Element*\n\n"
        table += "| Safety Goal | FSR-ID | Type | ASIL | Component | Component Type | Rationale |\n"
        table += "|-------------|--------|------|------|-----------|----------------|----------|\n"
        
        for fsr in fsrs:
            if fsr.get('allocated_to') == 'TBD':
                continue
            
            rationale = fsr.get('allocation_rationale', 'N/A')[:40]
            rationale = rationale.replace('\n', ' ').replace('|', '\\|')
            
            table += f"| {fsr.get('safety_goal_id', 'N/A')} | "
            table += f"{fsr.get('id', 'N/A')} | "
            table += f"{fsr.get('type', 'N/A')} | "
            table += f"{fsr.get('asil', 'QM')} | "
            table += f"{fsr.get('allocated_to', 'TBD')} | "
            table += f"{fsr.get('allocation_type', 'N/A')} | "
            table += f"{rationale}... |\n"
        
        table += "\n*ISO 26262-3:2018, Clause 7.4.2.8 - Allocation to Architectural Elements*\n"
        
        return table
    
    @staticmethod
    def format_strategy_table(strategies: List[Dict], goals) -> str:
        """Safety Strategies - ISO 26262-3:2018 Clause 7.4.2.3"""
        if not strategies:
            return ""
        
        table = "\n\n## 🎯 Safety Strategies\n\n"
        table += "| SG-ID | Strategy Type | Strategy Description |\n"
        table += "|-------|---------------|----------------------|\n"
        
        for strategy_data in strategies:
            sg_id = strategy_data.get('safety_goal_id', 'N/A')
            strats = strategy_data.get('strategies', {})
            
            # List all 10 required strategies
            for strat_type, strat_desc in strats.items():
                desc = strat_desc[:70] + "..." if len(strat_desc) > 70 else strat_desc
                desc = desc.replace('\n', ' ').replace('|', '\\|')
                
                table += f"| {sg_id} | {strat_type} | {desc} |\n"
        
        table += "\n*ISO 26262-3:2018, Clause 7.4.2.3 - Safety Strategies*\n"
        
        return table
    
    @staticmethod
    def format_mechanism_table(mechanisms: List[Dict], mappings: List[Dict]) -> str:
        """Safety Mechanisms - ISO 26262-4:2018 Clause 6.4.5.4"""
        if not mechanisms:
            return ""
        
        table = "\n\n## 🛡️ Safety Mechanisms\n\n"
        table += "| Mechanism ID | Name | Type | ASIL | Coverage | FSRs Covered |\n"
        table += "|--------------|------|------|------|----------|-------------|\n"
        
        for mech in mechanisms:
            fsrs_covered = ', '.join(mech.get('applicable_fsrs', [])[:3])
            if len(mech.get('applicable_fsrs', [])) > 3:
                fsrs_covered += "..."
            
            table += f"| {mech.get('id', 'N/A')} | "
            table += f"{mech.get('name', 'N/A')} | "
            table += f"{mech.get('mechanism_type', 'N/A')} | "
            table += f"{', '.join(mech.get('asil_suitability', ['QM']))} | "
            table += f"{mech.get('diagnostic_coverage', 'N/A')} | "
            table += f"{fsrs_covered} |\n"
        
        table += "\n*ISO 26262-4:2018, Clause 6.4.5.4 - Safety Mechanisms*\n"
        
        return table


# ============================================================================
# OPTIONAL: LLM for Executive Summaries and Insights Only
# ============================================================================

class LLMInsightGenerator:
    """
    Use LLM ONLY for generating insights, not for formatting tables.
    This keeps compliance data deterministic while adding value.
    """
    
    def __init__(self, llm_function):
        self.llm = llm_function
    
    def generate_executive_summary(self, context: Dict) -> str:
        """
        Generate high-level summary with insights.
        This is safe for LLM because it's interpretive, not prescriptive.
        """
        
        prompt = f"""You are an ISO 26262 expert reviewing FSC development progress.

**Context:**
- System: {context.get('system_name')}
- Safety Goals: {context.get('goal_count', 0)}
- FSRs: {context.get('fsr_count', 0)}
- ASIL D FSRs: {context.get('asil_d_count', 0)}
- Allocated: {context.get('allocated_count', 0)}/{context.get('fsr_count', 0)}

**Task:** Write a brief (3-4 sentences) executive summary highlighting:
1. Current progress status
2. Key risk indicators (ASIL D percentage)
3. One actionable recommendation

Keep it professional and concise. No tables or lists."""
        
        try:
            summary = self.llm(prompt)
            return f"\n\n## 📊 Executive Summary\n\n{summary}\n"
        except Exception as e:
            log.error(f"LLM summary failed: {e}")
            return ""
    
    def analyze_asil_distribution(self, fsrs: List[Dict]) -> str:
        """
        Analyze ASIL distribution and provide insights.
        Again, interpretive rather than prescriptive.
        """
        
        asil_counts = {}
        for fsr in fsrs:
            asil = fsr.get('asil', 'QM')
            asil_counts[asil] = asil_counts.get(asil, 0) + 1
        
        prompt = f"""As an ISO 26262 expert, analyze this ASIL distribution:

ASIL D: {asil_counts.get('D', 0)} FSRs
ASIL C: {asil_counts.get('C', 0)} FSRs
ASIL B: {asil_counts.get('B', 0)} FSRs
ASIL A: {asil_counts.get('A', 0)} FSRs
QM: {asil_counts.get('QM', 0)} FSRs

Write 2-3 sentences about:
1. Risk profile
2. Development implications
3. Verification intensity needed

Be specific and actionable."""
        
        try:
            analysis = self.llm(prompt)
            return f"\n\n### ASIL Analysis\n\n{analysis}\n"
        except Exception as e:
            log.error(f"LLM analysis failed: {e}")
            return ""


# ============================================================================
# MAIN FORMATTER - Hybrid with Rule Priority
# ============================================================================

class FSCHybridFormatter:
    """
    Hybrid formatter optimized for FSC development:
    - Rules: All tables, matrices, traceability (70% of operations)
    - LLM: Summaries, insights, analysis only (30% of operations)
    """
    
    def __init__(self, llm_function):
        self.table_formatter = FSCTableFormatter()
        self.insight_generator = LLMInsightGenerator(llm_function)
    
    def format(self, content: str, cat) -> str:
        """
        Main formatting logic for FSC plugin.
        """
        
        formatted = content
        last_op = cat.working_memory.get('last_operation')
        
        # RULE-BASED: Standard FSC operations (deterministic)
        if last_op == 'fsr_derivation':
            fsrs = cat.working_memory.get('fsc_functional_requirements', [])
            formatted += self.table_formatter.format_fsr_table(fsrs)
            # Add statistics (rule-based)
            formatted += self._add_fsr_statistics(fsrs)
        
        elif last_op == 'fsr_allocation':
            fsrs = cat.working_memory.get('fsc_functional_requirements', [])
            formatted += self.table_formatter.format_allocation_matrix(fsrs)
            # Add allocation stats (rule-based)
            formatted += self._add_allocation_statistics(fsrs)
        
        elif last_op == 'strategy_development':
            # ADDED: Format strategies as table when they are developed
            strategies = cat.working_memory.get('fsc_safety_strategies', [])
            goals = cat.working_memory.get('fsc_safety_goals', [])
            formatted += self.table_formatter.format_strategy_table(strategies, goals)
            
            # # Add strategy statistics
            # formatted += self._add_strategy_statistics(strategies, goals)
        
        elif last_op == 'fsr_mechanisms':
            mechanisms = cat.working_memory.get('fsc_safety_mechanisms', [])
            mappings = cat.working_memory.get('fsc_mechanism_mappings', [])
            formatted += self.table_formatter.format_mechanism_table(mechanisms, mappings)
        
        # LLM-BASED: Executive summaries and insights only
        elif last_op == 'fsc_verification':
            # Verification report stays rule-based for compliance
            # But add LLM insight at the end
            context = self._gather_context(cat)
            formatted += self.insight_generator.generate_executive_summary(context)
        
        return formatted
    
    def _add_fsr_statistics(self, fsrs: List[Dict]) -> str:
        """Rule-based statistics (deterministic)."""
        
        stats = "\n\n### Statistics\n\n"
        
        # Count by ASIL
        asil_counts = {}
        for fsr in fsrs:
            asil = fsr.get('asil', 'QM')
            asil_counts[asil] = asil_counts.get(asil, 0) + 1
        
        stats += f"- **Total FSRs:** {len(fsrs)}\n"
        for asil in ['D', 'C', 'B', 'A', 'QM']:
            if asil in asil_counts:
                percentage = (asil_counts[asil] / len(fsrs) * 100) if fsrs else 0
                stats += f"- **ASIL {asil}:** {asil_counts[asil]} ({percentage:.1f}%)\n"
        
        return stats
    
    def _add_allocation_statistics(self, fsrs: List[Dict]) -> str:
        """Rule-based allocation stats."""
        
        allocated = [f for f in fsrs if f.get('allocated_to') != 'TBD']
        percentage = (len(allocated) / len(fsrs) * 100) if fsrs else 0
        
        stats = "\n\n### Allocation Status\n\n"
        stats += f"- **Allocated:** {len(allocated)}/{len(fsrs)} ({percentage:.1f}%)\n"
        stats += f"- **Pending:** {len(fsrs) - len(allocated)}\n"
        
        if percentage < 100:
            stats += f"\n⚠️  **Action Required:** Complete allocation for {len(fsrs) - len(allocated)} FSRs\n"
        else:
            stats += f"\n✅ **All FSRs allocated** - Ready for mechanism identification\n"
        
        return stats
    
    def _gather_context(self, cat) -> Dict:
        """Gather context for LLM insights."""
        
        fsrs = cat.working_memory.get('fsc_functional_requirements', [])
        goals = cat.working_memory.get('fsc_safety_goals', [])
        
        return {
            'system_name': cat.working_memory.get('system_name', 'Unknown'),
            'goal_count': len(goals),
            'fsr_count': len(fsrs),
            'asil_d_count': len([f for f in fsrs if f.get('asil') == 'D']),
            'allocated_count': len([f for f in fsrs if f.get('allocated_to') != 'TBD']),
        }


# ============================================================================
# ENRICHMENT FUNCTIONS
# ============================================================================

def add_export_offers(message: Dict, cat) -> Dict:
    """Add export options - rule-based."""
    
    last_op = cat.working_memory.get('last_operation')
    
    offers = {
        'fsr_derivation': "\n\n📊 **Export:** `export FSRs to excel`",
        'fsr_allocation': "\n\n📊 **Export:** `export allocation matrix to excel`",
        'strategy_development': "\n\n📊 **Export:** `export strategies to excel`",
        'fsr_mechanisms': "\n\n📊 **Export:** `export mechanisms to excel`",
        'validation_criteria_specification': "\n\n📊 **Export:** `export validation criteria to excel`",
        'fsc_verification': "\n\n📄 **Generate:** `generate FSC document` (Word)",
    }
    
    if last_op in offers:
        message['content'] += offers[last_op]
    
    return message


def add_next_steps(message: Dict, cat) -> Dict:
    """Add workflow guidance - rule-based."""
    
    stage = cat.working_memory.get('fsc_stage')
    content = message.get('content', '')
    
    # Don't add if already present
    if 'next step' in content.lower() or '➡️' in content:
        return message
    
    next_steps = {
        'hara_loaded': "\n\n➡️ **Next:** `develop safety strategies for all goals`",
        'strategies_developed': "\n\n➡️ **Next:** `derive FSRs for all goals`",
        'fsrs_derived': "\n\n➡️ **Next:** `allocate FSRs`",
        'fsrs_allocated': "\n\n➡️ **Next:** `identify safety mechanisms`",
        'mechanisms_identified': "\n\n➡️ **Next:** `specify validation criteria`",
        'validation_criteria_specified': "\n\n➡️ **Next:** `verify FSC`",
    }
    
    if stage in next_steps:
        message['content'] += next_steps[stage]
    
    return message


# ============================================================================
# MAIN HOOK
# ============================================================================

@hook(priority=5)
def before_cat_sends_message(message, cat):
    """
    Main formatting hook for FSC Developer plugin.
    
    Strategy:
    - Rule-based for all ISO 26262 artifacts (compliance)
    - LLM for insights only (added value)
    - No LLM for tables/matrices (determinism)
    """
    
    content = message.get("content", "")
    
    # Skip if too short
    if not content or len(content.strip()) < 20:
        return message
    
    try:
        # Apply formatting
        formatter = FSCHybridFormatter(cat.llm)
        message['content'] = formatter.format(content, cat)
        
        # Add enrichments (always rule-based)
        message = add_export_offers(message, cat)
        message = add_next_steps(message, cat)
        
        log.info("✅ FSC formatting complete")
        
    except Exception as e:
        log.error(f"❌ Formatting error: {e}")
        import traceback
        log.error(traceback.format_exc())
    
    return message