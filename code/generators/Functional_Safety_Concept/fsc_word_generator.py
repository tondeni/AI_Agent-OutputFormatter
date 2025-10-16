# ==============================================================================
# AI_Agent-OutputFormatter/code/generators/fsc_word_generator.py
# Self-contained FSC Word document generator with completeness warnings
# ==============================================================================

import os
import json
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn


class FSCWordGenerator:
    """
    Self-contained generator for FSC Word documents.
    
    Handles:
    - Data validation
    - Completeness checking
    - Statistics calculation
    - Document structure
    - Word file creation
    """
    
    def __init__(self, plugin_folder: Optional[str] = None):
        """Initialize generator with optional plugin folder for templates"""
        self.plugin_folder = plugin_folder
        self.fsc_structure = self._load_fsc_structure() if plugin_folder else None
        self.completeness_warnings = []
    
    def _load_fsc_structure(self) -> Optional[Dict]:
        """Load FSC structure template from JSON"""
        if not self.plugin_folder:
            return None
        
        template_path = os.path.join(self.plugin_folder, "templates", "fsc_structure.json")
        
        if not os.path.exists(template_path):
            return None
        
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return None
    
    def check_completeness(
        self,
        goals_data: List[Dict],
        fsrs_data: List[Dict],
        strategies_data: Optional[Dict],
        fsc_data: Optional[Dict]
    ) -> List[str]:
        """
        Check completeness of FSC data and return warnings.
        
        Returns:
            List of warning messages for incomplete sections
        """
        warnings = []
        
        # Check safety strategies
        if not strategies_data or len(strategies_data) == 0:
            warnings.append(
                "⚠️ No safety strategies defined - Section 2.2 will be incomplete. "
                "Run: 'develop safety strategies for all goals'"
            )
        elif len(strategies_data) < len(goals_data):
            warnings.append(
                f"⚠️ Only {len(strategies_data)}/{len(goals_data)} safety goals have strategies. "
                "Some goals in Section 2.2 will be incomplete."
            )
        
        # Check allocation
        if not fsc_data:
            fsc_data = {}
        
        allocated_count = sum(1 for fsr in fsrs_data if fsr.get('allocated_to'))
        if allocated_count == 0:
            warnings.append(
                "⚠️ No FSRs allocated - Section 4.1 (Allocation Matrix) will be incomplete. "
                "Run: 'allocate all FSRs'"
            )
        elif allocated_count < len(fsrs_data):
            alloc_pct = allocated_count / len(fsrs_data) * 100
            warnings.append(
                f"⚠️ Only {allocated_count}/{len(fsrs_data)} FSRs allocated ({alloc_pct:.0f}%). "
                "Section 4.1 (Allocation Matrix) is incomplete."
            )
        
        # Check validation criteria
        validation_data = fsc_data.get('validation', [])
        if not validation_data or len(validation_data) == 0:
            warnings.append(
                "⚠️ No validation criteria defined - Section 7.2 will be incomplete. "
                "Run: 'specify validation criteria'"
            )
        
        # Check safety mechanisms
        mechanisms_data = fsc_data.get('mechanisms', [])
        if not mechanisms_data or len(mechanisms_data) == 0:
            warnings.append(
                "⚠️ No safety mechanisms defined - Section 5 will contain placeholder text only. "
                "Safety mechanisms should be specified."
            )
        
        # Check ASIL decomposition
        decomposition_data = fsc_data.get('decomposition', [])
        has_high_asil = any(fsr.get('asil', '') in ['C', 'D'] for fsr in fsrs_data)
        if has_high_asil and (not decomposition_data or len(decomposition_data) == 0):
            warnings.append(
                "ℹ️ No ASIL decomposition applied. Consider if decomposition could simplify "
                "development for ASIL C/D requirements."
            )
        
        # Check FSR coverage
        goals_with_fsrs = set()
        for fsr in fsrs_data:
            parent_sg = fsr.get('parent_safety_goal', fsr.get('parent_sg', []))
            if isinstance(parent_sg, list):
                goals_with_fsrs.update(parent_sg)
            elif parent_sg:
                goals_with_fsrs.add(parent_sg)
        
        if len(goals_with_fsrs) < len(goals_data):
            missing_count = len(goals_data) - len(goals_with_fsrs)
            warnings.append(
                f"⚠️ {missing_count} safety goal(s) have no FSRs. "
                "All goals should have at least one FSR."
            )
        
        return warnings
    
    def validate_data(
        self,
        goals_data: List[Dict],
        fsrs_data: List[Dict]
    ) -> Tuple[bool, List[str], List[str]]:
        """
        Validate FSC data for Word document generation.
        
        Returns:
            Tuple of (is_valid, warnings, errors)
        """
        warnings = []
        errors = []
        
        # Check required data
        if not goals_data:
            errors.append("No safety goals available")
        
        if not fsrs_data:
            errors.append("No FSRs available")
        
        if errors:
            return False, warnings, errors
        
        # Check coverage (warning, not error)
        goals_with_fsrs = set()
        for fsr in fsrs_data:
            parent_sg = fsr.get('parent_safety_goal', fsr.get('parent_sg', []))
            if isinstance(parent_sg, list):
                goals_with_fsrs.update(parent_sg)
            elif parent_sg:
                goals_with_fsrs.add(parent_sg)
        
        coverage = len(goals_with_fsrs) / len(goals_data) * 100 if goals_data else 0
        
        if coverage < 100:
            warnings.append(
                f"Only {len(goals_with_fsrs)}/{len(goals_data)} safety goals have FSRs ({coverage:.0f}%)"
            )
        
        # Check allocation (warning, not error)
        allocated = sum(1 for fsr in fsrs_data if fsr.get('allocated_to'))
        if allocated < len(fsrs_data):
            alloc_pct = allocated / len(fsrs_data) * 100
            if alloc_pct < 50:
                warnings.append(f"Only {allocated}/{len(fsrs_data)} FSRs allocated ({alloc_pct:.0f}%)")
        
        # Check FSR types (warning, not error)
        has_detection = any(fsr.get('type', '').lower() == 'detection' for fsr in fsrs_data)
        has_reaction = any(fsr.get('type', '').lower() == 'reaction' for fsr in fsrs_data)
        
        if not has_detection:
            warnings.append("No detection requirements defined")
        
        if not has_reaction:
            warnings.append("No reaction requirements defined")
        
        return True, warnings, errors
    
    def calculate_statistics(
        self,
        goals_data: List[Dict],
        fsrs_data: List[Dict],
        strategies_data: Dict
    ) -> Dict:
        """Calculate document statistics"""
        
        # FSRs by type
        fsr_by_type = {
            'detection': 0,
            'reaction': 0,
            'indication': 0,
            'other': 0
        }
        
        for fsr in fsrs_data:
            fsr_type = fsr.get('type', 'other').lower()
            if fsr_type in fsr_by_type:
                fsr_by_type[fsr_type] += 1
            else:
                fsr_by_type['other'] += 1
        
        # ASIL distribution
        asil_dist = {}
        for fsr in fsrs_data:
            asil = fsr.get('asil', 'QM')
            asil_dist[asil] = asil_dist.get(asil, 0) + 1
        
        # Allocation
        allocated = sum(1 for fsr in fsrs_data if fsr.get('allocated_to'))
        
        # Coverage
        goals_with_fsrs = set()
        for fsr in fsrs_data:
            parent_sg = fsr.get('parent_safety_goal', fsr.get('parent_sg', []))
            if isinstance(parent_sg, list):
                goals_with_fsrs.update(parent_sg)
            elif parent_sg:
                goals_with_fsrs.add(parent_sg)
        
        # Page estimate
        pages_estimate = len(goals_data) * 3 + len(fsrs_data) // 2 + 15
        
        return {
            'total_goals': len(goals_data),
            'total_fsrs': len(fsrs_data),
            'fsr_by_type': fsr_by_type,
            'asil_distribution': asil_dist,
            'allocated_fsrs': allocated,
            'allocation_pct': round(allocated / len(fsrs_data) * 100, 1) if fsrs_data else 0,
            'goals_with_fsrs': len(goals_with_fsrs),
            'coverage_pct': round(len(goals_with_fsrs) / len(goals_data) * 100, 1) if goals_data else 0,
            'total_strategies': len(strategies_data),
            'estimated_pages': pages_estimate
        }
    
    def generate(
        self,
        system_name: str,
        goals_data: List[Dict],
        fsrs_data: List[Dict],
        strategies_data: Optional[Dict] = None,
        fsc_data: Optional[Dict] = None
    ) -> Document:
        """
        Generate complete FSC Word document.
        
        Args:
            system_name: Name of the system/item
            goals_data: Safety goals from HARA
            fsrs_data: Functional Safety Requirements
            strategies_data: Safety strategies (optional)
            fsc_data: Additional FSC data (optional)
            
        Returns:
            Document object ready to save
        """
        
        if strategies_data is None:
            strategies_data = {}
        
        if fsc_data is None:
            fsc_data = {}
        
        # Check completeness and store warnings
        self.completeness_warnings = self.check_completeness(
            goals_data, fsrs_data, strategies_data, fsc_data
        )
        
        doc = Document()
        
        # Add styles
        self._add_styles(doc)
        
        # Title Page
        self._add_title_page(doc, system_name)
        
        # Add completeness notice if there are warnings
        if self.completeness_warnings:
            self._add_completeness_notice(doc)
        
        doc.add_page_break()
        
        # 1. Introduction
        self._add_introduction(doc, system_name)
        
        doc.add_page_break()
        
        # 2. Safety Goals Overview
        self._add_safety_goals(doc, goals_data, strategies_data)
        
        doc.add_page_break()
        
        # 3. Functional Safety Requirements
        self._add_fsrs(doc, fsrs_data)
        
        doc.add_page_break()
        
        # 4. FSR Allocation
        self._add_allocation(doc, fsrs_data, fsc_data.get('allocation', {}))
        
        doc.add_page_break()
        
        # 5. Functional Safety Mechanisms
        self._add_safety_mechanisms(doc, fsc_data.get('mechanisms', []))
        
        doc.add_page_break()
        
        # 6. ASIL Decomposition
        self._add_asil_decomposition(doc, fsc_data.get('decomposition', []))
        
        doc.add_page_break()
        
        # 7. Verification and Validation
        self._add_verification(doc, fsc_data.get('validation', []))
        
        doc.add_page_break()
        
        # 8. Traceability
        self._add_traceability(doc, goals_data, fsrs_data)
        
        doc.add_page_break()
        
        # 9. Approvals
        self._add_approvals(doc)
        
        return doc
    
    def _add_styles(self, doc):
        """Add custom styles"""
        try:
            # Title
            title = doc.styles.add_style('FSC_Title', 1)
            title.font.name = 'Calibri'
            title.font.size = Pt(26)
            title.font.bold = True
            title.font.color.rgb = RGBColor(31, 78, 121)
            title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Heading 1
            h1 = doc.styles.add_style('FSC_H1', 1)
            h1.font.name = 'Calibri'
            h1.font.size = Pt(16)
            h1.font.bold = True
            h1.font.color.rgb = RGBColor(31, 78, 121)
            h1.paragraph_format.space_before = Pt(18)
            h1.paragraph_format.space_after = Pt(6)
            
            # Heading 2
            h2 = doc.styles.add_style('FSC_H2', 1)
            h2.font.name = 'Calibri'
            h2.font.size = Pt(14)
            h2.font.bold = True
            h2.font.color.rgb = RGBColor(68, 114, 196)
            h2.paragraph_format.space_before = Pt(12)
            h2.paragraph_format.space_after = Pt(6)
            
            # Body
            body = doc.styles.add_style('FSC_Body', 1)
            body.font.name = 'Calibri'
            body.font.size = Pt(11)
            body.paragraph_format.line_spacing = 1.15
            body.paragraph_format.space_after = Pt(6)
            
            # ISO Reference
            iso = doc.styles.add_style('FSC_ISO', 1)
            iso.font.name = 'Calibri'
            iso.font.size = Pt(9)
            iso.font.italic = True
            iso.font.color.rgb = RGBColor(89, 89, 89)
            
            # Warning style
            warning = doc.styles.add_style('FSC_Warning', 1)
            warning.font.name = 'Calibri'
            warning.font.size = Pt(10)
            warning.font.italic = True
            warning.font.color.rgb = RGBColor(255, 140, 0)  # Orange
            warning.paragraph_format.space_after = Pt(6)
            
            # Incomplete marker style
            incomplete = doc.styles.add_style('FSC_Incomplete', 1)
            incomplete.font.name = 'Calibri'
            incomplete.font.size = Pt(10)
            incomplete.font.italic = True
            incomplete.font.color.rgb = RGBColor(255, 0, 0)  # Red
            incomplete.paragraph_format.space_after = Pt(3)
        except:
            pass  # Styles may exist
    
    def _add_completeness_notice(self, doc):
        """Add completeness notice to title page"""
        doc.add_paragraph()
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("⚠️ DOCUMENT COMPLETENESS NOTICE ⚠️")
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 140, 0)
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.style = 'FSC_Warning'
        p.add_run(
            "This document contains incomplete sections. "
            "Please review the warnings below and complete the missing workflow steps."
        )
        
        doc.add_paragraph()
        
        # List warnings
        for warning in self.completeness_warnings:
            p = doc.add_paragraph(warning, style='List Bullet')
            p.style = 'FSC_Warning'
    
    def _add_incomplete_marker(self, doc, section_name: str, missing_step: str):
        """Add incomplete section marker"""
        p = doc.add_paragraph()
        p.style = 'FSC_Incomplete'
        run = p.add_run(
            f"⚠️ INCOMPLETE: {section_name} - {missing_step}"
        )
        run.font.bold = True
    
    def _add_title_page(self, doc, system_name):
        """Add title page"""
        p = doc.add_paragraph()
        p.style = 'FSC_Title'
        p.add_run("Functional Safety Concept\n").bold = True
        p.add_run(f"\n{system_name}\n\n").font.size = Pt(20)
        p.add_run("\nISO 26262-3:2018, Clause 7\n").font.size = Pt(14)
        
        info = doc.add_paragraph()
        info.alignment = WD_ALIGN_PARAGRAPH.CENTER
        info.add_run(f"Version: 1.0\n").font.size = Pt(11)
        info.add_run(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n").font.size = Pt(11)
    
    def _add_introduction(self, doc, system_name):
        """Add introduction section"""
        doc.add_heading('1. Introduction', level=1).style = 'FSC_H1'
        
        # 1.1 Purpose
        doc.add_heading('1.1 Purpose and Scope', level=2).style = 'FSC_H2'
        self._add_iso_ref(doc, "ISO 26262-3:2018, Clause 7.5")
        
        p = doc.add_paragraph()
        p.style = 'FSC_Body'
        p.add_run(
            f"This Functional Safety Concept defines the functional safety requirements "
            f"derived from safety goals, their allocation to the {system_name} architecture, "
            f"and functional-level safety mechanisms per ISO 26262-3:2018, Clause 7."
        )
        
        # 1.2 References
        doc.add_heading('1.2 Referenced Documents', level=2).style = 'FSC_H2'
        self._add_iso_ref(doc, "ISO 26262-8:2018, Clause 6.4.2")
        
        refs = [
            f"Item Definition for {system_name}",
            f"HARA for {system_name}",
            "ISO 26262-3:2018",
            "ISO 26262-9:2018"
        ]
        for ref in refs:
            doc.add_paragraph(ref, style='List Bullet')
        
        # 1.3 Terms
        doc.add_heading('1.3 Terms and Abbreviations', level=2).style = 'FSC_H2'
        
        table = self._add_table(doc, 6, 2)
        table.rows[0].cells[0].text = "Term"
        table.rows[0].cells[1].text = "Definition"
        
        terms = [
            ("FSR", "Functional Safety Requirement"),
            ("ASIL", "Automotive Safety Integrity Level"),
            ("FTTI", "Fault Tolerant Time Interval"),
            ("Safe State", "Operating mode without unreasonable risk"),
            ("FHTI", "Fault Handling Time Interval")
        ]
        
        for i, (term, defn) in enumerate(terms, 1):
            table.rows[i].cells[0].text = term
            table.rows[i].cells[1].text = defn
    
    def _add_safety_goals(self, doc, goals_data, strategies_data):
        """Add safety goals section"""
        doc.add_heading('2. Safety Goals Overview', level=1).style = 'FSC_H1'
        
        doc.add_heading('2.1 Safety Goals Summary', level=2).style = 'FSC_H2'
        self._add_iso_ref(doc, "ISO 26262-3:2018, Clause 7.4.1")
        
        if goals_data:
            table = self._add_table(doc, len(goals_data) + 1, 5)
            
            # Headers
            headers = ["SG-ID", "Safety Goal", "ASIL", "Safe State", "FTTI"]
            for i, h in enumerate(headers):
                table.rows[0].cells[i].text = h
            
            # Data
            for i, goal in enumerate(goals_data, 1):
                table.rows[i].cells[0].text = goal.get('sg_id', goal.get('id', f'SG-{i:03d}'))
                table.rows[i].cells[1].text = goal.get('statement', goal.get('goal', ''))
                table.rows[i].cells[2].text = goal.get('asil', 'QM')
                table.rows[i].cells[3].text = goal.get('safe_state', 'TBD')
                table.rows[i].cells[4].text = str(goal.get('ftti_ms', goal.get('ftti', 'TBD')))
        
        doc.add_heading('2.2 Safety Goal Rationale', level=2).style = 'FSC_H2'
        
        if not strategies_data or len(strategies_data) == 0:
            self._add_incomplete_marker(
                doc,
                "Safety Goal Rationale",
                "Run 'develop safety strategies for all goals'"
            )
            p = doc.add_paragraph()
            p.style = 'FSC_Body'
            p.add_run(
                "Safety strategies describe the high-level approach for achieving each safety goal. "
                "This section will be completed when safety strategies are developed."
            )
        else:
            # Add strategies
            strategies_added = 0
            for goal in goals_data:
                sg_id = goal.get('sg_id', goal.get('id', ''))
                
                if sg_id in strategies_data:
                    strategy = strategies_data[sg_id]
                    if isinstance(strategy, dict):
                        narrative = strategy.get('narrative', '')
                        if narrative:
                            p = doc.add_paragraph()
                            p.style = 'FSC_Body'
                            run = p.add_run(f"{sg_id}: ")
                            run.bold = True
                            p.add_run(narrative)
                            strategies_added += 1
                else:
                    # Mark missing strategy
                    p = doc.add_paragraph()
                    p.style = 'FSC_Warning'
                    p.add_run(f"⚠️ {sg_id}: Strategy not defined")
            
            if strategies_added < len(goals_data):
                p = doc.add_paragraph()
                p.style = 'FSC_Warning'
                p.add_run(
                    f"\n⚠️ Note: Only {strategies_added}/{len(goals_data)} safety goals have strategies defined."
                )
    
    def _add_fsrs(self, doc, fsrs_data):
        """Add FSRs section"""
        doc.add_heading('3. Functional Safety Requirements', level=1).style = 'FSC_H1'
        
        doc.add_heading('3.1 FSR Overview', level=2).style = 'FSC_H2'
        self._add_iso_ref(doc, "ISO 26262-3:2018, Clause 7.4.2")
        
        p = doc.add_paragraph()
        p.style = 'FSC_Body'
        p.add_run("FSRs specify WHAT functional behavior is required to prevent safety goal violations.")
        
        # Group by type
        fsr_by_type = {'detection': [], 'reaction': [], 'indication': [], 'other': []}
        
        for fsr in fsrs_data:
            fsr_type = fsr.get('type', 'other').lower()
            if fsr_type in fsr_by_type:
                fsr_by_type[fsr_type].append(fsr)
            else:
                fsr_by_type['other'].append(fsr)
        
        # Detection
        doc.add_heading('3.2 Detection Requirements', level=2).style = 'FSC_H2'
        self._add_iso_ref(doc, "ISO 26262-3:2018, Clause 7.4.2.3.b")
        
        if not fsr_by_type['detection']:
            self._add_incomplete_marker(
                doc,
                "Detection Requirements",
                "No detection FSRs derived"
            )
        else:
            self._add_fsr_table(doc, fsr_by_type['detection'])
        
        # Reaction
        doc.add_heading('3.3 Reaction Requirements', level=2).style = 'FSC_H2'
        self._add_iso_ref(doc, "ISO 26262-3:2018, Clause 7.4.2.3.c")
        
        if not fsr_by_type['reaction']:
            self._add_incomplete_marker(
                doc,
                "Reaction Requirements",
                "No reaction FSRs derived"
            )
        else:
            self._add_fsr_table(doc, fsr_by_type['reaction'], include_fhti=True)
        
        # Indication
        doc.add_heading('3.4 Indication Requirements', level=2).style = 'FSC_H2'
        self._add_iso_ref(doc, "ISO 26262-3:2018, Clause 7.4.2.3.f,g")
        
        if not fsr_by_type['indication']:
            self._add_incomplete_marker(
                doc,
                "Indication Requirements",
                "No indication FSRs derived"
            )
        else:
            self._add_fsr_table(doc, fsr_by_type['indication'])
    
    def _add_fsr_table(self, doc, fsrs, include_fhti=False):
        """Add FSR table"""
        if not fsrs:
            return
        
        cols = 5 if include_fhti else 4
        table = self._add_table(doc, len(fsrs) + 1, cols)
        
        # Headers
        table.rows[0].cells[0].text = "FSR-ID"
        table.rows[0].cells[1].text = "Requirement"
        table.rows[0].cells[2].text = "ASIL"
        table.rows[0].cells[3].text = "Parent SG"
        if include_fhti:
            table.rows[0].cells[4].text = "FHTI (ms)"
        
        # Data
        for i, fsr in enumerate(fsrs, 1):
            table.rows[i].cells[0].text = fsr.get('id', fsr.get('fsr_id', ''))
            table.rows[i].cells[1].text = fsr.get('description', fsr.get('requirement', ''))[:100]
            table.rows[i].cells[2].text = fsr.get('asil', 'QM')
            
            parent_sg = fsr.get('parent_safety_goal', fsr.get('parent_sg', []))
            if isinstance(parent_sg, list):
                table.rows[i].cells[3].text = ', '.join(parent_sg)
            else:
                table.rows[i].cells[3].text = str(parent_sg)
            
            if include_fhti:
                table.rows[i].cells[4].text = str(fsr.get('fhti_ms', fsr.get('ftti_ms', 'TBD')))
    
    def _add_allocation(self, doc, fsrs_data, allocation_matrix):
        """Add allocation section"""
        doc.add_heading('4. FSR Allocation', level=1).style = 'FSC_H1'
        
        doc.add_heading('4.1 Allocation Matrix', level=2).style = 'FSC_H2'
        self._add_iso_ref(doc, "ISO 26262-3:2018, Clause 7.4.2.8")
        
        # Check allocation status
        allocated_count = sum(1 for fsr in fsrs_data if fsr.get('allocated_to'))
        
        if allocated_count == 0:
            self._add_incomplete_marker(
                doc,
                "Allocation Matrix",
                "Run 'allocate all FSRs'"
            )
            p = doc.add_paragraph()
            p.style = 'FSC_Body'
            p.add_run(
                "FSRs must be allocated to architectural elements (hardware, software, external systems). "
                "This section will be completed when FSRs are allocated."
            )
        else:
            if allocated_count < len(fsrs_data):
                p = doc.add_paragraph()
                p.style = 'FSC_Warning'
                alloc_pct = allocated_count / len(fsrs_data) * 100
                p.add_run(
                    f"⚠️ Partial allocation: {allocated_count}/{len(fsrs_data)} FSRs allocated ({alloc_pct:.0f}%)"
                )
            
            # Create allocation table
            table = self._add_table(doc, min(len(fsrs_data) + 1, 21), 5)
            
            table.rows[0].cells[0].text = "FSR-ID"
            table.rows[0].cells[1].text = "Requirement"
            table.rows[0].cells[2].text = "ASIL"
            table.rows[0].cells[3].text = "Allocated To"
            table.rows[0].cells[4].text = "Interface"
            
            for i, fsr in enumerate(fsrs_data[:20], 1):
                table.rows[i].cells[0].text = fsr.get('id', fsr.get('fsr_id', ''))
                table.rows[i].cells[1].text = fsr.get('description', fsr.get('requirement', ''))[:50] + "..."
                table.rows[i].cells[2].text = fsr.get('asil', 'QM')
                
                allocated_to = fsr.get('allocated_to', 'TBD')
                cell = table.rows[i].cells[3]
                cell.text = allocated_to
                
                # Highlight unallocated
                if allocated_to == 'TBD' or not allocated_to:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(255, 0, 0)
                
                table.rows[i].cells[4].text = fsr.get('interface_type', 'Internal')
    
    def _add_safety_mechanisms(self, doc, mechanisms_data):
        """Add safety mechanisms section"""
        doc.add_heading('5. Functional Safety Mechanisms', level=1).style = 'FSC_H1'
        self._add_iso_ref(doc, "ISO 26262-3:2018, Clause 7.4.2.3")
        
        if not mechanisms_data or len(mechanisms_data) == 0:
            self._add_incomplete_marker(
                doc,
                "Safety Mechanisms",
                "Safety mechanisms should be defined"
            )
        
        sections = [
            ('5.1 Detection Strategies', 'Functional strategies for fault detection'),
            ('5.2 Reaction Strategies', 'Safe state transition strategies'),
            ('5.3 Fault Tolerance Strategies', 'Redundancy and diversity at functional level'),
            ('5.4 Warning Strategies', 'Driver warning and degradation approaches')
        ]
        
        for title, desc in sections:
            doc.add_heading(title, level=2).style = 'FSC_H2'
            p = doc.add_paragraph()
            p.style = 'FSC_Body'
            p.add_run(desc)
            
            if mechanisms_data:
                # Add actual mechanisms if available
                section_key = title.split()[0].replace('.', '_').lower()
                section_mechanisms = [m for m in mechanisms_data if m.get('category', '').lower() == section_key]
                
                if section_mechanisms:
                    for mech in section_mechanisms:
                        doc.add_paragraph(
                            f"• {mech.get('name', 'Unnamed')}: {mech.get('description', '')}",
                            style='List Bullet'
                        )
    
    def _add_asil_decomposition(self, doc, decomposition_data):
        """Add ASIL decomposition section"""
        doc.add_heading('6. ASIL Decomposition (if applicable)', level=1).style = 'FSC_H1'
        self._add_iso_ref(doc, "ISO 26262-9:2018, Clause 5")
        
        if not decomposition_data or len(decomposition_data) == 0:
            p = doc.add_paragraph()
            p.style = 'FSC_Body'
            p.add_run(
                "No ASIL decomposition has been applied in this FSC. "
                "If ASIL decomposition is applied in future, document: "
                "original requirement, decomposed requirements, independence justification."
            )
        else:
            for decomp in decomposition_data:
                p = doc.add_paragraph()
                p.style = 'FSC_Body'
                p.add_run(f"Original: {decomp.get('original_id', '')} (ASIL {decomp.get('original_asil', '')})\n").bold = True
                p.add_run(f"Decomposed to: {decomp.get('decomposed_requirements', '')}\n")
                p.add_run(f"Independence: {decomp.get('independence_rationale', '')}")
    
    def _add_verification(self, doc, validation_data):
        """Add verification section"""
        doc.add_heading('7. Verification and Validation', level=1).style = 'FSC_H1'
        
        doc.add_heading('7.1 FSC Verification', level=2).style = 'FSC_H2'
        self._add_iso_ref(doc, "ISO 26262-3:2018, Clause 7.4.4")
        
        p = doc.add_paragraph()
        p.style = 'FSC_Body'
        p.add_run("FSC verified through reviews, walkthroughs, and traceability analysis.")
        
        doc.add_heading('7.2 Safety Validation Criteria', level=2).style = 'FSC_H2'
        self._add_iso_ref(doc, "ISO 26262-3:2018, Clause 7.4.3")
        
        if not validation_data or len(validation_data) == 0:
            self._add_incomplete_marker(
                doc,
                "Validation Criteria",
                "Run 'specify validation criteria'"
            )
            p = doc.add_paragraph()
            p.style = 'FSC_Body'
            p.add_run(
                "Safety validation criteria specify acceptance criteria for safety validation "
                "at vehicle level. This section will be completed when validation criteria are specified."
            )
        else:
            p = doc.add_paragraph()
            p.style = 'FSC_Body'
            p.add_run(
                f"A total of {len(validation_data)} validation criteria have been defined "
                "for safety validation at the vehicle level."
            )
            
            # Add summary table if validation data available
            if len(validation_data) > 0:
                table = self._add_table(doc, min(len(validation_data) + 1, 11), 3)
                table.rows[0].cells[0].text = "Criterion ID"
                table.rows[0].cells[1].text = "Description"
                table.rows[0].cells[2].text = "Related FSR"
                
                for i, criterion in enumerate(validation_data[:10], 1):
                    table.rows[i].cells[0].text = criterion.get('id', '')
                    table.rows[i].cells[1].text = criterion.get('description', '')[:80]
                    table.rows[i].cells[2].text = criterion.get('related_fsr', '')
    
    def _add_traceability(self, doc, goals_data, fsrs_data):
        """Add traceability section"""
        doc.add_heading('8. Traceability', level=1).style = 'FSC_H1'
        self._add_iso_ref(doc, "ISO 26262-8:2018, Clause 6.4.3")
        
        table = self._add_table(doc, min(len(fsrs_data) + 1, 21), 4)
        
        table.rows[0].cells[0].text = "Safety Goal"
        table.rows[0].cells[1].text = "FSR-ID"
        table.rows[0].cells[2].text = "Allocated To"
        table.rows[0].cells[3].text = "Mechanism"
        
        for i, fsr in enumerate(fsrs_data[:20], 1):
            parent_sg = fsr.get('parent_safety_goal', fsr.get('parent_sg', ''))
            if isinstance(parent_sg, list):
                parent_sg = ', '.join(parent_sg)
            
            table.rows[i].cells[0].text = str(parent_sg)
            table.rows[i].cells[1].text = fsr.get('id', fsr.get('fsr_id', ''))
            table.rows[i].cells[2].text = fsr.get('allocated_to', 'TBD')
            table.rows[i].cells[3].text = fsr.get('safety_mechanism', 'TBD')
    
    def _add_approvals(self, doc):
        """Add approvals section"""
        doc.add_heading('9. Approvals', level=1).style = 'FSC_H1'
        
        p = doc.add_paragraph()
        p.style = 'FSC_Body'
        p.add_run("This FSC must be approved before proceeding to Technical Safety Concept.")
        
        table = self._add_table(doc, 4, 4, heading_row=False)
        
        table.rows[0].cells[0].text = "Role"
        table.rows[0].cells[1].text = "Name"
        table.rows[0].cells[2].text = "Signature"
        table.rows[0].cells[3].text = "Date"
        
        roles = ["Safety Manager", "System Engineer", "Project Manager"]
        for i, role in enumerate(roles, 1):
            table.rows[i].cells[0].text = role
    
    def _add_table(self, doc, rows, cols, heading_row=True):
        """Create styled table"""
        table = doc.add_table(rows=rows, cols=cols)
        table.style = 'Light Grid Accent 1'
        
        if heading_row and rows > 0:
            for cell in table.rows[0].cells:
                if cell.paragraphs[0].runs:
                    cell.paragraphs[0].runs[0].font.bold = True
                shading = OxmlElement('w:shd')
                shading.set(qn('w:fill'), 'D9E2F3')
                cell._element.get_or_add_tcPr().append(shading)
        
        return table
    
    def _add_iso_ref(self, doc, clause):
        """Add ISO reference"""
        p = doc.add_paragraph()
        p.style = 'FSC_ISO'
        p.add_run(f"[{clause}]")


# Convenience function for direct use
def generate_fsc_word(
    system_name: str,
    goals_data: List[Dict],
    fsrs_data: List[Dict],
    strategies_data: Optional[Dict] = None,
    fsc_data: Optional[Dict] = None,
    plugin_folder: Optional[str] = None
) -> Document:
    """
    Generate FSC Word document.
    
    Convenience function that creates generator and generates document.
    """
    generator = FSCWordGenerator(plugin_folder)

    
    return generator.generate(
        system_name, goals_data, fsrs_data, strategies_data, fsc_data
    )