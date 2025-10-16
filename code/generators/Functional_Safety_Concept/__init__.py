# ==============================================================================
# AI_Agent-OutputFormatter/generators/Functional_Safety_Concept/__init__.py
# Subpackage for Functional Safety Concept generators
# ==============================================================================

"""
Generators for Functional Safety Concept (FSC) related documents.
"""

# Import specific classes/functions for easier access
from .fsr_excel_generator import generate_fsr_excel
from .fsc_excel_generator import FSCExcelGenerator, generate_fsc_excel
from .fsc_word_generator import FSCWordGenerator, generate_fsc_word

# Define what's available when someone does: from generators.Functional_Safety_Concept import *
__all__ = [
    'generate_fsr_excel',
    'generate_fsc_excel',
    'generate_fsc_word',
    'FSCWordGenerator',
]

# Package metadata
__version__ = '1.0.0'
__author__ = 'Functional Safety Team'