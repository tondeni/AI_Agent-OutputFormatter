# ==============================================================================
# AI_Agent-OutputFormatter/generators/__init__.py
# Package initialization for document generators
# ==============================================================================

"""
Document generators for ISO 26262 work products.

This package contains generators for various output formats:
- Excel workbooks for FSR listings and allocation matrices
- Word documents for FSC reports
- PDF reports (future)
"""

# Import main generator functions for easier access
from .Functional_Safety_Concept.fsr_excel_generator import generate_fsr_excel
from .Functional_Safety_Concept.fsc_excel_generator import FSCExcelGenerator, generate_fsc_excel
from .Functional_Safety_Concept.fsc_word_generator import FSCWordGenerator

# Optional: Import other generators when you create them
# from .fsc_word_generator import generate_fsc_word
# from .hara_excel_generator import generate_hara_excel

# Define what's available when someone does: from generators import *
__all__ = [
    'generate_fsr_excel',
    'generate_fsc_excel',
    FSCWordGenerator
]

# Package metadata
__version__ = '1.0.0'
__author__ = 'Functional Safety Team'