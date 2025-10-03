# hook_formatter.py - Main hook entrypoint for document formatting
from cat.mad_hatter.decorators import hook
from cat.log import log
import os
from datetime import datetime
from .item_definition_dev_doc import create_item_definition_docx
from .item_definition_rev_doc import create_review_docx
from .item_definition_rev_xls import create_review_excel
from .hara_dev_xls import create_hara_excel
from .utils import parse_review_content, detect_document_type

@hook(priority=1)
def before_cat_sends_message(message, cat):
    """
    Main hook to detect LLM output type and format accordingly.
    Saves files to organized folders based on document type.
    
    CRITICAL: Only formats complete documents, not intermediate workflow steps.
    """
    content = message.get("content", "")
    
    # Detect document type from content or working memory
    doc_type = detect_document_type(content, cat.working_memory)
    
    # CRITICAL: Check HARA workflow stage before attempting to format
    hara_stage = cat.working_memory.get("hara_stage", "")
    
    # Only format HARA if workflow is complete (Step 4 or Step 5)
    if doc_type == "hara":
        if hara_stage not in ["table_generated", "safety_goals_derived"]:
            log.info(f"HARA workflow incomplete (stage: {hara_stage}). Skipping document generation.")
            return message  # Don't format yet
    
    if not doc_type:
        return message
    
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        plugin_folder = os.path.dirname(__file__)
        
        # Check if this is a template
        is_template = cat.working_memory.get("is_template", False)
        
        # Determine output directory based on type
        if is_template:
            output_dir = os.path.join(plugin_folder, "generated_documents", "00_Templates")
        elif doc_type == "item_definition":
            output_dir = os.path.join(plugin_folder, "generated_documents", "01_Item_Definition")
        elif doc_type == "item_definition_review":
            output_dir = os.path.join(plugin_folder, "generated_documents", "02_Item_Definition_Review_Checklist_Report")
        elif doc_type == "hara":
            output_dir = os.path.join(plugin_folder, "generated_documents", "03_HARA")
        elif doc_type == "safety_goals":
            output_dir = os.path.join(plugin_folder, "generated_documents", "04_Safety_Goals")
        else:
            return message  # Unknown type, skip
        
        # Create directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        if doc_type == "item_definition":
            filename = format_item_definition(content, plugin_folder, output_dir, 
                                            timestamp, is_template)
            if filename:
                doc_type_str = "Template" if is_template else "Item Definition"
                folder_name = "00_Templates" if is_template else "01_Item_Definition"
                message["content"] += f"\n\n📄 *{doc_type_str} saved:* `generated_documents/{folder_name}/{filename}`"
                log.info(f"✅ {doc_type_str} formatted: {filename}")
        
        elif doc_type == "item_definition_review":
            filenames = format_review(content, plugin_folder, output_dir, 
                                    timestamp, is_template)
            if filenames:
                file_list = "\n- ".join(filenames)
                doc_type_str = "Review templates" if is_template else "Review documents"
                folder_name = "00_Templates" if is_template else "02_Item_Definition_Review_Checklist_Report"
                message["content"] += f"\n\n📄 *{doc_type_str} generated in `generated_documents/{folder_name}/`:*\n- {file_list}"
                log.info(f"✅ {doc_type_str} formatted: {filenames}")
        
        elif doc_type == "hara" and hara_stage == "table_generated":
            # Only format HARA after Step 4 is complete
            filenames = format_hara_table(content, plugin_folder, output_dir, 
                                         timestamp, cat.working_memory)
            if filenames:
                file_list = "\n- ".join(filenames)
                message["content"] += f"\n\n📊 *HARA documents generated in `generated_documents/03_HARA/`:*\n- {file_list}"
                log.info(f"✅ HARA formatted: {filenames}")
        
        elif doc_type == "safety_goals" and hara_stage == "safety_goals_derived":
            # Only format safety goals after Step 5 is complete
            filenames = format_safety_goals(content, plugin_folder, output_dir,
                                           timestamp, cat.working_memory)
            if filenames:
                file_list = "\n- ".join(filenames)
                message["content"] += f"\n\n📋 *Safety Goals documents generated in `generated_documents/04_Safety_Goals/`:*\n- {file_list}"
                log.info(f"✅ Safety Goals formatted: {filenames}")
        
        cleanup_working_memory(cat.working_memory)
        
    except Exception as e:
        log.error(f"❌ Document formatting failed: {e}")
        # Don't add error to user message - fail silently for better UX
        log.error(f"Formatting error details: {str(e)}")
    
    return message


def format_item_definition(content, plugin_folder, output_dir, timestamp, is_template=False):
    """Format Item Definition content into Word document."""
    try:
        system_name = extract_system_name(content)
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in system_name).replace(" ", "_")
        
        prefix = "TEMPLATE_" if is_template else ""
        filename = f"{prefix}ItemDefinition_{safe_name}_{timestamp}.docx"
        filepath = os.path.join(output_dir, filename)
        
        doc = create_item_definition_docx(content, plugin_folder, system_name)
        doc.save(filepath)
        
        return filename
    except Exception as e:
        log.error(f"Item definition formatting error: {e}")
        return None


def format_review(content, plugin_folder, output_dir, timestamp, is_template=False):
    """Format Review content into Word and Excel documents."""
    try:
        reviews = parse_review_content(content)
        if not reviews:
            log.warning("No review items found in content")
            return None
        
        filenames = []
        prefix = "TEMPLATE_" if is_template else ""
        
        # Create Word document
        doc = create_review_docx(reviews, plugin_folder, timestamp)
        docx_filename = f"{prefix}ItemDefinition_Review_{timestamp}.docx"
        docx_path = os.path.join(output_dir, docx_filename)
        doc.save(docx_path)
        filenames.append(f"Word: {docx_filename}")
        
        # Create Excel file
        wb = create_review_excel(reviews, timestamp)
        excel_filename = f"{prefix}ItemDefinition_Review_{timestamp}.xlsx"
        excel_path = os.path.join(output_dir, excel_filename)
        wb.save(excel_path)
        filenames.append(f"Excel: {excel_filename}")
        
        return filenames
    except Exception as e:
        log.error(f"Review formatting error: {e}")
        return None


def format_hara_table(content, plugin_folder, output_dir, timestamp, working_memory):
    """Format HARA table into Excel workbook."""
    try:
        # Import HARA formatters (create these if they don't exist)
        from .hara_dev_xls import create_hara_excel
        
        hara_table = working_memory.get("hara_table", "")
        hazop_analysis = working_memory.get("hazop_analysis", "")
        exposure_assessments = working_memory.get("exposure_assessments", "")
        item_name = working_memory.get("hara_item_name", "Unknown_System")
        
        if not hara_table:
            log.warning("No HARA table found in working memory")
            return None
        
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in item_name).replace(" ", "_")
        
        filename = f"HARA_{safe_name}_{timestamp}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        # Create Excel workbook
        wb = create_hara_excel(hara_table, hazop_analysis, exposure_assessments, item_name, timestamp)
        wb.save(filepath)
        
        return [f"Excel: {filename}"]
    except ImportError:
        log.warning("HARA Excel formatter not available yet")
        return None
    except Exception as e:
        log.error(f"HARA formatting error: {e}")
        return None


def format_safety_goals(content, plugin_folder, output_dir, timestamp, working_memory):
    """Format safety goals into Word document."""
    try:
        # Import safety goals formatter (create if it doesn't exist)
        from .safety_goals_formatter_doc import create_safety_goals_docx
        
        safety_goals_doc = working_memory.get("safety_goals_document", "")
        item_name = working_memory.get("hara_item_name", "Unknown_System")
        
        if not safety_goals_doc:
            log.warning("No safety goals document found in working memory")
            return None
        
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" 
                           for c in item_name).replace(" ", "_")
        
        filename = f"SafetyGoals_{safe_name}_{timestamp}.docx"
        filepath = os.path.join(output_dir, filename)
        
        # Create Word document
        doc = create_safety_goals_docx(safety_goals_doc, plugin_folder, item_name)
        doc.save(filepath)
        
        return [f"Word: {filename}"]
    except ImportError:
        log.warning("Safety Goals Word formatter not available yet")
        return None
    except Exception as e:
        log.error(f"Safety Goals formatting error: {e}")
        return None


def extract_system_name(content):
    """Extract system name from Item Definition content."""
    lines = content.split("\n")
    first_line = lines[0] if lines else "Unknown System"
    
    if ": " in first_line:
        return first_line.split(": ", 1)[1].strip()
    
    return "Unknown System"


def cleanup_working_memory(working_memory):
    """Clean up working memory keys used for document formatting."""
    # Don't cleanup hara_stage - needed for workflow progression
    keys_to_remove = ["document_type", "is_template"]
    for key in keys_to_remove:
        if key in working_memory:
            del working_memory[key]