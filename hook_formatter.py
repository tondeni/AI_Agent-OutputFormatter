# hook_formatter.py - Main hook entrypoint for document formatting
from cat.mad_hatter.decorators import hook
from cat.log import log
import os
from datetime import datetime
from .item_definition_dev_doc import create_item_definition_docx
from .item_definition_rev_doc import create_review_docx
from .item_definition_rev_xls import create_review_excel
from .utils import parse_review_content, detect_document_type

@hook(priority=1)
def before_cat_sends_message(message, cat):
    """
    Main hook to detect LLM output type and format accordingly.
    Saves files to organized folders based on document type.
    """
    content = message.get("content", "")
    
    # Detect document type from content or working memory
    doc_type = detect_document_type(content, cat.working_memory)
    
    if not doc_type:
        return message
    
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        plugin_folder = os.path.dirname(__file__)
        
        # Check if this is a template
        is_template = cat.working_memory.get("is_template", False)
        
        # Determine output directory based on type
        if is_template:
            output_dir = os.path.join(plugin_folder, "generated_documents", "templates")
        elif doc_type == "item_definition":
            output_dir = os.path.join(plugin_folder, "generated_documents", "item_definition_work_product")
        elif doc_type == "item_definition_review":
            output_dir = os.path.join(plugin_folder, "generated_documents", "item_definition_review_checklist_report")
        
        # Create directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        if doc_type == "item_definition":
            filename = format_item_definition(content, plugin_folder, output_dir, 
                                            timestamp, is_template)
            if filename:
                doc_type_str = "Template" if is_template else "Item Definition"
                folder_name = "templates" if is_template else "item_definition_work_product"
                message.content += f"\n\n📄 *{doc_type_str} saved:* `{folder_name}/{filename}`"
                log.info(f"✅ {doc_type_str} formatted: {filename}")
        
        elif doc_type == "item_definition_review":
            filenames = format_review(content, plugin_folder, output_dir, 
                                    timestamp, is_template)
            if filenames:
                file_list = "\n- ".join(filenames)
                doc_type_str = "Review templates" if is_template else "Review documents"
                folder_name = "templates" if is_template else "item_definition_review_checklist_report"
                message["content"] += f"\n\n📄 *{doc_type_str} generated in `{folder_name}/`:*\n- {file_list}"
                log.info(f"✅ {doc_type_str} formatted: {filenames}")
        
        cleanup_working_memory(cat.working_memory)
        
    except Exception as e:
        log.error(f"❌ Document formatting failed: {e}")
        message["content"] += f"\n\n⚠️ *Document formatting error:* {str(e)}"
    
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

def extract_system_name(content):
    """Extract system name from Item Definition content."""
    lines = content.split("\n")
    first_line = lines[0] if lines else "Unknown System"
    
    if ": " in first_line:
        return first_line.split(": ", 1)[1].strip()
    
    return "Unknown System"

def cleanup_working_memory(working_memory):
    """Clean up working memory keys used for document formatting."""
    keys_to_remove = ["document_type", "reviewed_item", "system_name", "is_template"]
    for key in keys_to_remove:
        if key in working_memory:
            del working_memory[key]