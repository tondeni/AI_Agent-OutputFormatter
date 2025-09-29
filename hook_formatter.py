# hook_formatter.py - Main hook entrypoint for document formatting
from cat.mad_hatter.decorators import hook
from cat.log import log
import os
from datetime import datetime
# from .claude_item_definition_dev_doc import create_item_definition_docx
from .item_definition_rev_doc import create_review_docx
from .item_definition_rev_xls import create_review_excel
from .utils import parse_review_content, detect_document_type

@hook(priority=1)
def before_cat_sends_message(message, cat):
    """
    Main hook to detect LLM output type and format accordingly.
    Supports both Item Definition development and review formatting.
    """
    content = message.get("content", "")
    
    # Detect document type from content or working memory
    doc_type = detect_document_type(content, cat.working_memory)
    
    if not doc_type:
        return message  # Not a document we can format
    
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        plugin_folder = os.path.dirname(__file__)
        output_dir = os.path.join(plugin_folder, "generated_documents")
        os.makedirs(output_dir, exist_ok=True)
        
        if doc_type == "item_definition":
            # Handle Item Definition development formatting
            filename = format_item_definition(content, plugin_folder, output_dir, timestamp)
            if filename:
                message["content"] += f"\n\n📄 *Item Definition saved:* `{filename}`"
                log.info(f"✅ Item Definition formatted: {filename}")
        
        elif doc_type == "item_definition_review":
            # Handle Item Definition review formatting
            filenames = format_review(content, plugin_folder, output_dir, timestamp)
            if filenames:
                file_list = "\n- ".join(filenames)
                message.content += f"\n\n📄 *Review documents generated:*\n- {file_list}"
                log.info(f"✅ Review documents formatted: {filenames}")
        
        # Clean up working memory
        cleanup_working_memory(cat.working_memory)
        
    except Exception as e:
        log.error(f"❌ Document formatting failed: {e}")
        message["content"] += f"\n\n⚠️ *Document formatting error:* {str(e)}"
    
    return message

def format_item_definition(content, plugin_folder, output_dir, timestamp):
    """Format Item Definition content into Word document."""
    try:
        # Extract system name for filename
        system_name = extract_system_name(content)
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" for c in system_name).replace(" ", "_")
        
        filename = f"ItemDefinition_{safe_name}_{timestamp}.docx"
        filepath = os.path.join(output_dir, filename)
        
        doc = create_item_definition_docx(content, plugin_folder, system_name)
        doc.save(filepath)
        
        return filename
    except Exception as e:
        log.error(f"Item definition formatting error: {e}")
        return None

def format_review(content, plugin_folder, output_dir, timestamp):
    """Format Review content into Word and Excel documents."""
    try:
        reviews = parse_review_content(content)
        if not reviews:
            log.warning("No review items found in content")
            return None
        
        filenames = []
        
        # Create Word document
        doc = create_review_docx(reviews, plugin_folder, timestamp)
        docx_filename = f"ItemDefinition_Review_{timestamp}.docx"
        docx_path = os.path.join(output_dir, docx_filename)
        doc.save(docx_path)
        filenames.append(f"Word: {docx_filename}")
        
        # Create Excel file
        wb = create_review_excel(reviews, timestamp)
        excel_filename = f"ItemDefinition_Review_{timestamp}.xlsx"
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
    keys_to_remove = ["document_type", "reviewed_item", "system_name"]
    for key in keys_to_remove:
        if key in working_memory:
            del working_memory[key]