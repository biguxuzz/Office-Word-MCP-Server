"""
Content tools for Word Document Server.

These tools add various types of content to Word documents,
including headings, paragraphs, tables, images, and page breaks.
"""
import os
from typing import List, Optional, Dict, Any
from docx import Document
from docx.shared import Inches, Pt, RGBColor

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension
from word_document_server.utils.document_utils import find_and_replace_text, insert_header_near_text, insert_numbered_list_near_text, insert_line_or_paragraph_near_text, replace_paragraph_block_below_header, replace_block_between_manual_anchors
from word_document_server.core.styles import ensure_heading_style, ensure_table_style
from word_document_server.utils.track_changes_utils import (
    open_document_with_track_changes, close_document_with_track_changes,
    check_pywin32_available, get_track_changes_error_message
)


async def add_heading(filename: str, text: str, level: int = 1,
                      font_name: Optional[str] = None, font_size: Optional[int] = None,
                      bold: Optional[bool] = None, italic: Optional[bool] = None,
                      border_bottom: bool = False, track_changes: bool = False,
                      change_author: str = "") -> str:
    """Add a heading to a Word document with optional formatting.

    Args:
        filename: Path to the Word document
        text: Heading text
        level: Heading level (1-9, where 1 is the highest level)
        font_name: Font family (e.g., 'Helvetica')
        font_size: Font size in points (e.g., 14)
        bold: True/False for bold text
        italic: True/False for italic text
        border_bottom: True to add bottom border (for section headers)
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
    """
    filename = ensure_docx_extension(filename)

    # Ensure level is converted to integer
    try:
        level = int(level)
    except (ValueError, TypeError):
        return "Invalid parameter: level must be an integer between 1 and 9"

    # Validate level range
    if level < 1 or level > 9:
        return f"Invalid heading level: {level}. Level must be between 1 and 9."

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        # Suggest creating a copy
        return f"Cannot modify document: {error_message}. Consider creating a copy first or creating a new document."

    # Check if Track Changes is requested
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        
        try:
            word, doc, saved_state = open_document_with_track_changes(filename, change_author)
            
            # Убеждаемся, что Track Changes включен
            if not doc.TrackRevisions:
                doc.TrackRevisions = True
            
            # Устанавливаем автора изменений
            if change_author:
                word.UserName = change_author
            
            # Add heading via COM API at the end of document
            range_obj = doc.Range()
            range_obj.Collapse(0)  # wdCollapseEnd = 0
            range_obj.InsertAfter(text + "\r")
            range_obj.Collapse(0)
            
            # Set heading style
            style_name = f"Heading {level}"
            try:
                range_obj.Style = style_name
            except:
                # If style doesn't exist, format manually
                range_obj.Font.Bold = True
                if level == 1:
                    range_obj.Font.Size = 16
                elif level == 2:
                    range_obj.Font.Size = 14
                else:
                    range_obj.Font.Size = 12
            
            # Apply formatting
            if font_name:
                range_obj.Font.Name = font_name
            if font_size:
                range_obj.Font.Size = font_size
            if bold is not None:
                range_obj.Font.Bold = bold
            if italic is not None:
                range_obj.Font.Italic = italic
            
            # Add bottom border if requested
            if border_bottom:
                range_obj.Borders(10).LineStyle = 1  # wdLineStyleSingle = 1
                range_obj.Borders(10).LineWidth = 4
            
            close_document_with_track_changes(word, doc, saved_state)
            return f"Heading '{text}' (level {level}) added to {filename} with Track Changes"
        except Exception as e:
            return f"Failed to add heading with Track Changes: {str(e)}"

    try:
        doc = Document(filename)

        # Ensure heading styles exist
        ensure_heading_style(doc)

        # Try to add heading with style
        try:
            heading = doc.add_heading(text, level=level)
        except Exception as style_error:
            # If style-based approach fails, use direct formatting
            heading = doc.add_paragraph(text)
            heading.style = doc.styles['Normal']
            if heading.runs:
                run = heading.runs[0]
                run.bold = True
                # Adjust size based on heading level
                if level == 1:
                    run.font.size = Pt(16)
                elif level == 2:
                    run.font.size = Pt(14)
                else:
                    run.font.size = Pt(12)

        # Apply formatting to all runs in the heading
        if any([font_name, font_size, bold is not None, italic is not None]):
            for run in heading.runs:
                if font_name:
                    run.font.name = font_name
                if font_size:
                    run.font.size = Pt(font_size)
                if bold is not None:
                    run.font.bold = bold
                if italic is not None:
                    run.font.italic = italic

        # Add bottom border if requested
        if border_bottom:
            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn

            pPr = heading._element.get_or_add_pPr()
            pBdr = OxmlElement('w:pBdr')

            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '4')  # 0.5pt border
            bottom.set(qn('w:space'), '0')
            bottom.set(qn('w:color'), '000000')

            pBdr.append(bottom)
            pPr.append(pBdr)

        doc.save(filename)
        return f"Heading '{text}' (level {level}) added to {filename}"
    except Exception as e:
        return f"Failed to add heading: {str(e)}"


async def add_paragraph(filename: str, text: str, style: Optional[str] = None,
                        font_name: Optional[str] = None, font_size: Optional[int] = None,
                        bold: Optional[bool] = None, italic: Optional[bool] = None,
                        color: Optional[str] = None, track_changes: bool = False,
                        change_author: str = "") -> str:
    """Add a paragraph to a Word document with optional formatting.

    Args:
        filename: Path to the Word document
        text: Paragraph text
        style: Optional paragraph style name
        font_name: Font family (e.g., 'Helvetica', 'Times New Roman')
        font_size: Font size in points (e.g., 14, 36)
        bold: True/False for bold text
        italic: True/False for italic text
        color: RGB color as hex string (e.g., '000000' for black)
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        # Suggest creating a copy
        return f"Cannot modify document: {error_message}. Consider creating a copy first or creating a new document."

    # Check if Track Changes is requested
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        
        try:
            word, doc, saved_state = open_document_with_track_changes(filename, change_author)
            
            # Убеждаемся, что Track Changes включен
            if not doc.TrackRevisions:
                doc.TrackRevisions = True
            
            # Устанавливаем автора изменений
            if change_author:
                word.UserName = change_author
            
            # Add paragraph via COM API at the end of document
            range_obj = doc.Range()
            range_obj.Collapse(0)  # wdCollapseEnd = 0
            range_obj.InsertAfter(text + "\r")
            range_obj.Collapse(0)
            
            # Apply style if provided
            if style:
                try:
                    range_obj.Style = style
                except:
                    pass  # Style doesn't exist, continue without it
            
            # Apply formatting
            if font_name:
                range_obj.Font.Name = font_name
            if font_size:
                range_obj.Font.Size = font_size
            if bold is not None:
                range_obj.Font.Bold = bold
            if italic is not None:
                range_obj.Font.Italic = italic
            if color:
                color_hex = color.lstrip('#')
                # Convert hex to RGB
                r = int(color_hex[0:2], 16)
                g = int(color_hex[2:4], 16)
                b = int(color_hex[4:6], 16)
                range_obj.Font.Color = (r << 16) | (g << 8) | b
            
            close_document_with_track_changes(word, doc, saved_state)
            return f"Paragraph added to {filename} with Track Changes"
        except Exception as e:
            return f"Failed to add paragraph with Track Changes: {str(e)}"

    try:
        doc = Document(filename)
        paragraph = doc.add_paragraph(text)

        if style:
            try:
                paragraph.style = style
            except KeyError:
                # Style doesn't exist, use normal and report it
                paragraph.style = doc.styles['Normal']
                doc.save(filename)
                return f"Style '{style}' not found, paragraph added with default style to {filename}"

        # Apply formatting to all runs in the paragraph
        if any([font_name, font_size, bold is not None, italic is not None, color]):
            for run in paragraph.runs:
                if font_name:
                    run.font.name = font_name
                if font_size:
                    run.font.size = Pt(font_size)
                if bold is not None:
                    run.font.bold = bold
                if italic is not None:
                    run.font.italic = italic
                if color:
                    # Remove any '#' prefix if present
                    color_hex = color.lstrip('#')
                    run.font.color.rgb = RGBColor.from_string(color_hex)

        doc.save(filename)
        return f"Paragraph added to {filename}"
    except Exception as e:
        return f"Failed to add paragraph: {str(e)}"


async def add_table(filename: str, rows: int, cols: int, data: Optional[List[List[str]]] = None,
                    track_changes: bool = False, change_author: str = "") -> str:
    """Add a table to a Word document.
    
    Args:
        filename: Path to the Word document
        rows: Number of rows in the table
        cols: Number of columns in the table
        data: Optional 2D array of data to fill the table
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        # Suggest creating a copy
        return f"Cannot modify document: {error_message}. Consider creating a copy first or creating a new document."

    # Check if Track Changes is requested
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        
        try:
            word, doc, saved_state = open_document_with_track_changes(filename, change_author)
            
            # Убеждаемся, что Track Changes включен
            if not doc.TrackRevisions:
                doc.TrackRevisions = True
            
            # Устанавливаем автора изменений
            if change_author:
                word.UserName = change_author
            
            # Add table via COM API at the end of document
            range_obj = doc.Range()
            range_obj.Collapse(0)  # wdCollapseEnd = 0
            table = doc.Tables.Add(range_obj, rows, cols)
            
            # Fill table with data if provided
            if data:
                for i, row_data in enumerate(data):
                    if i >= rows:
                        break
                    for j, cell_text in enumerate(row_data):
                        if j >= cols:
                            break
                        table.Cell(i + 1, j + 1).Range.Text = str(cell_text)
            
            close_document_with_track_changes(word, doc, saved_state)
            return f"Table ({rows}x{cols}) added to {filename} with Track Changes"
        except Exception as e:
            return f"Failed to add table with Track Changes: {str(e)}"

    try:
        doc = Document(filename)
        table = doc.add_table(rows=rows, cols=cols)
        
        # Try to set the table style
        try:
            table.style = 'Table Grid'
        except KeyError:
            # If style doesn't exist, add basic borders
            pass
        
        # Fill table with data if provided
        if data:
            for i, row_data in enumerate(data):
                if i >= rows:
                    break
                for j, cell_text in enumerate(row_data):
                    if j >= cols:
                        break
                    table.cell(i, j).text = str(cell_text)
        
        doc.save(filename)
        return f"Table ({rows}x{cols}) added to {filename}"
    except Exception as e:
        return f"Failed to add table: {str(e)}"


async def add_picture(filename: str, image_path: str, width: Optional[float] = None,
                      track_changes: bool = False, change_author: str = "") -> str:
    """Add an image to a Word document.
    
    Args:
        filename: Path to the Word document
        image_path: Path to the image file
        width: Optional width in inches (proportional scaling)
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
    """
    filename = ensure_docx_extension(filename)
    
    # Validate document existence
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Get absolute paths for better diagnostics
    abs_filename = os.path.abspath(filename)
    abs_image_path = os.path.abspath(image_path)
    
    # Validate image existence with improved error message
    if not os.path.exists(abs_image_path):
        return f"Image file not found: {abs_image_path}"
    
    # Check image file size
    try:
        image_size = os.path.getsize(abs_image_path) / 1024  # Size in KB
        if image_size <= 0:
            return f"Image file appears to be empty: {abs_image_path} (0 KB)"
    except Exception as size_error:
        return f"Error checking image file: {str(size_error)}"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(abs_filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first or creating a new document."

    # Check if Track Changes is requested
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        
        try:
            word, doc, saved_state = open_document_with_track_changes(abs_filename, change_author)
            
            # Убеждаемся, что Track Changes включен
            if not doc.TrackRevisions:
                doc.TrackRevisions = True
            
            # Устанавливаем автора изменений
            if change_author:
                word.UserName = change_author
            
            # Add picture via COM API at the end of document
            range_obj = doc.Range()
            range_obj.Collapse(0)  # wdCollapseEnd = 0
            
            if width:
                # Convert inches to points (1 inch = 72 points)
                width_points = width * 72
                doc.InlineShapes.AddPicture(abs_image_path, False, True, range_obj).Width = width_points
            else:
                doc.InlineShapes.AddPicture(abs_image_path, False, True, range_obj)
            
            close_document_with_track_changes(word, doc, saved_state)
            return f"Picture {image_path} added to {filename} with Track Changes"
        except Exception as e:
            return f"Failed to add picture with Track Changes: {str(e)}"

    try:
        doc = Document(abs_filename)
        # Additional diagnostic info
        diagnostic = f"Attempting to add image ({abs_image_path}, {image_size:.2f} KB) to document ({abs_filename})"
        
        try:
            if width:
                doc.add_picture(abs_image_path, width=Inches(width))
            else:
                doc.add_picture(abs_image_path)
            doc.save(abs_filename)
            return f"Picture {image_path} added to {filename}"
        except Exception as inner_error:
            # More detailed error for the specific operation
            error_type = type(inner_error).__name__
            error_msg = str(inner_error)
            return f"Failed to add picture: {error_type} - {error_msg or 'No error details available'}\nDiagnostic info: {diagnostic}"
    except Exception as outer_error:
        # Fallback error handling
        error_type = type(outer_error).__name__
        error_msg = str(outer_error)
        return f"Document processing error: {error_type} - {error_msg or 'No error details available'}"


async def add_page_break(filename: str, track_changes: bool = False,
                         change_author: str = "") -> str:
    """Add a page break to the document.
    
    Args:
        filename: Path to the Word document
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."

    # Check if Track Changes is requested
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        
        try:
            word, doc, saved_state = open_document_with_track_changes(filename, change_author)
            
            # Убеждаемся, что Track Changes включен
            if not doc.TrackRevisions:
                doc.TrackRevisions = True
            
            # Устанавливаем автора изменений
            if change_author:
                word.UserName = change_author
            
            # Add page break via COM API at the end of document
            range_obj = doc.Range()
            range_obj.Collapse(0)  # wdCollapseEnd = 0
            range_obj.InsertBreak(7)  # wdPageBreak = 7
            
            close_document_with_track_changes(word, doc, saved_state)
            return f"Page break added to {filename} with Track Changes."
        except Exception as e:
            return f"Failed to add page break with Track Changes: {str(e)}"

    try:
        doc = Document(filename)
        doc.add_page_break()
        doc.save(filename)
        return f"Page break added to {filename}."
    except Exception as e:
        return f"Failed to add page break: {str(e)}"


async def add_table_of_contents(filename: str, title: str = "Table of Contents", max_level: int = 3) -> str:
    """Add a table of contents to a Word document based on heading styles.
    
    Args:
        filename: Path to the Word document
        title: Optional title for the table of contents
        max_level: Maximum heading level to include (1-9)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."
    
    try:
        # Ensure max_level is within valid range
        max_level = max(1, min(max_level, 9))
        
        doc = Document(filename)
        
        # Collect headings and their positions
        headings = []
        for i, paragraph in enumerate(doc.paragraphs):
            # Check if paragraph style is a heading
            if paragraph.style and paragraph.style.name.startswith('Heading '):
                try:
                    # Extract heading level from style name
                    level = int(paragraph.style.name.split(' ')[1])
                    if level <= max_level:
                        headings.append({
                            'level': level,
                            'text': paragraph.text,
                            'position': i
                        })
                except (ValueError, IndexError):
                    # Skip if heading level can't be determined
                    pass
        
        if not headings:
            return f"No headings found in document {filename}. Table of contents not created."
        
        # Create a new document with the TOC
        toc_doc = Document()
        
        # Add title
        if title:
            toc_doc.add_heading(title, level=1)
        
        # Add TOC entries
        for heading in headings:
            # Indent based on level (using tab characters)
            indent = '    ' * (heading['level'] - 1)
            toc_doc.add_paragraph(f"{indent}{heading['text']}")
        
        # Add page break
        toc_doc.add_page_break()
        
        # Get content from original document
        for paragraph in doc.paragraphs:
            p = toc_doc.add_paragraph(paragraph.text)
            # Copy style if possible
            try:
                if paragraph.style:
                    p.style = paragraph.style.name
            except:
                pass
        
        # Copy tables
        for table in doc.tables:
            # Create a new table with the same dimensions
            new_table = toc_doc.add_table(rows=len(table.rows), cols=len(table.columns))
            # Copy cell contents
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    for paragraph in cell.paragraphs:
                        new_table.cell(i, j).text = paragraph.text
        
        # Save the new document with TOC
        toc_doc.save(filename)
        
        return f"Table of contents with {len(headings)} entries added to {filename}"
    except Exception as e:
        return f"Failed to add table of contents: {str(e)}"


async def delete_paragraph(filename: str, paragraph_index: int,
                          track_changes: bool = False, change_author: str = "") -> str:
    """Delete a paragraph from a document.
    
    Args:
        filename: Path to the Word document
        paragraph_index: Index of the paragraph to delete (0-based)
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."

    # Check if Track Changes is requested
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        
        try:
            word, doc, saved_state = open_document_with_track_changes(filename, change_author)
            
            # Убеждаемся, что Track Changes включен
            if not doc.TrackRevisions:
                doc.TrackRevisions = True
            
            # Устанавливаем автора изменений
            if change_author:
                word.UserName = change_author
            
            # Validate paragraph index (COM API uses 1-based indexing)
            if paragraph_index < 0 or paragraph_index >= doc.Paragraphs.Count:
                close_document_with_track_changes(word, doc, saved_state, save_changes=False)
                return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs (0-{doc.Paragraphs.Count-1})."
            
            # Delete paragraph via COM API (1-based index)
            # Удаление через Range.Delete() создаст ревизию при включенном TrackRevisions
            para = doc.Paragraphs(paragraph_index + 1)
            para.Range.Delete()
            
            close_document_with_track_changes(word, doc, saved_state)
            return f"Paragraph at index {paragraph_index} deleted successfully with Track Changes."
        except Exception as e:
            return f"Failed to delete paragraph with Track Changes: {str(e)}"

    try:
        doc = Document(filename)
        
        # Validate paragraph index
        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            return f"Invalid paragraph index. Document has {len(doc.paragraphs)} paragraphs (0-{len(doc.paragraphs)-1})."
        
        # Delete the paragraph (by removing its content and setting it empty)
        # Note: python-docx doesn't support true paragraph deletion, this is a workaround
        paragraph = doc.paragraphs[paragraph_index]
        p = paragraph._p
        p.getparent().remove(p)
        
        doc.save(filename)
        return f"Paragraph at index {paragraph_index} deleted successfully."
    except Exception as e:
        return f"Failed to delete paragraph: {str(e)}"


async def search_and_replace(filename: str, find_text: str, replace_text: str,
                             track_changes: bool = False, change_author: str = "",
                             occurrence_index: Optional[int] = None,
                             context_before: str = "", context_after: str = "") -> str:
    """Search for text and replace all occurrences or a specific occurrence.
    
    Args:
        filename: Path to the Word document
        find_text: Text to search for
        replace_text: Text to replace with
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
        occurrence_index: Optional index of specific occurrence to replace (1-based, e.g., 1 for first, 2 for second).
                         If None, replaces all occurrences.
        context_before: Optional text that should appear before the match (for precise identification)
        context_after: Optional text that should appear after the match (for precise identification)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."

    # Check if Track Changes is requested
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        
        try:
            word, doc, saved_state = open_document_with_track_changes(filename, change_author)
            
            # Убеждаемся, что Track Changes включен
            if not doc.TrackRevisions:
                doc.TrackRevisions = True
            
            # Устанавливаем автора изменений
            if change_author:
                word.UserName = change_author
            
            # Perform find and replace via COM API
            # Используем подход: находим все вхождения сначала, затем заменяем их в обратном порядке
            # Это предотвращает смещение позиций при замене
            # ВАЖНО: Не используем find_obj.Range = ..., вместо этого создаем новый Range для каждого поиска
            
            # Находим все вхождения и сохраняем их позиции ДО замены
            found_positions = []
            search_start = doc.Content.Start
            search_end = doc.Content.End
            
            # Используем цикл для поиска всех вхождений
            # Уменьшено максимальное количество итераций для предотвращения зависания
            # Если документ очень большой, лучше использовать python-docx подход
            max_iterations = 5000  # Защита от бесконечного цикла
            iteration = 0
            last_search_start = search_start  # Для отслеживания прогресса
            no_progress_count = 0  # Счетчик итераций без прогресса
            
            while search_start < search_end and iteration < max_iterations:
                iteration += 1
                
                # Создаем новый Range для поиска
                search_range = doc.Range(search_start, search_end)
                
                # Получаем Find объект для этого диапазона
                find_obj = search_range.Find
                find_obj.ClearFormatting()
                find_obj.Replacement.ClearFormatting()
                find_obj.Text = find_text
                find_obj.Forward = True
                find_obj.Wrap = 0  # wdFindStop = 0 (остановиться в конце документа)
                
                # Выполняем поиск
                found = find_obj.Execute()
                
                if not found:
                    # Больше не найдено
                    break
                
                # Получаем найденный диапазон
                found_range = find_obj.Parent
                start_pos = found_range.Start
                end_pos = found_range.End
                
                # Если указан контекст, проверяем его
                matches_context = True
                if context_before or context_after:
                    # Получаем контекст вокруг найденного текста
                    # Расширяем диапазон для получения достаточного контекста
                    context_start = max(0, start_pos - max(len(context_before), 100) if context_before else start_pos)
                    context_end = min(doc.Content.End, end_pos + max(len(context_after), 100) if context_after else end_pos)
                    context_range = doc.Range(context_start, context_end)
                    context_text = context_range.Text
                    
                    # Вычисляем позицию найденного текста в контексте относительно начала контекста
                    # Это более точно, чем использование find(), который может найти другое вхождение
                    found_pos_in_context = start_pos - context_start
                    
                    # Проверяем, что найденный текст действительно находится в этом диапазоне
                    if found_pos_in_context < 0 or found_pos_in_context + len(find_text) > len(context_text):
                        matches_context = False
                    else:
                        # Проверяем, что текст в этой позиции действительно соответствует find_text
                        actual_text = context_text[found_pos_in_context:found_pos_in_context + len(find_text)]
                        if actual_text != find_text:
                            matches_context = False
                        else:
                            # Проверяем контекст до найденного текста
                            if context_before:
                                text_before = context_text[:found_pos_in_context]
                                if not text_before.endswith(context_before):
                                    matches_context = False
                            # Проверяем контекст после найденного текста
                            if context_after and matches_context:
                                text_after_start = found_pos_in_context + len(find_text)
                                text_after = context_text[text_after_start:]
                                if not text_after.startswith(context_after):
                                    matches_context = False
                
                # Сохраняем только если соответствует контексту (или если контекст не указан)
                if matches_context:
                    found_positions.append((start_pos, end_pos))
                
                # Перемещаем начало поиска после найденного текста
                # ВАЖНО: Если найденный текст пустой (end_pos == start_pos), продвигаемся минимум на 1 символ
                # чтобы избежать бесконечного цикла
                if end_pos > start_pos:
                    search_start = end_pos
                else:
                    # Пустой текст найден - продвигаемся на 1 символ вперед
                    search_start = start_pos + 1
                
                # Если позиция не изменилась или достигли конца документа, прекращаем поиск
                if search_start >= search_end:
                    break
                
                # Дополнительная защита: если позиция не продвинулась после нескольких итераций, прекращаем
                if iteration > 1 and search_start <= start_pos:
                    break
                
                # Проверка прогресса: если позиция не изменилась за последние 10 итераций, прекращаем
                if search_start == last_search_start:
                    no_progress_count += 1
                    if no_progress_count >= 10:
                        break
                else:
                    no_progress_count = 0  # Сбрасываем счетчик при прогрессе
                last_search_start = search_start
            
            if iteration >= max_iterations:
                if doc:
                    close_document_with_track_changes(word, doc, saved_state, save_changes=False)
                return f"Search stopped due to iteration limit ({max_iterations}). Found {len(found_positions)} matching occurrence(s) before stopping. Possible infinite loop detected or document is too large."
            
            if no_progress_count >= 10:
                if doc:
                    close_document_with_track_changes(word, doc, saved_state, save_changes=False)
                return f"Search stopped: no progress detected after {no_progress_count} iterations. Found {len(found_positions)} matching occurrence(s) before stopping. Possible infinite loop detected."
            
            # Определяем, какие вхождения заменять
            if occurrence_index is not None:
                # Заменяем только указанное вхождение (1-based индекс)
                if occurrence_index < 1 or occurrence_index > len(found_positions):
                    if doc:
                        close_document_with_track_changes(word, doc, saved_state, save_changes=False)
                    return f"Invalid occurrence_index: {occurrence_index}. Found {len(found_positions)} occurrence(s). Valid range: 1-{len(found_positions)}."
                # Заменяем только одно вхождение по индексу
                start_pos, end_pos = found_positions[occurrence_index - 1]
                replace_range = doc.Range(start_pos, end_pos)
                replace_range.Text = replace_text
                count = 1
            else:
                # Заменяем все вхождения в обратном порядке (с конца к началу)
                count = len(found_positions)
                for start_pos, end_pos in reversed(found_positions):
                    # Создаем диапазон для замены по сохраненным позициям
                    replace_range = doc.Range(start_pos, end_pos)
                    # Заменяем текст - это создаст ревизию при включенном TrackRevisions
                    replace_range.Text = replace_text
            
            # Сохраняем и закрываем документ
            close_document_with_track_changes(word, doc, saved_state)
            word = None
            doc = None
            
            if count > 0:
                return f"Replaced {count} occurrence(s) of '{find_text}' with '{replace_text}' with Track Changes."
            else:
                return f"No occurrences of '{find_text}' found."
        except Exception as e:
            # Гарантируем закрытие документа даже при ошибке
            error_msg = str(e)
            try:
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                    except:
                        pass
                if word:
                    try:
                        word.Quit()
                    except:
                        pass
            except:
                pass
            return f"Failed to search and replace with Track Changes: {error_msg}"

    try:
        doc = Document(filename)
        
        # Find all occurrences first (with context filtering if needed)
        occurrences = []
        
        # Search in paragraphs
        for para_idx, para in enumerate(doc.paragraphs):
            # Skip TOC paragraphs
            if para.style and para.style.name.startswith("TOC"):
                continue
            para_text = para.text
            if find_text in para_text:
                # Find all occurrences in this paragraph
                start = 0
                while True:
                    pos = para_text.find(find_text, start)
                    if pos == -1:
                        break
                    
                    # Check context if specified
                    matches_context = True
                    if context_before or context_after:
                        context_start = max(0, pos - len(context_before) if context_before else 0)
                        context_end = min(len(para_text), pos + len(find_text) + len(context_after) if context_after else len(para_text))
                        context_text = para_text[context_start:context_end]
                        
                        if context_before:
                            found_pos_in_context = context_text.find(find_text)
                            if found_pos_in_context == -1 or not context_text[:found_pos_in_context].endswith(context_before):
                                matches_context = False
                        if context_after and matches_context:
                            found_pos_in_context = context_text.find(find_text)
                            if found_pos_in_context == -1 or not context_text[found_pos_in_context + len(find_text):].startswith(context_after):
                                matches_context = False
                    
                    if matches_context:
                        occurrences.append({
                            'para_idx': para_idx,
                            'para': para,
                            'text_pos': pos,
                            'para_text': para_text
                        })
                    
                    start = pos + 1
        
        # Search in tables
        for table_idx, table in enumerate(doc.tables):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    for para_idx, para in enumerate(cell.paragraphs):
                        # Skip TOC paragraphs in tables
                        if para.style and para.style.name.startswith("TOC"):
                            continue
                        para_text = para.text
                        if find_text in para_text:
                            # Find all occurrences in this paragraph
                            start = 0
                            while True:
                                pos = para_text.find(find_text, start)
                                if pos == -1:
                                    break
                                
                                # Check context if specified
                                matches_context = True
                                if context_before or context_after:
                                    context_start = max(0, pos - len(context_before) if context_before else 0)
                                    context_end = min(len(para_text), pos + len(find_text) + len(context_after) if context_after else len(para_text))
                                    context_text = para_text[context_start:context_end]
                                    
                                    if context_before:
                                        found_pos_in_context = context_text.find(find_text)
                                        if found_pos_in_context == -1 or not context_text[:found_pos_in_context].endswith(context_before):
                                            matches_context = False
                                    if context_after and matches_context:
                                        found_pos_in_context = context_text.find(find_text)
                                        if found_pos_in_context == -1 or not context_text[found_pos_in_context + len(find_text):].startswith(context_after):
                                            matches_context = False
                                
                                if matches_context:
                                    occurrences.append({
                                        'para_idx': para_idx,
                                        'para': para,
                                        'text_pos': pos,
                                        'para_text': para_text,
                                        'in_table': True,
                                        'table_idx': table_idx,
                                        'row_idx': row_idx,
                                        'cell_idx': cell_idx
                                    })
                                
                                start = pos + 1
        
        # Determine which occurrences to replace
        if occurrence_index is not None:
            if occurrence_index < 1 or occurrence_index > len(occurrences):
                return f"Invalid occurrence_index: {occurrence_index}. Found {len(occurrences)} occurrence(s). Valid range: 1-{len(occurrences)}."
            # Replace only the specified occurrence
            occurrences_to_replace = [occurrences[occurrence_index - 1]]
        else:
            # Replace all occurrences
            occurrences_to_replace = occurrences
        
        # Perform replacements
        count = 0
        for occ in occurrences_to_replace:
            para = occ['para']
            text_pos = occ['text_pos']
            para_text = occ['para_text']
            
            # Find the run containing this position
            current_pos = 0
            for run in para.runs:
                run_end = current_pos + len(run.text)
                if current_pos <= text_pos < run_end:
                    # This run contains the text
                    run_start_in_text = text_pos - current_pos
                    if find_text in run.text[run_start_in_text:run_start_in_text + len(find_text)]:
                        # Replace in this run
                        run_text = run.text
                        # Replace the first occurrence starting at run_start_in_text
                        new_run_text = run_text[:run_start_in_text] + replace_text + run_text[run_start_in_text + len(find_text):]
                        run.text = new_run_text
                        count += 1
                        break
                current_pos = run_end
        
        if count > 0:
            doc.save(filename)
            return f"Replaced {count} occurrence(s) of '{find_text}' with '{replace_text}'."
        else:
            return f"No occurrences of '{find_text}' found."
    except Exception as e:
        return f"Failed to search and replace: {str(e)}"

async def insert_header_near_text_tool(filename: str, target_text: str = None, header_title: str = "", position: str = 'after', header_style: str = 'Heading 1', target_paragraph_index: int = None,
                                        track_changes: bool = False, change_author: str = "") -> str:
    """Insert a header (with specified style) before or after the target paragraph. Specify by text or paragraph index.
    
    Args:
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
    """
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        # For Track Changes, use python-docx approach but note that document_utils functions don't support it
        # Fall back to regular implementation with a note
        return insert_header_near_text(filename, target_text, header_title, position, header_style, target_paragraph_index) + " (Note: Track Changes not fully supported for this operation)"
    return insert_header_near_text(filename, target_text, header_title, position, header_style, target_paragraph_index)

async def insert_numbered_list_near_text_tool(filename: str, target_text: str = None, list_items: list = None, position: str = 'after', target_paragraph_index: int = None, bullet_type: str = 'bullet',
                                              track_changes: bool = False, change_author: str = "") -> str:
    """Insert a bulleted or numbered list before or after the target paragraph. Specify by text or paragraph index.
    
    Args:
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
    """
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        # For Track Changes, use python-docx approach but note that document_utils functions don't support it
        return insert_numbered_list_near_text(filename, target_text, list_items, position, target_paragraph_index, bullet_type) + " (Note: Track Changes not fully supported for this operation)"
    return insert_numbered_list_near_text(filename, target_text, list_items, position, target_paragraph_index, bullet_type)

async def insert_line_or_paragraph_near_text_tool(filename: str, target_text: str = None, line_text: str = "", position: str = 'after', line_style: str = None, target_paragraph_index: int = None,
                                                   target_occurrence: int = 1, track_changes: bool = False, change_author: str = "") -> str:
    """Insert a new line or paragraph (with specified or matched style) before or after the target paragraph. Specify by text or paragraph index.
    
    Args:
        filename: Path to the Word document
        target_text: Text to search for to locate the target paragraph (optional)
        line_text: Text content for the new line or paragraph
        position: Position relative to target ('before' or 'after', default: 'after')
        line_style: Style name for the new line (optional, will match target style if not provided)
        target_paragraph_index: Index of the target paragraph (optional, alternative to target_text)
        target_occurrence: When using target_text, select this occurrence (1-based) if there are multiple matches
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Check if file is writeable
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}. Consider creating a copy first."
    
    # Check if Track Changes is requested
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        
        try:
            word, doc, saved_state = open_document_with_track_changes(filename, change_author)
            
            # Убеждаемся, что Track Changes включен
            if not doc.TrackRevisions:
                doc.TrackRevisions = True
            
            # Устанавливаем автора изменений
            if change_author:
                word.UserName = change_author
            
            # Validate occurrence
            if target_text and target_occurrence is not None and target_occurrence < 1:
                close_document_with_track_changes(word, doc, saved_state, save_changes=False)
                return "target_occurrence must be a positive integer when using target_text."
            
            # Find target paragraph
            target_para = None
            anchor_index = None
            
            if target_paragraph_index is not None:
                # IMPORTANT: Indices in COM API may differ from python-docx due to tables
                # Solution: Get paragraph text from python-docx first, then find it in COM API by text
                try:
                    from docx import Document as DocxDocument
                    docx_doc = DocxDocument(filename)
                    if target_paragraph_index < 0 or target_paragraph_index >= len(docx_doc.paragraphs):
                        close_document_with_track_changes(word, doc, saved_state, save_changes=False)
                        return f"Invalid target_paragraph_index: {target_paragraph_index}. Document has {len(docx_doc.paragraphs)} paragraphs (0-{len(docx_doc.paragraphs)-1})."
                    
                    # Get the target paragraph text from python-docx
                    target_para_docx = docx_doc.paragraphs[target_paragraph_index]
                    target_para_text = target_para_docx.text.strip()
                    
                    if not target_para_text:
                        # If paragraph is empty, try to find by position using Range
                        # This is a fallback for empty paragraphs
                        # Count paragraphs in COM API, skipping TOC and table paragraphs
                        para_count = 0
                        for i in range(1, doc.Paragraphs.Count + 1):
                            para = doc.Paragraphs(i)
                            try:
                                style_name = para.Style.NameLocal if para.Style else ""
                                if style_name and style_name.lower().startswith("toc"):
                                    continue
                                # Check if paragraph is in a table
                                if para.Range.Information(12):  # wdWithInTable = 12
                                    continue
                                if para_count == target_paragraph_index:
                                    target_para = para
                                    anchor_index = target_paragraph_index
                                    break
                                para_count += 1
                            except:
                                para_count += 1
                    else:
                        # Find paragraph by text in COM API, skipping TOC paragraphs
                        for i in range(1, doc.Paragraphs.Count + 1):
                            para = doc.Paragraphs(i)
                            # Skip TOC paragraphs
                            try:
                                style_name = para.Style.NameLocal if para.Style else ""
                                if style_name and style_name.lower().startswith("toc"):
                                    continue
                                # Skip paragraphs inside tables (they have different indexing)
                                if para.Range.Information(12):  # wdWithInTable = 12
                                    continue
                            except:
                                pass
                            
                            para_text = para.Range.Text.strip()
                            # Match by exact text or by containing the text
                            if para_text == target_para_text or target_para_text in para_text:
                                target_para = para
                                anchor_index = target_paragraph_index
                                break
                    
                    if target_para is None:
                        close_document_with_track_changes(word, doc, saved_state, save_changes=False)
                        return f"Target paragraph at index {target_paragraph_index} not found in COM API. Text: '{target_para_text[:50]}...'"
                except Exception as e:
                    close_document_with_track_changes(word, doc, saved_state, save_changes=False)
                    return f"Failed to find target paragraph by index: {str(e)}"
            else:
                # Find by text
                if not target_text:
                    close_document_with_track_changes(word, doc, saved_state, save_changes=False)
                    return "Either target_text or target_paragraph_index must be provided."
                
                match_count = 0
                # Search through paragraphs, skipping TOC and table paragraphs
                for i in range(1, doc.Paragraphs.Count + 1):
                    para = doc.Paragraphs(i)
                    # Skip TOC paragraphs
                    try:
                        style_name = para.Style.NameLocal if para.Style else ""
                        if style_name and style_name.lower().startswith("toc"):
                            continue
                        # Skip paragraphs inside tables
                        if para.Range.Information(12):  # wdWithInTable = 12
                            continue
                    except:
                        pass
                    para_text = para.Range.Text
                    if target_text in para_text:
                        match_count += 1
                        if match_count == target_occurrence:
                            target_para = para
                            # Try to find the index in python-docx for reporting
                            try:
                                from docx import Document as DocxDocument
                                docx_doc = DocxDocument(filename)
                                match_count_docx = 0
                                for idx, p in enumerate(docx_doc.paragraphs):
                                    if target_text in p.text:
                                        match_count_docx += 1
                                        if match_count_docx == target_occurrence:
                                            anchor_index = idx
                                            break
                            except:
                                anchor_index = None
                            break
                
                if target_para is None:
                    close_document_with_track_changes(word, doc, saved_state, save_changes=False)
                    occurrence_msg = ""
                    if match_count > 0:
                        occurrence_msg = f" Only {match_count} occurrence(s) found but target_occurrence={target_occurrence}."
                    return f"Target paragraph not found (by text '{target_text}'). (TOC paragraphs and table paragraphs are skipped in text search){occurrence_msg}"
            
            # Determine style: use provided or match target
            style_name = None
            if line_style:
                style_name = line_style
            else:
                # Match target paragraph style
                try:
                    style_name = target_para.Style.NameLocal if target_para.Style else "Normal"
                except:
                    style_name = "Normal"
            
            # Insert paragraph via COM API
            target_range = target_para.Range
            
            if position == 'before':
                # Insert before: create range at start of target paragraph
                insert_range = doc.Range(target_range.Start, target_range.Start)
                # Insert text with paragraph mark - this creates a new paragraph before
                insert_range.InsertBefore(line_text + "\r")
                # Collapse range to get the inserted paragraph
                insert_range.Collapse(1)  # wdCollapseStart = 1
                # Get the inserted paragraph by finding paragraph at this position
                inserted_para = insert_range.Paragraphs(1)
            else:
                # Insert after: use end of target paragraph range
                # InsertAfter adds text after the range and creates a new paragraph
                insert_range = doc.Range(target_range.End - 1, target_range.End - 1)
                # Insert text with paragraph mark - this creates a new paragraph after
                insert_range.InsertAfter(line_text + "\r")
                # Move range to the inserted paragraph
                insert_range.Collapse(0)  # wdCollapseEnd = 0
                # Get the inserted paragraph
                inserted_para = insert_range.Paragraphs(1)
            
            # Apply style to inserted paragraph
            try:
                inserted_para.Style = style_name
            except:
                # Style doesn't exist, try to set to Normal
                try:
                    inserted_para.Style = "Normal"
                except:
                    pass  # Continue without style if it fails
            
            # Save and close
            close_document_with_track_changes(word, doc, saved_state)
            
            if anchor_index is not None:
                return f"Line/paragraph inserted {position} paragraph (index {anchor_index}) with style '{style_name}' with Track Changes."
            else:
                return f"Line/paragraph inserted {position} the target paragraph with style '{style_name}' with Track Changes."
        except Exception as e:
            return f"Failed to insert line/paragraph with Track Changes: {str(e)}"
    
    # Fall back to python-docx approach if track_changes=False
    return insert_line_or_paragraph_near_text(filename, target_text, line_text, position, line_style, target_paragraph_index, target_occurrence)

async def replace_paragraph_block_below_header_tool(filename: str, header_text: str, new_paragraphs: list, detect_block_end_fn=None,
                                                    track_changes: bool = False, change_author: str = "") -> str:
    """Reemplaza el bloque de párrafos debajo de un encabezado, evitando modificar TOC.
    
    Args:
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
    """
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        # For Track Changes, use python-docx approach but note that document_utils functions don't support it
        return replace_paragraph_block_below_header(filename, header_text, new_paragraphs, detect_block_end_fn) + " (Note: Track Changes not fully supported for this operation)"
    return replace_paragraph_block_below_header(filename, header_text, new_paragraphs, detect_block_end_fn)

async def replace_block_between_manual_anchors_tool(filename: str, start_anchor_text: str, new_paragraphs: list, end_anchor_text: str = None, match_fn=None, new_paragraph_style: str = None,
                                                    track_changes: bool = False, change_author: str = "") -> str:
    """Replace all content between start_anchor_text and end_anchor_text (or next logical header if not provided).
    
    Args:
        track_changes: If True, changes will be tracked as revisions
        change_author: Author name for tracked changes (required if track_changes=True)
    """
    if track_changes:
        if not check_pywin32_available():
            return get_track_changes_error_message()
        # For Track Changes, use python-docx approach but note that document_utils functions don't support it
        return replace_block_between_manual_anchors(filename, start_anchor_text, new_paragraphs, end_anchor_text, match_fn, new_paragraph_style) + " (Note: Track Changes not fully supported for this operation)"
    return replace_block_between_manual_anchors(filename, start_anchor_text, new_paragraphs, end_anchor_text, match_fn, new_paragraph_style)
