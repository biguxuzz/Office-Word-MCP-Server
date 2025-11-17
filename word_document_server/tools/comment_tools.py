"""
Comment extraction tools for Word Document Server.

These tools provide high-level interfaces for extracting and analyzing
comments from Word documents through the MCP protocol.
"""
import os
import json
from typing import Dict, List, Optional, Any
from docx import Document

from word_document_server.utils.file_utils import ensure_docx_extension
from word_document_server.core.comments import (
    extract_all_comments,
    filter_comments_by_author,
    get_comments_for_paragraph,
    add_reply_to_comment
)


async def get_all_comments(filename: str) -> str:
    """
    Extract all comments from a Word document with paragraph index information.
    
    Args:
        filename: Path to the Word document
        
    Returns:
        JSON string containing all comments with metadata including paragraph_index
    
    Note:
        COM API enrichment (for paragraph_index) is disabled on Windows by default
        to prevent hanging. Set environment variable ENABLE_WORD_COM_API=1 to enable.
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return json.dumps({
            'success': False,
            'error': f'Document {filename} does not exist'
        }, indent=2)
    
    try:
        # Load the document
        doc = Document(filename)
        
        # Extract all comments
        comments = extract_all_comments(doc)
        
        # Try to enrich comments with paragraph_index using COM API if available
        # Skip on Windows with asyncio to avoid hangs (unless explicitly enabled)
        import platform
        enable_com_api = os.getenv('ENABLE_WORD_COM_API', '0') == '1'
        skip_com_api = platform.system() == 'Windows' and not enable_com_api
        
        try:
            if skip_com_api:
                # Skip COM API on Windows to avoid hangs with asyncio
                raise ImportError("COM API disabled on Windows for stability")
            
            import win32com.client
            import pythoncom
            
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            abs_filename = os.path.abspath(filename)
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            try:
                com_doc = word.Documents.Open(abs_filename, ReadOnly=True, AddToRecentFiles=False)
                
                # First, get all comments from COM API and find their paragraph indices
                com_comments_data = []
                for i in range(1, com_doc.Comments.Count + 1):
                    try:
                        com_comment = com_doc.Comments(i)
                        com_author = com_comment.Author if hasattr(com_comment, 'Author') else ""
                        com_text = com_comment.Range.Text if hasattr(com_comment, 'Range') else ""
                        com_date = com_comment.Date if hasattr(com_comment, 'Date') else None
                        comment_range = com_comment.Scope if hasattr(com_comment, 'Scope') else None
                        
                        if comment_range:
                            # Find the paragraph that contains or is closest to this range
                            # Skip paragraphs in tables - we only want main document paragraphs
                            best_para_idx = None
                            min_distance_before = float('inf')
                            min_distance_after = float('inf')
                            best_para_before = None
                            best_para_after = None
                            
                            for para_idx in range(len(doc.paragraphs)):
                                try:
                                    para = com_doc.Paragraphs(para_idx + 1)
                                    # Skip paragraphs in tables
                                    if para.Range.Information(12):  # wdWithInTable = 12
                                        continue
                                    
                                    para_range_start = para.Range.Start
                                    para_range_end = para.Range.End
                                    
                                    # Check if comment range overlaps this paragraph
                                    if comment_range.Start >= para_range_start and comment_range.Start <= para_range_end:
                                        best_para_idx = para_idx
                                        break
                                    
                                    # If comment is before this paragraph
                                    if comment_range.Start < para_range_start:
                                        distance = para_range_start - comment_range.Start
                                        if distance < min_distance_before:
                                            min_distance_before = distance
                                            best_para_before = para_idx - 1 if para_idx > 0 else None
                                    
                                    # If comment is after this paragraph
                                    if comment_range.Start > para_range_end:
                                        distance = comment_range.Start - para_range_end
                                        if distance < min_distance_after:
                                            min_distance_after = distance
                                            best_para_after = para_idx
                                except:
                                    continue
                            
                            # Determine best paragraph index and position
                            insert_position = None  # 'inside', 'before', or 'after'
                            if best_para_idx is not None:
                                # Comment is inside this paragraph
                                final_para_idx = best_para_idx
                                insert_position = 'inside'
                            else:
                                # Comment is between paragraphs - find the closest one
                                # Prefer the paragraph AFTER the comment (for insertion before it)
                                # but if the paragraph BEFORE is closer, use that (for insertion after it)
                                if best_para_after is not None and best_para_before is not None:
                                    # Choose the closest one
                                    if min_distance_after <= min_distance_before:
                                        final_para_idx = best_para_after
                                        insert_position = 'before'  # Insert before this paragraph
                                    else:
                                        final_para_idx = best_para_before
                                        insert_position = 'after'  # Insert after this paragraph
                                elif best_para_after is not None:
                                    final_para_idx = best_para_after
                                    insert_position = 'before'
                                elif best_para_before is not None and best_para_before >= 0:
                                    final_para_idx = best_para_before
                                    insert_position = 'after'
                                else:
                                    final_para_idx = None
                                    insert_position = None
                            
                            com_comments_data.append({
                                'author': com_author,
                                'text': com_text.strip()[:100],  # First 100 chars for matching
                                'date': com_date,
                                'paragraph_index': final_para_idx,
                                'comment_range_start': comment_range.Start if comment_range else None,
                                'insert_position': insert_position
                            })
                    except:
                        continue
                
                # Now match python-docx comments with COM API comments
                # Use more flexible matching to handle format differences
                for comment in comments:
                    comment_date = comment.get('date')
                    comment_author = comment.get('author', '')
                    comment_text = comment.get('text', '').strip()
                    
                    # Try to find matching COM API comment
                    best_match = None
                    best_score = 0
                    
                    for com_data in com_comments_data:
                        score = 0
                        
                        # Match by author (flexible matching)
                        if comment_author and com_data['author']:
                            comment_author_lower = comment_author.lower()
                            com_author_lower = com_data['author'].lower()
                            # Check if names match (allowing for partial matches)
                            if comment_author_lower == com_author_lower:
                                score += 5
                            elif comment_author_lower in com_author_lower or com_author_lower in comment_author_lower:
                                score += 3
                            # Check by initials or last name
                            comment_words = set(comment_author_lower.split())
                            com_words = set(com_author_lower.split())
                            if comment_words.intersection(com_words):
                                score += 2
                        
                        # Match by date (more flexible)
                        if comment_date and com_data['date']:
                            try:
                                from datetime import datetime
                                # Normalize dates
                                comment_date_str = str(comment_date).replace('Z', '+00:00')
                                com_date_str = str(com_data['date'])
                                
                                # Try to parse both dates
                                try:
                                    comment_dt = datetime.fromisoformat(comment_date_str.replace('Z', '+00:00'))
                                    # COM API date might be in different format, try to parse
                                    try:
                                        # Try various date formats
                                        if isinstance(com_data['date'], (int, float)):
                                            # COM date might be serial number
                                            com_dt = datetime.fromtimestamp(com_data['date'])
                                        else:
                                            com_dt = datetime.fromisoformat(com_date_str.replace('Z', '+00:00'))
                                        
                                        # Check if dates are very close (within 1 hour)
                                        time_diff = abs((comment_dt - com_dt).total_seconds())
                                        if time_diff < 3600:  # Within 1 hour
                                            score += 10
                                        elif time_diff < 86400:  # Within 1 day
                                            score += 5
                                    except:
                                        # Fallback to string matching
                                        if comment_date_str[:19] in com_date_str or com_date_str[:19] in comment_date_str:
                                            score += 5
                                except:
                                    # Fallback to string matching
                                    if comment_date_str[:19] in com_date_str or com_date_str[:19] in comment_date_str:
                                        score += 5
                            except:
                                pass
                        
                        # Match by text (partial, more flexible)
                        if comment_text and com_data['text']:
                            comment_text_lower = comment_text.lower()
                            com_text_lower = com_data['text'].lower()
                            # Check for significant text overlap
                            if len(comment_text_lower) > 20 and len(com_text_lower) > 20:
                                # Check if significant portion matches
                                if comment_text_lower[:50] in com_text_lower or com_text_lower[:50] in comment_text_lower:
                                    score += 3
                                # Check for key words match
                                comment_words = set(comment_text_lower.split()[:10])
                                com_words = set(com_text_lower.split()[:10])
                                common_words = comment_words.intersection(com_words)
                                if len(common_words) >= 3:  # At least 3 common words
                                    score += 2
                        
                        if score > best_score:
                            best_score = score
                            best_match = com_data
                    
                    # If we found a good match (score >= 5), use its paragraph_index
                    if best_match and best_score >= 5 and best_match['paragraph_index'] is not None:
                        comment['paragraph_index'] = best_match['paragraph_index']
                        comment['comment_range_start'] = best_match['comment_range_start']
                        # Set insertion position based on where the comment is located
                        insert_pos = best_match.get('insert_position', 'inside')
                        if insert_pos == 'before':
                            comment['insert_before'] = True
                        elif insert_pos == 'after':
                            comment['insert_after'] = True
                        elif insert_pos == 'inside':
                            # Comment is inside paragraph, default to insert after
                            comment['insert_after'] = True
                
                com_doc.Close(SaveChanges=False)
                word.Quit()
            except Exception as e:
                # Ensure Word is closed even if there's an error
                try:
                    if 'com_doc' in locals():
                        com_doc.Close(SaveChanges=False)
                except:
                    pass
                try:
                    if 'word' in locals():
                        word.Quit()
                except:
                    pass
            finally:
                # Uninitialize COM
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass
        except ImportError:
            # pywin32 not available or COM API disabled, use python-docx only
            pass
        except Exception as e:
            # COM API failed, continue with python-docx data
            pass
        
        # Return results
        return json.dumps({
            'success': True,
            'comments': comments,
            'total_comments': len(comments)
        }, indent=2)
        
    except Exception as e:
        return json.dumps({
            'success': False,
            'error': f'Failed to extract comments: {str(e)}'
        }, indent=2)


async def get_comments_by_author(filename: str, author: str) -> str:
    """
    Extract comments from a specific author in a Word document.
    
    Args:
        filename: Path to the Word document
        author: Name of the comment author to filter by
        
    Returns:
        JSON string containing filtered comments
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return json.dumps({
            'success': False,
            'error': f'Document {filename} does not exist'
        }, indent=2)
    
    if not author or not author.strip():
        return json.dumps({
            'success': False,
            'error': 'Author name cannot be empty'
        }, indent=2)
    
    try:
        # Load the document
        doc = Document(filename)
        
        # Extract all comments
        all_comments = extract_all_comments(doc)
        
        # Filter by author
        author_comments = filter_comments_by_author(all_comments, author)
        
        # Return results
        return json.dumps({
            'success': True,
            'author': author,
            'comments': author_comments,
            'total_comments': len(author_comments)
        }, indent=2)
        
    except Exception as e:
        return json.dumps({
            'success': False,
            'error': f'Failed to extract comments: {str(e)}'
        }, indent=2)


async def get_comments_for_paragraph(filename: str, paragraph_index: int) -> str:
    """
    Extract comments for a specific paragraph in a Word document.
    
    Args:
        filename: Path to the Word document
        paragraph_index: Index of the paragraph (0-based)
        
    Returns:
        JSON string containing comments for the specified paragraph
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return json.dumps({
            'success': False,
            'error': f'Document {filename} does not exist'
        }, indent=2)
    
    if paragraph_index < 0:
        return json.dumps({
            'success': False,
            'error': 'Paragraph index must be non-negative'
        }, indent=2)
    
    try:
        # Load the document
        doc = Document(filename)
        
        # Check if paragraph index is valid
        if paragraph_index >= len(doc.paragraphs):
            return json.dumps({
                'success': False,
                'error': f'Paragraph index {paragraph_index} is out of range. Document has {len(doc.paragraphs)} paragraphs.'
            }, indent=2)
        
        # Extract all comments
        all_comments = extract_all_comments(doc)
        
        # Filter for the specific paragraph
        from word_document_server.core.comments import get_comments_for_paragraph as core_get_comments_for_paragraph
        para_comments = core_get_comments_for_paragraph(all_comments, paragraph_index)
        
        # Get the paragraph text for context
        paragraph_text = doc.paragraphs[paragraph_index].text
        
        # Return results
        return json.dumps({
            'success': True,
            'paragraph_index': paragraph_index,
            'paragraph_text': paragraph_text,
            'comments': para_comments,
            'total_comments': len(para_comments)
        }, indent=2)
        
    except Exception as e:
        return json.dumps({
            'success': False,
            'error': f'Failed to extract comments: {str(e)}'
        }, indent=2)


async def reply_to_comment(filename: str, comment_id: str, reply_text: str, author: str = "", initials: str = "") -> str:
    """
    Add a reply to an existing comment in a Word document.
    
    Args:
        filename: Path to the Word document
        comment_id: Identifier of the comment to reply to. Can be:
            - Timestamp (ISO format, e.g., "2025-11-06T21:57:00+00:00") - recommended for reliable identification
            - Comment ID from comment data (e.g., "comment_1", "comment_2")
            - Numeric string (e.g., "0", "1") - less reliable as indices may shift
        reply_text: Text content for the reply
        author: Author name for the reply (optional)
        initials: Author initials for the reply (optional)
        
    Returns:
        JSON string indicating success or failure
    """
    filename = ensure_docx_extension(filename)
    
    if not os.path.exists(filename):
        return json.dumps({
            'success': False,
            'error': f'Document {filename} does not exist'
        }, indent=2)
    
    if not comment_id or not str(comment_id).strip():
        return json.dumps({
            'success': False,
            'error': 'Comment ID cannot be empty'
        }, indent=2)
    
    if not reply_text or not reply_text.strip():
        return json.dumps({
            'success': False,
            'error': 'Reply text cannot be empty'
        }, indent=2)
    
    try:
        # Add reply to the comment using COM interface
        # This creates a proper reply that appears as a separate comment in the thread
        success = add_reply_to_comment(filename, comment_id, reply_text, author, initials)
        
        if success:
            return json.dumps({
                'success': True,
                'message': f'Reply added to comment {comment_id}',
                'comment_id': comment_id,
                'reply_text': reply_text,
                'author': author if author else 'Current user'
            }, indent=2)
        else:
            return json.dumps({
                'success': False,
                'error': f'Comment with ID {comment_id} not found or could not be accessed. Make sure pywin32 is installed.'
            }, indent=2)
        
    except Exception as e:
        return json.dumps({
            'success': False,
            'error': f'Failed to add reply to comment: {str(e)}'
        }, indent=2)