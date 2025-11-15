"""
Utilities for working with Track Changes (режим правки) in Word documents via COM API.

This module provides functions to enable/disable Track Changes and manage
change tracking state in Word documents using Windows COM API.
"""
import os
from typing import Optional, Callable, Any, Tuple


def enable_track_changes(doc, author: str = "", initials: str = "") -> None:
    """
    Enable Track Changes in a Word document via COM API.
    
    Args:
        doc: Word Document object from COM API (win32com.client)
        author: Author name for tracked changes (optional)
        initials: Author initials for tracked changes (optional)
    """
    try:
        # Enable track revisions
        doc.TrackRevisions = True
        
        # Set author if provided
        if author:
            doc.Application.UserName = author
        
        # Set initials if provided
        if initials:
            doc.Application.UserInitials = initials
    except Exception as e:
        raise Exception(f"Failed to enable Track Changes: {str(e)}")


def disable_track_changes(doc) -> None:
    """
    Disable Track Changes in a Word document via COM API.
    
    Args:
        doc: Word Document object from COM API (win32com.client)
    """
    try:
        doc.TrackRevisions = False
    except Exception as e:
        raise Exception(f"Failed to disable Track Changes: {str(e)}")


def get_track_changes_status(doc) -> bool:
    """
    Get the current Track Changes status in a Word document.
    
    Args:
        doc: Word Document object from COM API (win32com.client)
        
    Returns:
        True if Track Changes is enabled, False otherwise
    """
    try:
        return bool(doc.TrackRevisions)
    except Exception:
        return False


def save_track_changes_state(doc) -> Tuple[bool, str, str]:
    """
    Save the current Track Changes state before modifying it.
    
    Args:
        doc: Word Document object from COM API (win32com.client)
        
    Returns:
        Tuple of (was_enabled, original_author, original_initials)
    """
    try:
        was_enabled = get_track_changes_status(doc)
        original_author = doc.Application.UserName if hasattr(doc.Application, 'UserName') else ""
        original_initials = doc.Application.UserInitials if hasattr(doc.Application, 'UserInitials') else ""
        return (was_enabled, original_author, original_initials)
    except Exception:
        return (False, "", "")


def restore_track_changes_state(doc, was_enabled: bool, original_author: str = "", original_initials: str = "") -> None:
    """
    Restore the Track Changes state after modification.
    
    Args:
        doc: Word Document object from COM API (win32com.client)
        was_enabled: Whether Track Changes was enabled before
        original_author: Original author name to restore
        original_initials: Original initials to restore
    """
    try:
        doc.TrackRevisions = was_enabled
        if original_author:
            doc.Application.UserName = original_author
        if original_initials:
            doc.Application.UserInitials = original_initials
    except Exception:
        pass  # Silently fail if restoration fails


def open_document_with_track_changes(filename: str, author: str = "", initials: str = "", 
                                     preserve_existing_state: bool = False) -> Tuple[Any, Any, Tuple[bool, str, str]]:
    """
    Open a Word document via COM API and enable Track Changes.
    
    Args:
        filename: Path to the Word document
        author: Author name for tracked changes
        initials: Author initials for tracked changes
        preserve_existing_state: If True, preserve existing Track Changes state and restore it later
        
    Returns:
        Tuple of (word_application, document, saved_state)
        saved_state is None if preserve_existing_state=False, otherwise tuple of (was_enabled, original_author, original_initials)
    """
    try:
        import win32com.client
    except ImportError:
        raise ImportError("pywin32 is required for Track Changes support. Install it with: pip install pywin32")
    
    abs_filename = os.path.abspath(filename)
    
    if not os.path.exists(abs_filename):
        raise FileNotFoundError(f"Document {abs_filename} does not exist")
    
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    try:
        # Устанавливаем автора ПЕРЕД открытием документа
        if author:
            word.UserName = author
        if initials:
            word.UserInitials = initials
        
        doc = word.Documents.Open(abs_filename)
        
        saved_state = None
        if preserve_existing_state:
            saved_state = save_track_changes_state(doc)
        
        # Enable Track Changes - ВАЖНО: включаем ПОСЛЕ открытия документа
        enable_track_changes(doc, author, initials)
        
        # Убеждаемся, что Track Changes действительно включен
        if not doc.TrackRevisions:
            doc.TrackRevisions = True
        
        return (word, doc, saved_state)
    except Exception as e:
        try:
            word.Quit()
        except:
            pass
        raise Exception(f"Failed to open document: {str(e)}")


def close_document_with_track_changes(word_app, doc, saved_state: Optional[Tuple[bool, str, str]] = None, 
                                     save_changes: bool = True) -> None:
    """
    Close a Word document and restore Track Changes state if needed.
    
    Args:
        word_app: Word Application object from COM API
        doc: Word Document object from COM API
        saved_state: Saved state tuple from open_document_with_track_changes (optional)
        save_changes: Whether to save changes before closing
    """
    try:
        # Сохраняем документ ПЕРЕД восстановлением состояния, чтобы правки сохранились
        if save_changes:
            doc.Save()
        
        # Восстанавливаем состояние только если оно было сохранено
        if saved_state:
            restore_track_changes_state(doc, saved_state[0], saved_state[1], saved_state[2])
            # Если мы сохранили изменения, нужно сохранить еще раз после восстановления состояния
            if save_changes:
                doc.Save()
        else:
            # Если состояние не сохранялось, просто закрываем документ
            if not save_changes:
                doc.Close(SaveChanges=False)
                word_app.Quit()
                return
        
        doc.Close()
        word_app.Quit()
    except Exception as e:
        try:
            doc.Close(SaveChanges=False)
        except:
            pass
        try:
            word_app.Quit()
        except:
            pass
        raise Exception(f"Failed to close document: {str(e)}")


def check_pywin32_available() -> bool:
    """
    Check if pywin32 is available for COM API access.
    
    Returns:
        True if pywin32 is available, False otherwise
    """
    try:
        import win32com.client
        return True
    except ImportError:
        return False


def get_track_changes_error_message() -> str:
    """
    Get error message for when pywin32 is not available.
    
    Returns:
        Error message string
    """
    return "Track Changes requires pywin32. Install it with: pip install pywin32. Note: Track Changes only works on Windows."

