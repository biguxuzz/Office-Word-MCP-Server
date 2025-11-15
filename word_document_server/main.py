"""
Main entry point for the Word Document MCP Server.
Acts as the central controller for the MCP server that handles Word document operations.
Supports multiple transports: stdio, sse, and streamable-http using standalone FastMCP.
"""

import os
import sys
from dotenv import load_dotenv

# Load environment variables from .env file
print("Loading configuration from .env file...")
load_dotenv()
# Set required environment variable for FastMCP 2.8.1+
os.environ.setdefault('FASTMCP_LOG_LEVEL', 'INFO')
from fastmcp import FastMCP
from word_document_server.tools import (
    document_tools,
    content_tools,
    format_tools,
    protection_tools,
    footnote_tools,
    extended_document_tools,
    comment_tools
)
from word_document_server.tools.content_tools import replace_paragraph_block_below_header_tool
from word_document_server.tools.content_tools import replace_block_between_manual_anchors_tool

def get_transport_config():
    """
    Get transport configuration from environment variables.
    
    Returns:
        dict: Transport configuration with type, host, port, and other settings
    """
    # Default configuration
    config = {
        'transport': 'stdio',  # Default to stdio for backward compatibility
        'host': '0.0.0.0',
        'port': 8000,
        'path': '/mcp',
        'sse_path': '/sse'
    }
    
    # Override with environment variables if provided
    transport = os.getenv('MCP_TRANSPORT', 'stdio').lower()
    print(f"Transport: {transport}")
    # Validate transport type
    valid_transports = ['stdio', 'streamable-http', 'sse']
    if transport not in valid_transports:
        print(f"Warning: Invalid transport '{transport}'. Falling back to 'stdio'.")
        transport = 'stdio'
    
    config['transport'] = transport
    config['host'] = os.getenv('MCP_HOST', config['host'])
    # Use PORT from Render if available, otherwise fall back to MCP_PORT or default
    config['port'] = int(os.getenv('PORT', os.getenv('MCP_PORT', config['port'])))
    config['path'] = os.getenv('MCP_PATH', config['path'])
    config['sse_path'] = os.getenv('MCP_SSE_PATH', config['sse_path'])
    
    return config


def setup_logging(debug_mode):
    """
    Setup logging based on debug mode.
    
    Args:
        debug_mode (bool): Whether to enable debug logging
    """
    import logging
    
    if debug_mode:
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        print("Debug logging enabled")
    else:
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )


# Initialize FastMCP server
mcp = FastMCP("Word Document Server")


def register_tools():
    """Register all tools with the MCP server using FastMCP decorators."""
    
    # Document tools (create, copy, info, etc.)
    @mcp.tool()
    def create_document(filename: str, title: str = None, author: str = None):
        """Create a new Word document with optional metadata.
        
        Args:
            filename: Path where the new Word document will be created
            title: Optional document title
            author: Optional document author name
            
        Returns:
            Status message indicating success or failure
        """
        return document_tools.create_document(filename, title, author)
    
    @mcp.tool()
    def copy_document(source_filename: str, destination_filename: str = None):
        """Create a copy of a Word document.
        
        Args:
            source_filename: Path to the source Word document
            destination_filename: Path for the copied document (optional, will auto-generate if not provided)
            
        Returns:
            Status message indicating success or failure
        """
        return document_tools.copy_document(source_filename, destination_filename)
    
    @mcp.tool()
    def get_document_info(filename: str):
        """Get information about a Word document.
        
        Args:
            filename: Path to the Word document
            
        Returns:
            Document information including title, author, creation date, and other metadata
        """
        return document_tools.get_document_info(filename)
    
    @mcp.tool()
    def get_document_text(filename: str):
        """Extract all text from a Word document.
        
        Args:
            filename: Path to the Word document
            
        Returns:
            All text content from the document
        """
        return document_tools.get_document_text(filename)
    
    @mcp.tool()
    def get_document_outline(filename: str):
        """Get the structure of a Word document.
        
        Args:
            filename: Path to the Word document
            
        Returns:
            Document outline showing headings and structure
        """
        return document_tools.get_document_outline(filename)
    
    @mcp.tool()
    def list_available_documents(directory: str = "."):
        """List all .docx files in the specified directory.
        
        Args:
            directory: Directory path to search for Word documents (default: current directory)
            
        Returns:
            List of available .docx files in the directory
        """
        return document_tools.list_available_documents(directory)
    
    @mcp.tool()
    def get_document_xml(filename: str):
        """Get the raw XML structure of a Word document.
        
        Args:
            filename: Path to the Word document
            
        Returns:
            Raw XML structure of the document
        """
        return document_tools.get_document_xml_tool(filename)
    
    @mcp.tool()
    def insert_header_near_text(filename: str, target_text: str = None, header_title: str = None, position: str = 'after', header_style: str = 'Heading 1', target_paragraph_index: int = None,
                                track_changes: bool = False, change_author: str = ""):
        """Insert a header (with specified style) before or after the target paragraph.
        
        Args:
            filename: Path to the Word document
            target_text: Text to search for to locate the target paragraph (optional)
            header_title: Text content for the header
            position: Position relative to target ('before' or 'after', default: 'after')
            header_style: Style name for the header (default: 'Heading 1')
            target_paragraph_index: Index of the target paragraph (optional, alternative to target_text)
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return content_tools.insert_header_near_text_tool(filename, target_text, header_title, position, header_style, target_paragraph_index, track_changes, change_author)
    
    @mcp.tool()
    def insert_line_or_paragraph_near_text(filename: str, target_text: str = None, line_text: str = None, position: str = 'after', line_style: str = None, target_paragraph_index: int = None,
                                           track_changes: bool = False, change_author: str = ""):
        """Insert a new line or paragraph (with specified or matched style) before or after the target paragraph.
        
        Args:
            filename: Path to the Word document
            target_text: Text to search for to locate the target paragraph (optional)
            line_text: Text content for the new line or paragraph
            position: Position relative to target ('before' or 'after', default: 'after')
            line_style: Style name for the new line (optional, will match target style if not provided)
            target_paragraph_index: Index of the target paragraph (optional, alternative to target_text)
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return content_tools.insert_line_or_paragraph_near_text_tool(filename, target_text, line_text, position, line_style, target_paragraph_index, track_changes, change_author)
    
    @mcp.tool()
    def insert_numbered_list_near_text(filename: str, target_text: str = None, list_items: list = None, position: str = 'after', target_paragraph_index: int = None, bullet_type: str = 'bullet',
                                       track_changes: bool = False, change_author: str = ""):
        """Insert a bulleted or numbered list before or after the target paragraph.
        
        Args:
            filename: Path to the Word document
            target_text: Text to search for to locate the target paragraph (optional)
            list_items: List of strings for each list item
            position: Position relative to target ('before' or 'after', default: 'after')
            target_paragraph_index: Index of the target paragraph (optional, alternative to target_text)
            bullet_type: Type of list ('bullet' for bullets or 'number' for numbered lists, default: 'bullet')
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return content_tools.insert_numbered_list_near_text_tool(filename, target_text, list_items, position, target_paragraph_index, bullet_type, track_changes, change_author)
    # Content tools (paragraphs, headings, tables, etc.)
    @mcp.tool()
    def add_paragraph(filename: str, text: str, style: str = None,
                      font_name: str = None, font_size: int = None,
                      bold: bool = None, italic: bool = None, color: str = None,
                      track_changes: bool = False, change_author: str = ""):
        """Add a paragraph to a Word document with optional formatting.

        Args:
            filename: Path to Word document
            text: Paragraph text content
            style: Optional paragraph style name
            font_name: Font family (e.g., 'Helvetica', 'Times New Roman')
            font_size: Font size in points (e.g., 14, 36)
            bold: Make text bold
            italic: Make text italic
            color: Text color as hex RGB (e.g., '000000')
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
        """
        return content_tools.add_paragraph(filename, text, style, font_name, font_size, bold, italic, color, track_changes, change_author)
    
    @mcp.tool()
    def add_heading(filename: str, text: str, level: int = 1,
                    font_name: str = None, font_size: int = None,
                    bold: bool = None, italic: bool = None, border_bottom: bool = False,
                    track_changes: bool = False, change_author: str = ""):
        """Add a heading to a Word document with optional formatting.

        Args:
            filename: Path to Word document
            text: Heading text
            level: Heading level (1-9)
            font_name: Font family (e.g., 'Helvetica')
            font_size: Font size in points (e.g., 14)
            bold: Make heading bold
            italic: Make heading italic
            border_bottom: Add bottom border (for section headers)
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
        """
        return content_tools.add_heading(filename, text, level, font_name, font_size, bold, italic, border_bottom, track_changes, change_author)
    
    @mcp.tool()
    def add_picture(filename: str, image_path: str, width: float = None,
                    track_changes: bool = False, change_author: str = ""):
        """Add an image to a Word document.
        
        Args:
            filename: Path to the Word document
            image_path: Path to the image file to insert
            width: Optional width for the image in points
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return content_tools.add_picture(filename, image_path, width, track_changes, change_author)
    
    @mcp.tool()
    def add_table(filename: str, rows: int, cols: int, data: list = None,
                  track_changes: bool = False, change_author: str = ""):
        """Add a table to a Word document.
        
        Args:
            filename: Path to the Word document
            rows: Number of rows in the table
            cols: Number of columns in the table
            data: Optional list of lists containing table data (rows x cols)
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return content_tools.add_table(filename, rows, cols, data, track_changes, change_author)
    
    @mcp.tool()
    def add_page_break(filename: str, track_changes: bool = False, change_author: str = ""):
        """Add a page break to the document.
        
        Args:
            filename: Path to the Word document
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return content_tools.add_page_break(filename, track_changes, change_author)
    
    @mcp.tool()
    def delete_paragraph(filename: str, paragraph_index: int,
                         track_changes: bool = False, change_author: str = ""):
        """Delete a paragraph from a document.
        
        Args:
            filename: Path to the Word document
            paragraph_index: Index of the paragraph to delete (0-based)
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return content_tools.delete_paragraph(filename, paragraph_index, track_changes, change_author)
    
    @mcp.tool()
    def search_and_replace(filename: str, find_text: str, replace_text: str,
                          track_changes: bool = False, change_author: str = "",
                          occurrence_index: int = None, context_before: str = "", context_after: str = "",
                          replace_all: bool = True):
        """Search for text and replace all occurrences or a specific occurrence.
        
        Args:
            filename: Path to the Word document
            find_text: Text to search for
            replace_text: Text to replace with
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            occurrence_index: Optional index of specific occurrence to replace (1-based, e.g., 1 for first, 2 for second).
                             If None, replaces all occurrences (or first occurrence if replace_all=False).
            context_before: Optional text that should appear before the match (for precise identification)
            context_after: Optional text that should appear after the match (for precise identification)
            replace_all: If True, replace all occurrences. If False, replace only first occurrence (deprecated, use occurrence_index=1 instead).
                        Ignored if occurrence_index is specified.
            
        Returns:
            Status message indicating success or failure, including number of replacements made
        """
        # Handle replace_all for backward compatibility
        # If replace_all=False and occurrence_index is not specified, replace only first occurrence
        if not replace_all and occurrence_index is None:
            occurrence_index = 1
        
        return content_tools.search_and_replace(filename, find_text, replace_text, track_changes, change_author,
                                                occurrence_index, context_before, context_after)
    
    # Format tools (styling, text formatting, etc.)
    @mcp.tool()
    def create_custom_style(filename: str, style_name: str, bold: bool = None, 
                          italic: bool = None, font_size: int = None, 
                          font_name: str = None, color: str = None, 
                          base_style: str = None,
                          track_changes: bool = False, change_author: str = ""):
        """Create a custom style in the document.
        
        Args:
            filename: Path to the Word document
            style_name: Name for the new custom style
            bold: Make text bold (optional)
            italic: Make text italic (optional)
            font_size: Font size in points (optional)
            font_name: Font family name (optional)
            color: Text color as hex RGB (optional, e.g., '000000')
            base_style: Base style to inherit from (optional)
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.create_custom_style(
            filename, style_name, bold, italic, font_size, font_name, color, base_style, track_changes, change_author
        )
    
    @mcp.tool()
    def format_text(filename: str, paragraph_index: int, start_pos: int, end_pos: int,
                   bold: bool = None, italic: bool = None, underline: bool = None,
                   color: str = None, font_size: int = None, font_name: str = None,
                   track_changes: bool = False, change_author: str = ""):
        """Format a specific range of text within a paragraph.
        
        Args:
            filename: Path to the Word document
            paragraph_index: Index of the paragraph (0-based)
            start_pos: Starting character position within the paragraph
            end_pos: Ending character position within the paragraph
            bold: Make text bold (optional)
            italic: Make text italic (optional)
            underline: Add underline (optional)
            color: Text color as hex RGB (optional)
            font_size: Font size in points (optional)
            font_name: Font family name (optional)
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.format_text(
            filename, paragraph_index, start_pos, end_pos, bold, italic, 
            underline, color, font_size, font_name, track_changes, change_author
        )
    
    @mcp.tool()
    def format_table(filename: str, table_index: int, has_header_row: bool = None,
                    border_style: str = None, shading: list = None,
                    track_changes: bool = False, change_author: str = ""):
        """Format a table with borders, shading, and structure.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table to format (0-based)
            has_header_row: Whether the table has a header row (optional)
            border_style: Border style name (optional)
            shading: List of shading colors for rows (optional)
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.format_table(filename, table_index, has_header_row, border_style, shading, track_changes, change_author)
    
    # New table cell shading tools
    @mcp.tool()
    def set_table_cell_shading(filename: str, table_index: int, row_index: int, 
                              col_index: int, fill_color: str, pattern: str = "clear",
                              track_changes: bool = False, change_author: str = ""):
        """Apply shading/filling to a specific table cell.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            row_index: Row index of the cell (0-based)
            col_index: Column index of the cell (0-based)
            fill_color: Fill color as hex RGB (e.g., 'FFFFFF')
            pattern: Pattern type (default: 'clear')
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.set_table_cell_shading(filename, table_index, row_index, col_index, fill_color, pattern, track_changes, change_author)
    
    @mcp.tool()
    def apply_table_alternating_rows(filename: str, table_index: int, 
                                   color1: str = "FFFFFF", color2: str = "F2F2F2",
                                   track_changes: bool = False, change_author: str = ""):
        """Apply alternating row colors to a table for better readability.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            color1: Color for odd rows as hex RGB (default: 'FFFFFF')
            color2: Color for even rows as hex RGB (default: 'F2F2F2')
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.apply_table_alternating_rows(filename, table_index, color1, color2, track_changes, change_author)
    
    @mcp.tool()
    def highlight_table_header(filename: str, table_index: int, 
                             header_color: str = "4472C4", text_color: str = "FFFFFF",
                             track_changes: bool = False, change_author: str = ""):
        """Apply special highlighting to table header row.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            header_color: Background color for header row as hex RGB (default: '4472C4')
            text_color: Text color for header row as hex RGB (default: 'FFFFFF')
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.highlight_table_header(filename, table_index, header_color, text_color, track_changes, change_author)
    
    # Cell merging tools
    @mcp.tool()
    def merge_table_cells(filename: str, table_index: int, start_row: int, start_col: int, 
                        end_row: int, end_col: int,
                        track_changes: bool = False, change_author: str = ""):
        """Merge cells in a rectangular area of a table.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            start_row: Starting row index (0-based)
            start_col: Starting column index (0-based)
            end_row: Ending row index (0-based)
            end_col: Ending column index (0-based)
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.merge_table_cells(filename, table_index, start_row, start_col, end_row, end_col, track_changes, change_author)
    
    @mcp.tool()
    def merge_table_cells_horizontal(filename: str, table_index: int, row_index: int, 
                                   start_col: int, end_col: int,
                                   track_changes: bool = False, change_author: str = ""):
        """Merge cells horizontally in a single row.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            row_index: Row index to merge cells in (0-based)
            start_col: Starting column index (0-based)
            end_col: Ending column index (0-based)
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.merge_table_cells_horizontal(filename, table_index, row_index, start_col, end_col, track_changes, change_author)
    
    @mcp.tool()
    def merge_table_cells_vertical(filename: str, table_index: int, col_index: int, 
                                 start_row: int, end_row: int,
                                 track_changes: bool = False, change_author: str = ""):
        """Merge cells vertically in a single column.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            col_index: Column index to merge cells in (0-based)
            start_row: Starting row index (0-based)
            end_row: Ending row index (0-based)
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.merge_table_cells_vertical(filename, table_index, col_index, start_row, end_row, track_changes, change_author)
    
    # Cell alignment tools
    @mcp.tool()
    def set_table_cell_alignment(filename: str, table_index: int, row_index: int, col_index: int,
                               horizontal: str = "left", vertical: str = "top",
                               track_changes: bool = False, change_author: str = ""):
        """Set text alignment for a specific table cell.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            row_index: Row index of the cell (0-based)
            col_index: Column index of the cell (0-based)
            horizontal: Horizontal alignment ('left', 'center', 'right', default: 'left')
            vertical: Vertical alignment ('top', 'center', 'bottom', default: 'top')
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.set_table_cell_alignment(filename, table_index, row_index, col_index, horizontal, vertical, track_changes, change_author)
    
    @mcp.tool()
    def set_table_alignment_all(filename: str, table_index: int, 
                              horizontal: str = "left", vertical: str = "top",
                              track_changes: bool = False, change_author: str = ""):
        """Set text alignment for all cells in a table.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            horizontal: Horizontal alignment for all cells ('left', 'center', 'right', default: 'left')
            vertical: Vertical alignment for all cells ('top', 'center', 'bottom', default: 'top')
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.set_table_alignment_all(filename, table_index, horizontal, vertical, track_changes, change_author)
    
    # Protection tools
    @mcp.tool()
    def protect_document(filename: str, password: str):
        """Add password protection to a Word document.
        
        Args:
            filename: Path to the Word document
            password: Password to protect the document with
            
        Returns:
            Status message indicating success or failure
        """
        return protection_tools.protect_document(filename, password)
    
    @mcp.tool()
    def unprotect_document(filename: str, password: str):
        """Remove password protection from a Word document.
        
        Args:
            filename: Path to the Word document
            password: Password to unlock the document
            
        Returns:
            Status message indicating success or failure
        """
        return protection_tools.unprotect_document(filename, password)
    
    # Footnote tools
    @mcp.tool()
    def add_footnote_to_document(filename: str, paragraph_index: int, footnote_text: str,
                                 track_changes: bool = False, change_author: str = ""):
        """Add a footnote to a specific paragraph in a Word document.
        
        Args:
            filename: Path to the Word document
            paragraph_index: Index of the paragraph to add footnote to (0-based)
            footnote_text: Text content for the footnote
            
        Returns:
            Status message indicating success or failure
        """
        return footnote_tools.add_footnote_to_document(filename, paragraph_index, footnote_text, track_changes, change_author)
    
    @mcp.tool()
    def add_footnote_after_text(filename: str, search_text: str, footnote_text: str, 
                               output_filename: str = None,
                               track_changes: bool = False, change_author: str = ""):
        """Add a footnote after specific text with proper superscript formatting.
        
        Args:
            filename: Path to the Word document
            search_text: Text to search for to locate insertion point
            footnote_text: Text content for the footnote
            output_filename: Optional output filename (if not provided, modifies original)
            
        Returns:
            Status message indicating success or failure
        """
        return footnote_tools.add_footnote_after_text(filename, search_text, footnote_text, output_filename, track_changes, change_author)
    
    @mcp.tool()
    def add_footnote_before_text(filename: str, search_text: str, footnote_text: str, 
                                output_filename: str = None,
                                track_changes: bool = False, change_author: str = ""):
        """Add a footnote before specific text with proper superscript formatting.
        
        Args:
            filename: Path to the Word document
            search_text: Text to search for to locate insertion point
            footnote_text: Text content for the footnote
            output_filename: Optional output filename (if not provided, modifies original)
            
        Returns:
            Status message indicating success or failure
        """
        return footnote_tools.add_footnote_before_text(filename, search_text, footnote_text, output_filename, track_changes, change_author)
    
    @mcp.tool()
    def add_footnote_enhanced(filename: str, paragraph_index: int, footnote_text: str,
                             output_filename: str = None,
                             track_changes: bool = False, change_author: str = ""):
        """Enhanced footnote addition with guaranteed superscript formatting.
        
        Args:
            filename: Path to the Word document
            paragraph_index: Index of the paragraph to add footnote to (0-based)
            footnote_text: Text content for the footnote
            output_filename: Optional output filename (if not provided, modifies original)
            
        Returns:
            Status message indicating success or failure
        """
        return footnote_tools.add_footnote_enhanced(filename, paragraph_index, footnote_text, output_filename, track_changes, change_author)
    
    @mcp.tool()
    def add_endnote_to_document(filename: str, paragraph_index: int, endnote_text: str,
                                 track_changes: bool = False, change_author: str = ""):
        """Add an endnote to a specific paragraph in a Word document.
        
        Args:
            filename: Path to the Word document
            paragraph_index: Index of the paragraph to add endnote to (0-based)
            endnote_text: Text content for the endnote
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return footnote_tools.add_endnote_to_document(filename, paragraph_index, endnote_text, track_changes, change_author)
    
    @mcp.tool()
    def customize_footnote_style(filename: str, numbering_format: str = "1, 2, 3",
                                start_number: int = 1, font_name: str = None,
                                font_size: int = None,
                                track_changes: bool = False, change_author: str = ""):
        """Customize footnote numbering and formatting in a Word document.
        
        Args:
            filename: Path to the Word document
            numbering_format: Numbering format (default: '1, 2, 3')
            start_number: Starting number for footnotes (default: 1)
            font_name: Font name for footnotes (optional)
            font_size: Font size for footnotes (optional)
            track_changes: If True, changes will be tracked as revisions
            change_author: Author name for tracked changes (required if track_changes=True)
            
        Returns:
            Status message indicating success or failure
        """
        return footnote_tools.customize_footnote_style(
            filename, numbering_format, start_number, font_name, font_size, track_changes, change_author
        )
    
    @mcp.tool()
    def delete_footnote_from_document(filename: str, footnote_id: int = None,
                                     search_text: str = None, output_filename: str = None,
                                     track_changes: bool = False, change_author: str = ""):
        """Delete a footnote from a Word document.
        
        Args:
            filename: Path to the Word document
            footnote_id: ID of the footnote to delete (1, 2, 3, etc., optional)
            search_text: Text near the footnote to locate it (optional, alternative to footnote_id)
            output_filename: Optional output filename (if not provided, modifies original)
            
        Returns:
            Status message indicating success or failure
        """
        return footnote_tools.delete_footnote_from_document(
            filename, footnote_id, search_text, output_filename, track_changes, change_author
        )
    
    # Robust footnote tools - Production-ready with comprehensive validation
    @mcp.tool()
    def add_footnote_robust(filename: str, search_text: str = None, 
                           paragraph_index: int = None, footnote_text: str = "",
                           validate_location: bool = True, auto_repair: bool = False,
                           track_changes: bool = False, change_author: str = ""):
        """Add footnote with robust validation and Word compliance.
        
        Args:
            filename: Path to the Word document
            search_text: Text to search for to locate insertion point (optional)
            paragraph_index: Index of the paragraph (optional, alternative to search_text)
            footnote_text: Text content for the footnote
            validate_location: Whether to validate insertion location (default: True)
            auto_repair: Whether to automatically repair document issues (default: False)
            
        Returns:
            Status message indicating success or failure
        """
        return footnote_tools.add_footnote_robust_tool(
            filename, search_text, paragraph_index, footnote_text, 
            validate_location, auto_repair, track_changes, change_author
        )
    
    @mcp.tool()
    def validate_document_footnotes(filename: str):
        """Validate all footnotes in document for coherence and compliance.
        
        Args:
            filename: Path to the Word document
            
        Returns:
            Detailed report on ID conflicts, orphaned content, missing styles, etc.
        """
        return footnote_tools.validate_footnotes_tool(filename)
    
    @mcp.tool()
    def delete_footnote_robust(filename: str, footnote_id: int = None,
                              search_text: str = None, clean_orphans: bool = True):
        """Delete footnote with comprehensive cleanup and orphan removal.
        
        Args:
            filename: Path to the Word document
            footnote_id: ID of the footnote to delete (optional)
            search_text: Text near the footnote to locate it (optional, alternative to footnote_id)
            clean_orphans: Whether to clean up orphaned content (default: True)
            
        Returns:
            Status message indicating success or failure
        """
        return footnote_tools.delete_footnote_robust_tool(
            filename, footnote_id, search_text, clean_orphans
        )
    
    # Extended document tools
    @mcp.tool()
    def get_paragraph_text_from_document(filename: str, paragraph_index: int):
        """Get text from a specific paragraph in a Word document.
        
        Args:
            filename: Path to the Word document
            paragraph_index: Index of the paragraph (0-based)
            
        Returns:
            Text content of the specified paragraph
        """
        return extended_document_tools.get_paragraph_text_from_document(filename, paragraph_index)
    
    @mcp.tool()
    def find_text_in_document(filename: str, text_to_find: str, match_case: bool = True,
                             whole_word: bool = False):
        """Find occurrences of specific text in a Word document.
        
        Args:
            filename: Path to the Word document
            text_to_find: Text to search for
            match_case: Whether to match case (default: True)
            whole_word: Whether to match whole words only (default: False)
            
        Returns:
            List of occurrences with their locations
        """
        return extended_document_tools.find_text_in_document(
            filename, text_to_find, match_case, whole_word
        )
    
    @mcp.tool()
    def convert_to_pdf(filename: str, output_filename: str = None):
        """Convert a Word document to PDF format.
        
        Args:
            filename: Path to the Word document
            output_filename: Optional output PDF filename (if not provided, uses same name with .pdf extension)
            
        Returns:
            Status message indicating success or failure
        """
        return extended_document_tools.convert_to_pdf(filename, output_filename)

    @mcp.tool()
    def replace_paragraph_block_below_header(filename: str, header_text: str, new_paragraphs: list, detect_block_end_fn=None):
        """Replace the paragraph block below a header, avoiding modification of TOC.
        
        Args:
            filename: Path to the Word document
            header_text: Text of the header to find
            new_paragraphs: List of new paragraph texts to replace the block
            detect_block_end_fn: Optional function to detect block end (optional)
            
        Returns:
            Status message indicating success or failure
        """
        return replace_paragraph_block_below_header_tool(filename, header_text, new_paragraphs, detect_block_end_fn)

    @mcp.tool()
    def replace_block_between_manual_anchors(filename: str, start_anchor_text: str, new_paragraphs: list, end_anchor_text: str = None, match_fn=None, new_paragraph_style: str = None):
        """Replace all content between start_anchor_text and end_anchor_text (or next logical header if not provided).
        
        Args:
            filename: Path to the Word document
            start_anchor_text: Text marking the start of the block to replace
            new_paragraphs: List of new paragraph texts to replace the block
            end_anchor_text: Text marking the end of the block (optional, uses next header if not provided)
            match_fn: Optional matching function (optional)
            new_paragraph_style: Style to apply to new paragraphs (optional)
            
        Returns:
            Status message indicating success or failure
        """
        return replace_block_between_manual_anchors_tool(filename, start_anchor_text, new_paragraphs, end_anchor_text, match_fn, new_paragraph_style)

    # Comment tools
    @mcp.tool()
    def get_all_comments(filename: str):
        """Extract all comments from a Word document.
        
        Args:
            filename: Path to the Word document
            
        Returns:
            List of all comments with their details. Each comment includes:
            - id: Sequential ID (e.g., "comment_1", "comment_2")
            - comment_id: XML comment ID
            - author: Comment author name
            - date: Timestamp in ISO format (e.g., "2025-11-06T21:57:00+00:00") - use this for reliable identification in reply_to_comment
            - text: Comment text content
            - Other metadata fields
        """
        return comment_tools.get_all_comments(filename)
    
    @mcp.tool()
    def get_comments_by_author(filename: str, author: str):
        """Extract comments from a specific author in a Word document.
        
        Args:
            filename: Path to the Word document
            author: Name of the author to filter comments by
            
        Returns:
            List of comments from the specified author
        """
        return comment_tools.get_comments_by_author(filename, author)
    
    @mcp.tool()
    def get_comments_for_paragraph(filename: str, paragraph_index: int):
        """Extract comments for a specific paragraph in a Word document.
        
        Args:
            filename: Path to the Word document
            paragraph_index: Index of the paragraph (0-based)
            
        Returns:
            List of comments associated with the specified paragraph
        """
        return comment_tools.get_comments_for_paragraph(filename, paragraph_index)
    
    @mcp.tool()
    def reply_to_comment(filename: str, comment_id: str, reply_text: str, author: str = "", initials: str = ""):
        """Add a reply to an existing comment in a Word document.
        
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
            Status message indicating success or failure
        """
        return comment_tools.reply_to_comment(filename, comment_id, reply_text, author, initials)
    # New table column width tools
    @mcp.tool()
    def set_table_column_width(filename: str, table_index: int, col_index: int, 
                              width: float, width_type: str = "points",
                              track_changes: bool = False, change_author: str = ""):
        """Set the width of a specific table column.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            col_index: Index of the column (0-based)
            width: Width value
            width_type: Unit type for width ('points', 'inches', 'cm', default: 'points')
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.set_table_column_width(filename, table_index, col_index, width, width_type, track_changes, change_author)

    @mcp.tool()
    def set_table_column_widths(filename: str, table_index: int, widths: list, 
                               width_type: str = "points",
                               track_changes: bool = False, change_author: str = ""):
        """Set the widths of multiple table columns.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            widths: List of width values (one per column)
            width_type: Unit type for widths ('points', 'inches', 'cm', default: 'points')
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.set_table_column_widths(filename, table_index, widths, width_type, track_changes, change_author)

    @mcp.tool()
    def set_table_width(filename: str, table_index: int, width: float, 
                       width_type: str = "points",
                       track_changes: bool = False, change_author: str = ""):
        """Set the overall width of a table.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            width: Overall table width
            width_type: Unit type for width ('points', 'inches', 'cm', default: 'points')
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.set_table_width(filename, table_index, width, width_type, track_changes, change_author)

    @mcp.tool()
    def auto_fit_table_columns(filename: str, table_index: int,
                               track_changes: bool = False, change_author: str = ""):
        """Set table columns to auto-fit based on content.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.auto_fit_table_columns(filename, table_index, track_changes, change_author)

    # New table cell text formatting and padding tools
    @mcp.tool()
    def format_table_cell_text(filename: str, table_index: int, row_index: int, col_index: int,
                               text_content: str = None, bold: bool = None, italic: bool = None,
                               underline: bool = None, color: str = None, font_size: int = None,
                               font_name: str = None,
                               track_changes: bool = False, change_author: str = ""):
        """Format text within a specific table cell.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            row_index: Row index of the cell (0-based)
            col_index: Column index of the cell (0-based)
            text_content: Text to set in the cell (optional)
            bold: Make text bold (optional)
            italic: Make text italic (optional)
            underline: Add underline (optional)
            color: Text color as hex RGB (optional)
            font_size: Font size in points (optional)
            font_name: Font family name (optional)
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.format_table_cell_text(filename, table_index, row_index, col_index,
                                                   text_content, bold, italic, underline, color, font_size, font_name,
                                                   track_changes, change_author)

    @mcp.tool()
    def set_table_cell_padding(filename: str, table_index: int, row_index: int, col_index: int,
                               top: float = None, bottom: float = None, left: float = None, 
                               right: float = None, unit: str = "points",
                               track_changes: bool = False, change_author: str = ""):
        """Set padding/margins for a specific table cell.
        
        Args:
            filename: Path to the Word document
            table_index: Index of the table (0-based)
            row_index: Row index of the cell (0-based)
            col_index: Column index of the cell (0-based)
            top: Top padding (optional)
            bottom: Bottom padding (optional)
            left: Left padding (optional)
            right: Right padding (optional)
            unit: Unit type for padding ('points', 'inches', 'cm', default: 'points')
            
        Returns:
            Status message indicating success or failure
        """
        return format_tools.set_table_cell_padding(filename, table_index, row_index, col_index,
                                                   top, bottom, left, right, unit, track_changes, change_author)



def run_server():
    """Run the Word Document MCP Server with configurable transport."""
    # Get transport configuration
    config = get_transport_config()
    
    # Setup logging
    # setup_logging(config['debug'])
    
    # Register all tools
    register_tools()
    
    # Print startup information
    transport_type = config['transport']
    print(f"Starting Word Document MCP Server with {transport_type} transport...")
    
    # if config['debug']:
    #     print(f"Configuration: {config}")
    
    try:
        if transport_type == 'stdio':
            # Run with stdio transport (default, backward compatible)
            print("Server running on stdio transport")
            mcp.run(transport='stdio')
            
        elif transport_type == 'streamable-http':
            # Run with streamable HTTP transport
            print(f"Server running on streamable-http transport at http://{config['host']}:{config['port']}{config['path']}")
            mcp.run(
                transport='streamable-http',
                host=config['host'],
                port=config['port'],
                path=config['path']
            )
            
        elif transport_type == 'sse':
            # Run with SSE transport
            print(f"Server running on SSE transport at http://{config['host']}:{config['port']}{config['sse_path']}")
            mcp.run(
                transport='sse',
                host=config['host'],
                port=config['port'],
                path=config['sse_path']
            )
            
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"Error starting server: {e}")
        if config['debug']:
            import traceback
            traceback.print_exc()
        sys.exit(1)
    
    return mcp


def main():
    """Main entry point for the server."""
    run_server()


if __name__ == "__main__":
    main()
