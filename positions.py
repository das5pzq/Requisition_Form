"""
Coordinate positions for placing text on the requisition form.
Based on the original PHP/LaTeX coordinate system.

The form uses a coordinate system where:
- x coordinate ranges from 0 to ~15 units (left to right)
- y coordinate ranges from 0 to ~14 units (bottom to top)

These positions are used by both the PDF generation and PNG overlay functions.
"""

# Main element positions
POSITIONS = {
    # Vendor information
    "vendor_name": (2.5, 1.5),      # Position for vendor name
    # "vendor": (2.5, 1.5),           # Adding this alias for "vendor_name"
    "vendor_address": (2.5, 2.2),   # Position for vendor street address
    "vendor_city_state": (2.5, 2.6), # Position for vendor city/state/zip
    
    # Contact information
    "phone": (1.2, 3.25),
    "fax": (1.2, 3.8),
    
    # Order details
    "date": (10.5, 0.9),
    "ptao": (10.5, 1.4),
    
    # Table positions
    "table_start": (1.5, 5.5),
    "items_start": (1.5, 5.5),  # Alias for "table_start"
    "table_width": 12,
    
    # Signature fields
    "signature": (2.5, 11.5),
    "notes": (1.0, 11.0),
    
    # Footer information
    "total": (13.4, 10.30),         # Total amount
    "requestor": (3.25, 11.85),      # Requested by (purchaser name)
    "approver": (3.25, 12.45)       # Approved by
}

# Column positions for the items table
# These are x-coordinates for each column
TABLE_COLUMNS = {
    "quantity": 0.445,              # Quantity column
    "catalog": 0.445 + 1.55,        # Catalog number column
    "description": 0.445 + 1.55 + 1.85,  # Description column
    "unit_price": 0.445 + 1.55 + 1.85 + 7.65,  # Unit price column
    "total_price": 0.445 + 1.55 + 1.85 + 7.65 + 2.05  # Total price column
}

# Spacing configuration
SPACING = {
    "font_size": 10,
    "line_height": 14,
    "address_line": 20,  # Spacing between address lines
    "table_row": 20,     # Height of each table row
    "table_header_height": 25,  # Height of table header row
    "col_widths": [6, 1.5, 2, 2.5],  # Column widths for the items table
    "col_positions": [0, 6, 7.5, 9.5]  # Starting positions for each column
}

# Debugging options
DEBUG = {
    "draw_grid": True,              # Whether to draw coordinate grid in debug mode
    "draw_position_markers": True   # Whether to mark key positions in debug mode
} 