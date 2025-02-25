from flask import Flask
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
import sys

# Import the positions module
from positions import POSITIONS, TABLE_COLUMNS, SPACING, DEBUG

# Path to your template PNG file
TEMPLATE_IMAGE_PATH = "ir-raw.png"

# Create a minimal Flask app for logging
app = Flask(__name__)

def generate_pdf(data):
    """Generate PDF with form data placed at predefined positions based on PHP format"""
    try:
        # Create Req_Forms directory if it doesn't exist
        output_dir = "Req_Forms"
        os.makedirs(output_dir, exist_ok=True)
        
        # Create a unique filename based on timestamp
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        pdf_filename = os.path.join(output_dir, f"test_order_{timestamp}.pdf")
        
        c = canvas.Canvas(pdf_filename, pagesize=letter)
        
        # Ensure template image exists
        if not os.path.exists(TEMPLATE_IMAGE_PATH):
            raise FileNotFoundError(f"Template image '{TEMPLATE_IMAGE_PATH}' not found")
            
        # Add background image with absolute path
        template_path = os.path.abspath(TEMPLATE_IMAGE_PATH)
        c.drawImage(template_path, 0, 0, width=letter[0], height=letter[1])
        
        # Set default font
        c.setFont("Helvetica", SPACING["font_size"])
        
        # PDF dimensions and coordinate system conversion
        width, height = letter
        
        # Scale factors to map from LaTeX coordinates to PDF points
        x_scale = width / 15  # 15 units in LaTeX width
        y_scale = height / 14  # 14 units in LaTeX height
        
        # LaTeX coordinates origin is at bottom left, PDF at top left
        # So we need to convert y-coordinates
        def y_pos(latex_y):
            return height - (latex_y * y_scale)
        
        # Extract form data with appropriate defaults
        vendor_name = data.get("vendor_name", "")
        address = data.get("vendor_address", "")
        tel = data.get("tel", "")
        fax = data.get("fax", "")
        ptao = data.get("ptao", "")
        date = data.get("date", datetime.now().strftime("%B %d, %Y"))
        notes = data.get("notes", "")
        req = data.get("purchaser", "")
        approver = "Dustin Keller"  # Hardcoded approver name
        
        # Position text elements using scaled coordinates from positions.py
        
        # Vendor name
        vendor_name_x, vendor_name_y = POSITIONS["vendor_name"]
        c.drawString(vendor_name_x * x_scale, y_pos(vendor_name_y), vendor_name)
        
        # Supplier/vendor address
        vendor_x, vendor_y = POSITIONS["vendor_address"]
        y_position = y_pos(vendor_y)
        for line in address.split("\n"):
            c.drawString(vendor_x * x_scale, y_position, line)
            y_position -= SPACING["address_line"]
        
        # Phone
        phone_x, phone_y = POSITIONS["phone"]
        c.drawString(phone_x * x_scale, y_pos(phone_y), tel)
        
        # Fax
        fax_x, fax_y = POSITIONS["fax"]
        c.drawString(fax_x * x_scale, y_pos(fax_y), fax)
        
        
        # Date
        date_x, date_y = POSITIONS["date"]
        c.drawString(date_x * x_scale, y_pos(date_y), date)
        
        # PTAO
        ptao_x, ptao_y = POSITIONS["ptao"]
        c.drawString(ptao_x * x_scale, y_pos(ptao_y), ptao)
        
        # Items table
        items_x, items_y = POSITIONS["items_start"]
        y_start = y_pos(items_y)
        
        # Define column positions based on TABLE_COLUMNS
        x_quantity = TABLE_COLUMNS["quantity"] * x_scale
        x_catalog = TABLE_COLUMNS["catalog"] * x_scale
        x_desc = TABLE_COLUMNS["description"] * x_scale
        x_unit_price = TABLE_COLUMNS["unit_price"] * x_scale
        x_total_price = TABLE_COLUMNS["total_price"] * x_scale
        
        # Process items
        items = data.get("items", [])
        row_height = SPACING["table_row"]
        total_amount = 0
        
        for i, item in enumerate(items):
            if i < 9:  # Limit to 9 items as in PHP
                quantity = float(item.get("quantity", 0))
                catalog_num = item.get("catalog_number", "")
                description = item.get("name", "")
                unit_price = float(item.get("price", 0))
                
                # Calculate total for this item
                total_price = quantity * unit_price
                total_amount += total_price
                
                # Format as currency
                formatted_unit_price = f"${unit_price:.2f}"
                formatted_total_price = f"${total_price:.2f}"
                
                # Position for this row
                row_y = y_start - (i * row_height)
                
                # Draw item row
                c.drawString(x_quantity, row_y, str(quantity))
                c.drawString(x_catalog, row_y, catalog_num)
                c.drawString(x_desc, row_y, description)
                c.drawString(x_unit_price, row_y, formatted_unit_price)
                c.drawString(x_total_price, row_y, formatted_total_price)
        
        # Total amount
        total_x, total_y = POSITIONS["total"]
        c.drawString(total_x * x_scale, y_pos(total_y), f"${total_amount:.2f}")
        
        # Notes
        notes_x, notes_y = POSITIONS["notes"]
        if notes:
            c.drawString(notes_x * x_scale, y_pos(notes_y), f"Notes: {notes}")
        
        # Requested by (purchaser name)
        req_x, req_y = POSITIONS["requestor"]
        c.drawString(req_x * x_scale, y_pos(req_y), req)
        
        # Approved by
        approver_x, approver_y = POSITIONS["approver"]
        c.drawString(approver_x * x_scale, y_pos(approver_y), approver)
        
        c.save()
        print(f"PDF generated successfully: {pdf_filename}")
        return pdf_filename
    except Exception as e:
        print(f"Error generating PDF: {str(e)}")
        raise

def create_test_data():
    """Create sample data for testing the PDF generation"""
    return {
        "vendor_name": "ABC Supplies Inc.",
        "vendor_address": "123 Main Street\nCharlottesville, VA 22903",
        "tel": "434-555-1234",
        "fax": "434-555-5678",
        "purpose": "Lab Equipment - PHYS 2010",
        "date": datetime.now().strftime("%B %d, %Y"),
        "ptao": "GRBLAHBLAHBLAH",
        "purchaser": "John Doe",
        "notes": "Please deliver to Physics Building Room 123",
        "items": [
            {
                "name": "Digital Oscilloscope",
                "quantity": 1,
                "price": 1299.99,
                "catalog_number": "OSC-2022"
            },
            {
                "name": "Function Generator",
                "quantity": 2,
                "price": 799.50,
                "catalog_number": "FG-100"
            },
            {
                "name": "Multimeter",
                "quantity": 5,
                "price": 129.95,
                "catalog_number": "MM-500"
            }
        ]
    }

if __name__ == "__main__":
    print("Generating test PDF...")
    test_data = create_test_data()
    
    # Allow command line arguments to modify test data
    if len(sys.argv) > 1:
        vendor_name = sys.argv[1]
        test_data["vendor_name"] = vendor_name
    
    pdf_path = generate_pdf(test_data)
    print(f"Test PDF generated at: {pdf_path}")
    print("Done!") 