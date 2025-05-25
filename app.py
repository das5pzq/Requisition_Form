from flask import Flask, request, send_file, render_template, send_from_directory
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image
import cv2
import numpy as np
import yagmail
import requests
import os
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from positions import POSITIONS, TABLE_COLUMNS, SPACING, DEBUG
from google.oauth2 import service_account
from googleapiclient.discovery import build
import json
from googleapiclient.http import MediaFileUpload
import textwrap
from reportlab.pdfbase.pdfmetrics import stringWidth

app = Flask(__name__)

# Google Form URL (Change to your form URL)
GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeL97014pGTbbwJR-t5BktA3kjc_v0Af3t2NBhzcnEp5MLL7g/formResponse"

# Email credentials
EMAIL_SENDER = "xxdrp3pp3rxlov3rxx@gmail.com"
EMAIL_PASSWORD = "nyty jsgq btqd xfzu"
EMAIL_RECEIVER = "xxdrp3pp3rxlov3rxx@gmail.com"

# Path to your template PNG file
TEMPLATE_IMAGE_PATH = "ir-raw.png"

def truncate_text(text, max_chars, suffix="..."):
    """Truncate text to fit within character limit"""
    if len(text) <= max_chars:
        return text
    return text[:max_chars-len(suffix)] + suffix

def wrap_text_to_width(text, max_width, font_name, font_size):
    """Wrap text to fit within pixel width using reportlab's stringWidth"""
    if not text:
        return [""]
    
    words = text.split()
    lines = []
    current_line = []
    
    for word in words:
        # Test if adding this word exceeds width
        test_line = " ".join(current_line + [word])
        if stringWidth(test_line, font_name, font_size) <= max_width:
            current_line.append(word)
        else:
            # Start new line
            if current_line:
                lines.append(" ".join(current_line))
                current_line = [word]
            else:
                # Single word is too long, truncate it
                truncated_word = truncate_text(word, int(max_width / (font_size * 0.6)))
                lines.append(truncated_word)
    
    if current_line:
        lines.append(" ".join(current_line))
    
    return lines

def get_text_width_pil(text, font):
    """Calculate text width for PIL fonts with better error handling"""
    if not text:
        return 0
        
    try:
        # For PIL fonts, use textbbox to get width
        bbox = font.getbbox(text)
        return bbox[2] - bbox[0]  # width = right - left
    except AttributeError:
        # Fallback for older PIL versions
        try:
            return font.getsize(text)[0]
        except:
            # Final fallback: estimate based on character count
            return len(text) * 7
    except Exception as e:
        app.logger.warning(f"Error calculating text width: {e}")
        # Fallback: estimate based on character count and font size
        font_size = getattr(font, 'size', 12)
        return len(text) * (font_size * 0.6)

def smart_text_placement(draw_func, x, y, text, max_width, max_lines=1, font_name="Helvetica", font_size=10, pil_font=None):
    """Enhanced text placement with proper width calculation"""
    if not text:
        return y
    
    # Calculate text width based on context
    if pil_font:  # PNG context
        def get_width(test_text):
            return get_text_width_pil(test_text, pil_font)
    else:  # PDF context
        def get_width(test_text):
            return stringWidth(test_text, font_name, font_size)
    
    # Smart truncation
    if get_width(text) <= max_width:
        draw_func(x, y, text)
        return y
    
    # Text is too long, truncate with ellipsis
    ellipsis = "..."
    ellipsis_width = get_width(ellipsis)
    available_width = max_width - ellipsis_width
    
    # Binary search to find the longest text that fits
    left, right = 0, len(text)
    best_length = 0
    
    while left <= right:
        mid = (left + right) // 2
        test_text = text[:mid]
        
        if get_width(test_text) <= available_width:
            best_length = mid
            left = mid + 1
        else:
            right = mid - 1
    
    truncated_text = text[:best_length] + ellipsis
    draw_func(x, y, truncated_text)
    return y

def generate_pdf(data):
    """Generate PDF with enhanced text handling"""
    try:
        # Create Req_Forms directory if it doesn't exist
        output_dir = "Req_Forms"
        os.makedirs(output_dir, exist_ok=True)
        
        # Generate unique filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pdf_filename = os.path.join(output_dir, f"requisition_{timestamp}.pdf")
        c = canvas.Canvas(pdf_filename, pagesize=letter)
        
        # Ensure template image exists
        if not os.path.exists(TEMPLATE_IMAGE_PATH):
            raise FileNotFoundError(f"Template image '{TEMPLATE_IMAGE_PATH}' not found")
            
        # Add background image
        template_path = os.path.abspath(TEMPLATE_IMAGE_PATH)
        c.drawImage(template_path, 0, 0, width=letter[0], height=letter[1])
        
        # Set default font
        font_name = "Helvetica"
        font_size = SPACING["font_size"]
        c.setFont(font_name, font_size)
        
        # PDF dimensions and coordinate system conversion
        width, height = letter
        x_scale = width / 15
        y_scale = height / 14
        
        def y_pos(latex_y):
            return height - (latex_y * y_scale)
        
        def pdf_draw_text(x, y, text):
            """Helper function for PDF text drawing"""
            c.drawString(x, y, text)
        
        # Extract and format data with character limits
        vendor_name = data.get("vendor", "") or data.get("Vendor Name", "")
        ptao = data.get("purpose", "") or data.get("Purpose of Purchase", "")
        date = data.get("date", datetime.now().strftime("%m/%d/%Y"))
        req = data.get("name", "") or data.get("Name of Purchaser", "")
        
        # Field-specific width limits (in pixels)
        field_widths = {
            "vendor_name": 200,
            "ptao": 150,
            "requestor": 120,
            "description": 250,
            "catalog": 80
        }
        
        # Vendor name with smart text placement
        vendor_x, vendor_y = POSITIONS["vendor_name"]
        smart_text_placement(
            pdf_draw_text,
            vendor_x * x_scale, 
            y_pos(vendor_y),
            vendor_name,
            field_widths["vendor_name"],
            max_lines=1,
            font_name=font_name,
            font_size=font_size
        )
        
        # Date
        date_x, date_y = POSITIONS["date"]
        c.drawString(date_x * x_scale, y_pos(date_y), date)
        
        # PTAO with text wrapping
        ptao_x, ptao_y = POSITIONS["ptao"]
        smart_text_placement(
            pdf_draw_text,
            ptao_x * x_scale,
            y_pos(ptao_y),
            ptao,
            field_widths["ptao"],
            max_lines=1,
            font_name=font_name,
            font_size=font_size
        )
        
        # Process items with enhanced formatting
        items = data.get("items", [])
        row_height = SPACING["table_row"]
        total_amount = 0
        
        items_x, items_y = POSITIONS["items_start"]
        y_start = y_pos(items_y)
        
        for i, item in enumerate(items[:9]):  # Limit to 9 items
            try:
                quantity = float(item.get("quantity", 0))
                catalog_num = item.get("catalog_number", "")
                description = item.get("description", "") or item.get("name", "")
                unit_price = float(item.get("unit_price", 0)) or float(item.get("price", 0))
                
                total_price = quantity * unit_price
                total_amount += total_price
                
                row_y = y_start - (i * row_height)
                
                # Quantity (simple)
                c.drawString(TABLE_COLUMNS["quantity"] * x_scale, row_y, str(int(quantity)))
                
                # Catalog number with truncation
                smart_text_placement(
                    pdf_draw_text,
                    TABLE_COLUMNS["catalog"] * x_scale,
                    row_y,
                    catalog_num,
                    field_widths["catalog"],
                    max_lines=1,
                    font_name=font_name,
                    font_size=font_size
                )
                
                # Description with smart wrapping
                smart_text_placement(
                    pdf_draw_text,
                    TABLE_COLUMNS["description"] * x_scale,
                    row_y,
                    description,
                    field_widths["description"],
                    max_lines=1,
                    font_name=font_name,
                    font_size=font_size
                )
                
                # Prices
                c.drawString(TABLE_COLUMNS["unit_price"] * x_scale, row_y, f"${unit_price:.2f}")
                c.drawString(TABLE_COLUMNS["total_price"] * x_scale, row_y, f"${total_price:.2f}")
                
            except (ValueError, TypeError) as e:
                app.logger.warning(f"Error processing item {i}: {e}")
                continue
        
        # Total amount
        total_x, total_y = POSITIONS["total"]
        c.drawString(total_x * x_scale, y_pos(total_y), f"${total_amount:.2f}")
        
        # Requestor with truncation
        req_x, req_y = POSITIONS["requestor"]
        smart_text_placement(
            pdf_draw_text,
            req_x * x_scale,
            y_pos(req_y),
            req,
            field_widths["requestor"],
            max_lines=1,
            font_name=font_name,
            font_size=font_size
        )
        
        # Approver
        approver_x, approver_y = POSITIONS["approver"]
        c.drawString(approver_x * x_scale, y_pos(approver_y), "Dustin Keller")
        
        c.save()
        return pdf_filename
        
    except Exception as e:
        app.logger.error(f"Error generating PDF: {str(e)}")
        raise

# Enhanced overlay_on_png function
def load_font(font_size=12):
    """Load font with fallbacks for different systems"""
    font_paths = [
        "arial.ttf",
        "/System/Library/Fonts/Arial.ttf",  # macOS
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux
        "C:\\Windows\\Fonts\\arial.ttf",  # Windows
    ]
    
    for font_path in font_paths:
        try:
            return ImageFont.truetype(font_path, font_size)
        except (OSError, IOError):
            continue
    
    # Final fallback
    try:
        return ImageFont.load_default()
    except:
        app.logger.warning("Could not load any font, using basic fallback")
        return None

def overlay_on_png(data, debug=False):
    """Enhanced PNG overlay with proper text handling"""
    try:
        if not os.path.exists(TEMPLATE_IMAGE_PATH):
            raise FileNotFoundError(f"Template image '{TEMPLATE_IMAGE_PATH}' not found")
            
        img = Image.open(TEMPLATE_IMAGE_PATH)
        draw = ImageDraw.Draw(img)

        # Get image dimensions
        width, height = img.size
        
        # Scale factors
        x_scale = width / 15
        y_scale = height / 14
        
        # Load font with better fallback
        font = load_font(12)
        if font is None:
            app.logger.error("Could not load font for PNG generation")
            # You might want to raise an exception or use text without font
        
        def png_draw_text(x, y, text):
            """Helper function for PNG text drawing"""
            draw.text((x, y), text, font=font, fill="black")
        
        # Field-specific width limits (in pixels for PNG)
        field_widths = {
            "vendor_name": 300,
            "ptao": 200,
            "requestor": 150,
            "description": 350,
            "catalog": 100
        }
        
        # Extract data with correct keys
        vendor_name = data.get("vendor", "")
        ptao = data.get("purpose", "")
        date = data.get("date", datetime.now().strftime("%m/%d/%Y"))
        req = data.get("name", "")
        
        # BUG: Missing other fields that might be needed
        telephone = data.get("telephone", "")
        fax = data.get("fax", "")
        vendor_address = data.get("vendor_address", "")
        notes = data.get("notes", "")
        
        # Vendor name with smart placement
        vendor_name_x, vendor_name_y = POSITIONS["vendor_name"]
        smart_text_placement(
            png_draw_text,
            vendor_name_x * x_scale,
            vendor_name_y * y_scale,
            vendor_name,
            field_widths["vendor_name"],
            max_lines=1,
            font_name="arial",
            font_size=12,
            pil_font=font  # Pass PIL font for width calculation
        )
        
        # Date
        date_x, date_y = POSITIONS["date"]
        draw.text((date_x * x_scale, date_y * y_scale), date, font=font, fill="black")
        
        # PTAO
        ptao_x, ptao_y = POSITIONS["ptao"]
        smart_text_placement(
            png_draw_text,
            ptao_x * x_scale,
            ptao_y * y_scale,
            ptao,
            field_widths["ptao"],
            max_lines=1,
            pil_font=font
        )
        
        # Add missing fields to PNG if needed
        if telephone:
            phone_x, phone_y = POSITIONS["phone"]
            draw.text((phone_x * x_scale, phone_y * y_scale), telephone, font=font, fill="black")
        
        if fax:
            fax_x, fax_y = POSITIONS["fax"]
            draw.text((fax_x * x_scale, fax_y * y_scale), fax, font=font, fill="black")
        
        # Add vendor address with line breaks
        if vendor_address:
            vendor_addr_x, vendor_addr_y = POSITIONS["vendor_address"]
            y_position = vendor_addr_y * y_scale
            for line in vendor_address.split("\n")[:3]:  # Limit to 3 lines
                smart_text_placement(
                    png_draw_text,
                    vendor_addr_x * x_scale,
                    y_position,
                    line,
                    field_widths["vendor_name"],
                    pil_font=font
                )
                y_position += SPACING["address_line"]
        
        # Process items
        items = data.get("items", [])
        y_start = POSITIONS["items_start"][1] * y_scale
        total_amount = 0
        
        for i, item in enumerate(items[:9]):
            try:
                quantity = float(item.get("quantity", 0))
                catalog_num = item.get("catalog_number", "")
                description = item.get("description", "")
                unit_price = float(item.get("unit_price", 0))
                
                total_price = quantity * unit_price
                total_amount += total_price
                
                row_y = y_start + (i * SPACING["table_row"])
                
                # Draw with smart text placement
                draw.text((TABLE_COLUMNS["quantity"] * x_scale, row_y), str(int(quantity)), font=font, fill="black")
                
                smart_text_placement(
                    png_draw_text,
                    TABLE_COLUMNS["catalog"] * x_scale,
                    row_y,
                    catalog_num,
                    field_widths["catalog"],
                    pil_font=font
                )
                
                smart_text_placement(
                    png_draw_text,
                    TABLE_COLUMNS["description"] * x_scale,
                    row_y,
                    description,
                    field_widths["description"],
                    pil_font=font
                )
                
                draw.text((TABLE_COLUMNS["unit_price"] * x_scale, row_y), f"${unit_price:.2f}", font=font, fill="black")
                draw.text((TABLE_COLUMNS["total_price"] * x_scale, row_y), f"${total_price:.2f}", font=font, fill="black")
                
            except (ValueError, TypeError) as e:
                app.logger.warning(f"Error processing PNG item {i}: {e}")
                continue
        
        # Total and signatures
        draw.text((POSITIONS["total"][0] * x_scale, POSITIONS["total"][1] * y_scale), 
                  f"${total_amount:.2f}", font=font, fill="black")
        
        smart_text_placement(
            png_draw_text,
            POSITIONS["requestor"][0] * x_scale,
            POSITIONS["requestor"][1] * y_scale,
            req,
            field_widths["requestor"],
            pil_font=font
        )
        
        draw.text((POSITIONS["approver"][0] * x_scale, POSITIONS["approver"][1] * y_scale), 
                  "Dustin Keller", font=font, fill="black")
        
        # Save file
        suffix = "_debug" if debug else ""
        file_base = f"{req.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d%H%M%S')}{suffix}"
        output_path = os.path.join("Req_Forms", f"{file_base}.png")
        os.makedirs("Req_Forms", exist_ok=True)
        img.save(output_path)
        return output_path
        
    except Exception as e:
        app.logger.error(f"Error overlaying on PNG: {str(e)}")
        raise

# Send email with the PDF and PNG attachment
def send_email(pdf_filename, png_filename, data):
    try:
        if pdf_filename and not os.path.exists(pdf_filename):
            raise FileNotFoundError(f"PDF file '{pdf_filename}' not found")
        if png_filename and not os.path.exists(png_filename):
            raise FileNotFoundError(f"PNG file '{png_filename}' not found")

        subject = f"Purchase Order Submission - {data['Vendor Name']}"
        body = f"""
        Dear {EMAIL_RECEIVER},

        Please find the attached purchase order details.

        **Purchase Details:**
        - Name of Purchaser: {data['Name of Purchaser']}
        - Vendor Name: {data['Vendor Name']}
        - Date of Purchase: {data['Date of Purchase']}
        - Purpose of Purchase: {data['Purpose of Purchase']}

        The completed purchase order is attached as a PDF and PNG.

        Best,
        {data['Name of Purchaser']}
        """

        try:
            app.logger.info("Initializing email connection...")
            yag = yagmail.SMTP(EMAIL_SENDER, EMAIL_PASSWORD)
            
            app.logger.info("Sending email...")
            yag.send(to=EMAIL_RECEIVER, subject=subject, contents=body, attachments=[pdf_filename, png_filename])
            
            app.logger.info("Email sent successfully")
        except Exception as email_error:
            app.logger.error(f"Detailed email error: {str(email_error)}")
            app.logger.error(f"Error type: {type(email_error).__name__}")
            raise
    except Exception as e:
        app.logger.error(f"Error sending email: {str(e)}")
        raise

# Google Sheets API setup
def get_sheets_service():
    try:
        # Path to your service account credentials JSON file
        SERVICE_ACCOUNT_FILE = 'service-account-credentials.json'
        
        # Check if credentials file exists
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            app.logger.error(f"Service account credentials file not found: {SERVICE_ACCOUNT_FILE}")
            return None
            
        # Create credentials from service account file
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, 
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        
        # Build the Sheets API service
        service = build('sheets', 'v4', credentials=credentials)
        return service
    except Exception as e:
        app.logger.error(f"Error setting up Google Sheets service: {str(e)}")
        return None

# Upload PDF to Google Drive and get shareable link
def upload_to_drive(file_path, file_name=None):
    try:
        if not os.path.exists(file_path):
            app.logger.error(f"File not found: {file_path}")
            return None
            
        # Get Drive service
        drive_service = get_drive_service()
        if not drive_service:
            app.logger.error("Failed to initialize Google Drive service")
            return None
            
        # Use the provided file name or extract from path
        if not file_name:
            file_name = os.path.basename(file_path)
            
        # File metadata
        file_metadata = {
            'name': file_name,
            'mimeType': 'application/pdf'
        }
        
        # Create media object
        media = MediaFileUpload(
            file_path,
            mimetype='application/pdf',
            resumable=True
        )
        
        # Upload file
        app.logger.info(f"Uploading file to Google Drive: {file_name}")
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id,webViewLink'
        ).execute()
        
        # Make the file publicly accessible (anyone with the link can view)
        permission = {
            'type': 'anyone',
            'role': 'reader'
        }
        drive_service.permissions().create(
            fileId=file.get('id'),
            body=permission
        ).execute()
        
        # Get the web view link
        file = drive_service.files().get(
            fileId=file.get('id'),
            fields='webViewLink'
        ).execute()
        
        app.logger.info(f"File uploaded successfully. Link: {file.get('webViewLink')}")
        return file.get('webViewLink')
        
    except Exception as e:
        app.logger.error(f"Error uploading file to Google Drive: {str(e)}")
        return None

# Submit data to Google Sheets
def submit_to_google_sheet(data, pdf_filename=None):
    try:
        # Google Sheet ID - replace with your actual sheet ID
        SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'
        RANGE_NAME = 'Sheet1!A:H'  # Adjust based on your sheet structure
        
        # Get Sheets service
        sheets_service = get_sheets_service()
        if not sheets_service:
            app.logger.error("Failed to initialize Google Sheets service")
            return False, "Failed to connect to Google Sheets"
        
        # Calculate total cost
        total_cost = 0
        for item in data.get("items", []):
            try:
                quantity = float(item.get("quantity", 0))
                price = float(item.get("price", 0))
                total_cost += quantity * price
            except (ValueError, TypeError):
                app.logger.warning(f"Could not calculate cost for item: {item}")
        
        # Format items as a list
        items_text = []
        for item in data.get("items", []):
            item_text = f"{item['name']} - Qty: {item['quantity']} - Price: ${item['price']} - Catalog: {item['catalog_number']}"
            items_text.append(item_text)
        
        # Combine all items into a single string
        items_combined = "\n".join(items_text)
        
        # Upload PDF to Google Drive if provided
        pdf_link = None
        if pdf_filename and os.path.exists(pdf_filename):
            pdf_link = upload_to_drive(pdf_filename)
            if not pdf_link:
                app.logger.warning(f"Failed to upload PDF to Google Drive: {pdf_filename}")
        
        # Prepare row data
        row_data = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # Timestamp
            data["Name of Purchaser"],                     # Name of Purchaser
            data["Vendor Name"],                           # Vendor
            items_combined,                                # Item(s)
            data["Date of Purchase"],                      # Date of Purchase
            f"${total_cost:.2f}",                          # Cost
            data["Purpose of Purchase"],                   # Purpose
            pdf_link if pdf_link else "No PDF available"   # PDF Link
        ]
        
        # Prepare the request body
        body = {
            'values': [row_data]
        }
        
        app.logger.info(f"Submitting to Google Sheet: {SPREADSHEET_ID}")
        app.logger.info(f"Row data: {row_data}")
        
        # Append the data to the sheet
        result = sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME,
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        
        # Check if the append was successful
        if 'updates' in result and result['updates']['updatedRows'] > 0:
            app.logger.info(f"Data successfully appended to Google Sheet. Updated {result['updates']['updatedRows']} rows.")
            return True, "Data successfully submitted to Google Sheet"
        else:
            app.logger.warning("Google Sheet update response doesn't contain expected success markers")
            return False, "Data submission may have failed. Check sheet configuration."
            
    except Exception as e:
        app.logger.error(f"Error submitting to Google Sheet: {str(e)}")
        return False, f"Error submitting data: {str(e)}"

@app.route("/", methods=["GET", "POST"])
def index():
    try:
        if request.method == "POST":
            app.logger.info("Form submitted. Processing data...")
            
            # Validate form data
            required_fields = ["name", "vendor", "date", "telephone", "purpose"]
            missing_fields = [field for field in required_fields if field not in request.form]
            if missing_fields:
                app.logger.warning(f"Missing required fields: {missing_fields}")
                return f"Missing required fields: {', '.join(missing_fields)}", 400

            # Standardized form data structure
            form_data = {
                "name": request.form["name"],
                "vendor": request.form["vendor"],
                "vendor_address": request.form.get("vendor_address", ""),
                "date": request.form["date"],
                "purpose": request.form["purpose"],
                "telephone": request.form["telephone"],
                "fax": request.form.get("fax", ""),
                "notes": request.form.get("notes", ""),
                "items": []
            }

            # Process items with correct field mapping
            item_names = request.form.getlist("item_name[]")
            if not item_names:
                return "Please add at least one item to the purchase order", 400
                
            item_quantities = request.form.getlist("item_quantity[]")
            item_prices = request.form.getlist("item_price[]")
            item_catalogs = request.form.getlist("catalog_number[]")

            for i in range(len(item_names)):
                try:
                    quantity = float(item_quantities[i])
                    price = float(item_prices[i])
                    if quantity <= 0 or price < 0:
                        raise ValueError
                except ValueError:
                    return "Invalid quantity or price values", 400

                form_data["items"].append({
                    "description": item_names[i],      # Changed from "name" to "description"
                    "quantity": quantity,              # Store as float
                    "unit_price": price,               # Changed from "price" to "unit_price"
                    "catalog_number": item_catalogs[i] if i < len(item_catalogs) else ""
                })

            # Generate documents
            pdf_filename = generate_pdf(form_data)
            png_filename = overlay_on_png(form_data)
            
            # Return the PDF file
            return send_file(pdf_filename, as_attachment=True, download_name="requisition_form.pdf")

        return render_template("form.html")
    except Exception as e:
        app.logger.error(f"Error: {str(e)}")
        return "An error occurred processing your request.", 500

def send_email_smtp(pdf_filename, png_filename, data):
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECEIVER
        msg['Subject'] = f"Purchase Order Submission - {data.get('vendor', 'Unknown Vendor')}"
        
        # Add body with correct data keys
        body = f"""
        Dear {EMAIL_RECEIVER},
        
        Please find the attached purchase order details.
        
        **Purchase Details:**
        - Name of Purchaser: {data.get('name', 'Unknown')}
        - Vendor Name: {data.get('vendor', 'Unknown')}
        - Date of Purchase: {data.get('date', 'Unknown')}
        - Purpose of Purchase: {data.get('purpose', 'Unknown')}
        - Phone: {data.get('telephone', 'N/A')}
        - Total Items: {len(data.get('items', []))}
        
        The completed purchase order is attached as a PDF and PNG.
        
        Best,
        {data.get('name', 'Unknown')}
        """
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach PDF
        with open(pdf_filename, 'rb') as f:
            pdf_attachment = MIMEApplication(f.read(), _subtype='pdf')
            pdf_attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_filename))
            msg.attach(pdf_attachment)
        
        # Attach PNG
        with open(png_filename, 'rb') as f:
            png_attachment = MIMEApplication(f.read(), _subtype='png')
            png_attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(png_filename))
            msg.attach(png_attachment)
        
        # Connect to server and send
        smtp_server = 'smtp.virginia.edu'  # Replace with your university's SMTP server
        smtp_port = 587  # Common port, may need to be changed
        
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        
        return True
    except Exception as e:
        app.logger.error(f"SMTP Error: {str(e)}")
        raise

@app.route('/Req_Forms/<filename>')
def uploaded_file(filename):
    """Serve files from the Req_Forms directory"""
    return send_from_directory('Req_Forms', filename)

@app.route("/calibrate")
def calibrate():
    """Route to generate test overlays for position calibration"""
    sample_data = {
        "vendor": "Sample Vendor Very Long Name That Might Overflow",
        "name": "Test User With Long Name",
        "date": datetime.now().strftime("%m/%d/%Y"),
        "purpose": "Very Long PTAO Purpose That Exceeds Character Limits",
        "telephone": "555-123-4567",
        "fax": "555-987-6543",
        "notes": "These are test notes for calibration",
        "items": [
            {
                "description": "Very Long Item Description That Should Be Truncated Properly",
                "quantity": 1,
                "unit_price": 1299.99,
                "catalog_number": "VERYLONGCATALOG123"
            },
            {
                "description": "Another Long Test Item",
                "quantity": 2,
                "unit_price": 799.50,
                "catalog_number": "SHORT"
            }
        ]
    }
    
    # Generate a debug overlay
    debug_path = overlay_on_png(sample_data, debug=True)
    
    return f"""
    <h1>Position Calibration Tool</h1>
    <p>A debug overlay has been generated: {os.path.basename(debug_path)}</p>
    <p>Open this file to see the coordinate grid and verify text positioning.</p>
    <p><a href="/Req_Forms/{os.path.basename(debug_path)}" target="_blank">View Debug Image</a></p>
    <p>Adjust the POSITIONS values in positions.py based on the grid if needed.</p>
    """

if __name__ == "__main__":
    # Create directories if they don't exist
    os.makedirs("Req_Forms", exist_ok=True)
    
    # Debug information
    print(f"Template image exists: {os.path.exists(TEMPLATE_IMAGE_PATH)}")
    print(f"Templates directory exists: {os.path.exists('templates')}")
    print(f"Form template exists: {os.path.exists('templates/form.html')}")
    
    app.run(debug=True, host="127.0.0.1", port=5000)
