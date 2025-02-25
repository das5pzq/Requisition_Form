from flask import Flask, request, send_file, render_template
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

app = Flask(__name__)

# Google Form URL (Change to your form URL)
GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeL97014pGTbbwJR-t5BktA3kjc_v0Af3t2NBhzcnEp5MLL7g/formResponse"

# Email credentials
EMAIL_SENDER = "xxdrp3pp3rxlov3rxx@gmail.com"
EMAIL_PASSWORD = "nyty jsgq btqd xfzu"
EMAIL_RECEIVER = "xxdrp3pp3rxlov3rxx@gmail.com"

# Path to your template PNG file
TEMPLATE_IMAGE_PATH = "ir-raw.png"

def generate_pdf(data):
    """Generate PDF with form data placed at predefined positions based on PHP format"""
    try:
        # Create Req_Forms directory if it doesn't exist
        output_dir = "Req_Forms"
        os.makedirs(output_dir, exist_ok=True)
        
        pdf_filename = os.path.join(output_dir, "purchase_order.pdf")
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
        vendor_name = data.get("Vendor Name", "")
        address = data.get("vendor_address", "")
        tel = data.get("tel", "")
        fax = data.get("fax", "")
        ptao = data.get("Purpose of Purchase", "")
        date = data.get("Date of Purchase", datetime.now().strftime("%B %d, %Y"))
        notes = data.get("notes", "")
        req = data.get("Name of Purchaser", "")
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
        
        # PTAO/Purpose
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
        return pdf_filename
    except Exception as e:
        app.logger.error(f"Error generating PDF: {str(e)}")
        raise

def overlay_on_png(data, debug=False):
    """Overlay form data onto the PNG template with optional debug grid"""
    try:
        if not os.path.exists(TEMPLATE_IMAGE_PATH):
            raise FileNotFoundError(f"Template image '{TEMPLATE_IMAGE_PATH}' not found")
            
        img = Image.open(TEMPLATE_IMAGE_PATH)
        draw = ImageDraw.Draw(img)

        # Get image dimensions
        width, height = img.size
        
        # Scale factors to convert from LaTeX coordinates to pixels
        x_scale = width / 15  # Assuming LaTeX uses ~15 units width
        y_scale = height / 14  # Assuming LaTeX uses ~14 units height
        
        # Draw debug grid if requested
        if debug:
            # Draw horizontal and vertical grid lines
            for i in range(0, 15):
                # Vertical lines
                x = i * x_scale
                draw.line([(x, 0), (x, height)], fill="red", width=1)
                draw.text((x, 10), f"{i}", fill="red")
                
            for i in range(0, 14):
                # Horizontal lines
                y = i * y_scale
                draw.line([(0, y), (width, y)], fill="blue", width=1)
                draw.text((10, y), f"{i}", fill="blue")
                
            # For key positions, draw small markers
            positions = [
                (3.2, 2.25, "Supplier"),
                (3.2, 4.1, "Phone"),
                (3.2, 4.7, "Fax"),
                (10, 1.75, "Date"),
                (10, 2.3, "PTAO"),
                (0.445, 6, "Items"),
                (13.4, 11.18, "Total"),
                (0.85, 11.5, "Notes"),
                (3.25, 12.65, "Requested by")
            ]
            
            for x, y, label in positions:
                px = x * x_scale
                py = y * y_scale
                r = 5  # radius
                draw.ellipse([(px-r, py-r), (px+r, py+r)], outline="green")
                draw.text((px+r+2, py), label, fill="green")
        
        # Current date
        date = data.get("Date of Purchase", datetime.now().strftime("%B %d, %Y"))
        
        # Extract supplier/vendor information
        vendor_name = data.get("Vendor Name", "")
        address = data.get("vendor_address", "")
        tel = data.get("tel", "")
        fax = data.get("fax", "")
        ptao = data.get("Purpose of Purchase", "")
        notes = data.get("notes", "")
        req = data.get("Name of Purchaser", "")
        approver = "Dustin Keller"  # Hardcoded approver name
        
        # Load font
        font = ImageFont.load_default()  # You should define or load a proper font
        
        # Vendor name
        vendor_name_x, vendor_name_y = POSITIONS["vendor_name"]
        draw.text((vendor_name_x * x_scale, vendor_name_y * y_scale), vendor_name, font=font, fill="black")
        
        # Supplier/vendor address
        y_pos = POSITIONS["vendor_address"][1] * y_scale
        for line in address.split("\n"):
            draw.text((POSITIONS["vendor_address"][0] * x_scale, y_pos), line, font=font, fill="black")
            y_pos += 20
        
        # Phone
        draw.text((POSITIONS["phone"][0] * x_scale, POSITIONS["phone"][1] * y_scale), tel, font=font, fill="black")
        
        # Fax
        draw.text((POSITIONS["fax"][0] * x_scale, POSITIONS["fax"][1] * y_scale), fax, font=font, fill="black")
        
        # Date
        draw.text((POSITIONS["date"][0] * x_scale, POSITIONS["date"][1] * y_scale), date, font=font, fill="black")
        
        # PTAO/Purpose
        draw.text((POSITIONS["ptao"][0] * x_scale, POSITIONS["ptao"][1] * y_scale), ptao, font=font, fill="black")
        
        # Items table
        y_start = POSITIONS["items_start"][1] * y_scale
        total_amount = 0
        
        # Define column positions based on TABLE_COLUMNS
        x_quantity = TABLE_COLUMNS["quantity"] * x_scale
        x_catalog = TABLE_COLUMNS["catalog"] * x_scale
        x_desc = TABLE_COLUMNS["description"] * x_scale
        x_unit_price = TABLE_COLUMNS["unit_price"] * x_scale
        x_total_price = TABLE_COLUMNS["total_price"] * x_scale
        
        # Process items
        items = data.get("items", [])
        row_height = SPACING["table_row"] * y_scale
        
        for i, item in enumerate(items):
            if i < 9:  # Limit to 9 items as in PHP
                quantity = float(item.get("quantity", 0))
                catalog_num = item.get("catalog_number", "")
                description = item.get("name", "")
                unit_price = float(item.get("price", 0))
                
                total_price = quantity * unit_price
                total_amount += total_price
                
                # Draw item row
                row_y = y_start + (i * row_height)
                draw.text((x_quantity, row_y), str(quantity), font=font, fill="black")
                draw.text((x_catalog, row_y), catalog_num, font=font, fill="black")
                draw.text((x_desc, row_y), description, font=font, fill="black")
                draw.text((x_unit_price, row_y), f"${unit_price:.2f}", font=font, fill="black")
                draw.text((x_total_price, row_y), f"${total_price:.2f}", font=font, fill="black")
        
        # Total amount
        draw.text((POSITIONS["total"][0] * x_scale, POSITIONS["total"][1] * y_scale), 
                  f"${total_amount:.2f}", font=font, fill="black")
        
        # Notes
        if notes:
            draw.text((POSITIONS["notes"][0] * x_scale, POSITIONS["notes"][1] * y_scale), 
                      f"Notes: {notes}", font=font, fill="black")
        
        # Requested by (purchaser name)
        draw.text((POSITIONS["requestor"][0] * x_scale, POSITIONS["requestor"][1] * y_scale), 
                  req, font=font, fill="black")
        
        # Approved by
        draw.text((POSITIONS["approver"][0] * x_scale, POSITIONS["approver"][1] * y_scale), 
                  approver, font=font, fill="black")
        
        # Generate unique filename with debug indicator
        suffix = "_debug" if debug else ""
        file_base = f"{req.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d%H%M%S')}{suffix}"
        output_path = os.path.join("Req_Forms", f"{file_base}.png")
        
        # Create Req_Forms directory if it doesn't exist
        os.makedirs("Req_Forms", exist_ok=True)
        
        # Save the modified image
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
            
            # Debug: Print all form data
            app.logger.info(f"Form data: {request.form}")
            
            # Validate form data - removed vendor_city, vendor_state, vendor_zip
            required_fields = ["name", "vendor", "vendor_address", "date", "telephone", "purpose"]
            missing_fields = [field for field in required_fields if field not in request.form]
            if missing_fields:
                app.logger.warning(f"Missing required fields: {missing_fields}")
                return f"Missing required fields: {', '.join(missing_fields)}", 400

            # Use vendor_address directly without combining with city/state/zip
            form_data = {
                "Name of Purchaser": request.form["name"],
                "Vendor Name": request.form["vendor"],
                "vendor_address": request.form["vendor_address"],
                "Date of Purchase": request.form["date"],
                "Purpose of Purchase": request.form["purpose"],
                "tel": request.form["telephone"],
                "fax": request.form.get("fax", ""),
                "notes": request.form.get("notes", ""),
                "items": []
            }

            # Check if there are any items
            item_names = request.form.getlist("item_name[]")
            if not item_names:
                app.logger.warning("No items provided in the form")
                return "Please add at least one item to the purchase order", 400
                
            item_quantities = request.form.getlist("item_quantity[]")
            item_prices = request.form.getlist("item_price[]")
            item_catalogs = request.form.getlist("catalog_number[]")

            # Validate items data
            if len(item_names) != len(item_quantities) or len(item_names) != len(item_prices):
                app.logger.warning(f"Item data mismatch: names={len(item_names)}, quantities={len(item_quantities)}, prices={len(item_prices)}")
                return "Invalid items data", 400

            for i in range(len(item_names)):
                try:
                    quantity = float(item_quantities[i])
                    price = float(item_prices[i])
                    if quantity <= 0 or price < 0:
                        raise ValueError
                except ValueError:
                    app.logger.warning(f"Invalid quantity or price: quantity={item_quantities[i]}, price={item_prices[i]}")
                    return "Invalid quantity or price values", 400

                form_data["items"].append({
                    "name": item_names[i],
                    "quantity": item_quantities[i],
                    "price": item_prices[i],
                    "catalog_number": item_catalogs[i] if i < len(item_catalogs) else ""
                })

            app.logger.info("Form data validated successfully. Generating documents...")
            
            try:
                # Generate both PDF and PNG
                pdf_filename = generate_pdf(form_data)
                png_filename = overlay_on_png(form_data)
                app.logger.info(f"Documents generated: PDF={pdf_filename}, PNG={png_filename}")
            except Exception as e:
                app.logger.error(f"Document generation failed: {str(e)}", exc_info=True)
                return "Error: Failed to generate documents. Please check the form data and try again.", 500

            try:
                send_email(pdf_filename, png_filename, form_data)
                app.logger.info("Email sent successfully")
            except Exception as e:
                app.logger.error(f"Email sending failed: {str(e)}", exc_info=True)
                return "Error: Failed to send email. The documents were generated but couldn't be sent.", 500

            try:
                # Use Google Sheets instead of Google Forms
                success, message = submit_to_google_sheet(form_data, pdf_filename)
                if not success:
                    app.logger.warning(f"Google Sheet submission issue: {message}")
                    return f"Documents were generated and email was sent, but there was an issue with Google Sheet submission: {message}", 200
                app.logger.info("Google Sheet updated successfully")
            except Exception as e:
                app.logger.error(f"Google Sheet submission failed: {str(e)}", exc_info=True)
                return "Documents were generated and sent via email, but Google Sheet submission failed.", 200

            return "Purchase order processed successfully! Documents were generated, email was sent, and data was recorded in Google Sheets."

        return render_template("form.html")
    except Exception as e:
        app.logger.error(f"Unexpected error in index route: {str(e)}", exc_info=True)
        return "An unexpected error occurred. Please check the server logs for details.", 500

def send_email_smtp(pdf_filename, png_filename, data):
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECEIVER
        msg['Subject'] = f"Purchase Order Submission - {data['Vendor Name']}"
        
        # Add body
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

@app.route("/calibrate")
def calibrate():
    """Route to generate test overlays for position calibration"""
    sample_data = {
        "Vendor Name": "Sample Vendor\nLine 2\nLine 3",
        "Name of Purchaser": "Test User",
        "Date of Purchase": datetime.now().strftime("%B %d, %Y"),
        "Purpose of Purchase": "Test PTAO",
        "tel": "555-123-4567",
        "fax": "555-987-6543",
        "notes": "These are test notes",
        "items": [
            {"name": "Test Item 1", "quantity": "1", "price": "10.00", "catalog_number": "ABC123"},
            {"name": "Test Item 2", "quantity": "2", "price": "20.00", "catalog_number": "DEF456"}
        ]
    }
    
    # Generate a debug overlay
    debug_path = overlay_on_png(sample_data, debug=True)
    
    return f"""
    <h1>Position Calibration Tool</h1>
    <p>A debug overlay has been generated with grid lines at: {debug_path}</p>
    <p>Open this file to see the coordinate grid.</p>
    <p>Adjust the x_scale and y_scale values in the code based on the grid.</p>
    <img src="/{debug_path}" style="max-width: 100%">
    """

if __name__ == "__main__":
    app.run(debug=True)
