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
import logging

app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.INFO)

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

def calculate_field_widths():
    """Calculate actual available width for each field based on form layout"""
    # PDF dimensions
    width, height = letter
    x_scale = width / 15
    
    return {
        "vendor_name": 6.5 * x_scale,  # Based on actual form space
        "ptao": 4.0 * x_scale,
        "requestor": 3.5 * x_scale,
        "description": 7.0 * x_scale,
        "catalog": 1.8 * x_scale,
        "phone": 2.5 * x_scale,
        "fax": 2.5 * x_scale,
        "notes": 8.0 * x_scale
    }

def smart_text_scaling_pdf(canvas_obj, x, y, text, max_width, font_name="Helvetica", 
                          initial_size=10, min_size=6, max_shrink_ratio=0.6):
    """
    Enhanced PDF text scaling with context-aware limits
    
    Args:
        max_shrink_ratio: Maximum allowed shrinkage (0.6 = can shrink to 60% of original)
    """
    if not text:
        return
    
    # Calculate minimum allowed font size based on field importance
    absolute_min_size = max(min_size, initial_size * max_shrink_ratio)
    
    current_size = initial_size
    
    while current_size >= absolute_min_size:
        text_width = stringWidth(text, font_name, current_size)
        
        if text_width <= max_width:
            canvas_obj.setFont(font_name, current_size)
            canvas_obj.drawString(x, y, text)
            return current_size  # Return the size used
        
        current_size -= 0.5
    
    # If we still can't fit, use minimum size and possibly truncate
    canvas_obj.setFont(font_name, absolute_min_size)
    
    # Check if we need to truncate even at minimum size
    text_width = stringWidth(text, font_name, absolute_min_size)
    if text_width > max_width:
        # Calculate how many characters we can fit
        char_width = text_width / len(text)
        max_chars = int(max_width / char_width) - 3  # Leave room for "..."
        text = text[:max(1, max_chars)] + "..."
    
    canvas_obj.drawString(x, y, text)
    return absolute_min_size

def smart_text_scaling_pdf_with_wrapping(canvas_obj, x, y, text, max_width, max_lines=2, 
                                        font_name="Helvetica", initial_size=10, 
                                        min_size=6, max_shrink_ratio=0.6, line_spacing=12):
    """
    Enhanced PDF text scaling with multi-line wrapping fallback
    
    Args:
        max_lines: Maximum number of lines to allow (default 2)
        line_spacing: Space between lines in points
    """
    if not text:
        return
    
    # Calculate minimum allowed font size based on field importance
    absolute_min_size = max(min_size, initial_size * max_shrink_ratio)
    
    # Try single-line scaling first
    current_size = initial_size
    while current_size >= absolute_min_size:
        text_width = stringWidth(text, font_name, current_size)
        
        if text_width <= max_width:
            # Fits on single line
            canvas_obj.setFont(font_name, current_size)
            canvas_obj.drawString(x, y, text)
            return current_size
        
        current_size -= 0.5
    
    # Single line doesn't fit even at minimum size - try multi-line
    if max_lines > 1:
        canvas_obj.setFont(font_name, absolute_min_size)
        wrapped_lines = wrap_text_to_lines(text, max_width, font_name, absolute_min_size, max_lines)
        
        for i, line in enumerate(wrapped_lines):
            line_y = y - (i * line_spacing)
            canvas_obj.drawString(x, line_y, line)
        
        return absolute_min_size
    
    # Fallback: truncate single line
    canvas_obj.setFont(font_name, absolute_min_size)
    text_width = stringWidth(text, font_name, absolute_min_size)
    if text_width > max_width:
        char_width = text_width / len(text)
        max_chars = int(max_width / char_width) - 3
        text = text[:max(1, max_chars)] + "..."
    
    canvas_obj.drawString(x, y, text)
    return absolute_min_size

def wrap_text_to_lines(text, max_width, font_name, font_size, max_lines):
    """
    Wrap text into multiple lines that fit within max_width
    
    Returns list of strings, one per line
    """
    if not text:
        return [""]
    
    words = text.split()
    lines = []
    current_line = []
    
    for word in words:
        # Test if adding this word exceeds width
        test_line = " ".join(current_line + [word])
        test_width = stringWidth(test_line, font_name, font_size)
        
        if test_width <= max_width:
            current_line.append(word)
        else:
            # Start new line
            if current_line:
                lines.append(" ".join(current_line))
                if len(lines) >= max_lines:
                    break
                current_line = [word]
            else:
                # Single word is too long - truncate it
                truncated_word = truncate_text_by_width(word, max_width, font_name, font_size)
                lines.append(truncated_word)
                if len(lines) >= max_lines:
                    break
    
    # Add remaining words to last line if there's room
    if current_line and len(lines) < max_lines:
        lines.append(" ".join(current_line))
    elif current_line and len(lines) == max_lines:
        # Merge remaining words into last line with truncation
        remaining_text = " ".join(current_line)
        last_line = lines[-1] + " " + remaining_text
        truncated_last = truncate_text_by_width(last_line, max_width, font_name, font_size)
        lines[-1] = truncated_last
    
    return lines[:max_lines]  # Ensure we don't exceed max_lines

def truncate_text_by_width(text, max_width, font_name, font_size, suffix="..."):
    """Truncate text to fit within pixel width"""
    if stringWidth(text, font_name, font_size) <= max_width:
        return text
    
    # Binary search for the right length
    left, right = 0, len(text)
    best_fit = ""
    
    while left <= right:
        mid = (left + right) // 2
        test_text = text[:mid] + suffix
        test_width = stringWidth(test_text, font_name, font_size)
        
        if test_width <= max_width:
            best_fit = test_text
            left = mid + 1
        else:
            right = mid - 1
    
    return best_fit if best_fit else suffix

def smart_text_scaling_png(draw, x, y, text, max_width, initial_size=12, 
                          min_size=7, max_shrink_ratio=0.6):
    """Enhanced PNG text scaling"""
    if not text:
        return
    
    absolute_min_size = max(min_size, initial_size * max_shrink_ratio)
    current_size = initial_size
    
    while current_size >= absolute_min_size:
        font = load_font(int(current_size))
        if font:
            text_width = get_text_width_pil(text, font)
            if text_width <= max_width:
                draw.text((x, y), text, font=font, fill="black")
                return current_size
        current_size -= 1
    
    # Use minimum size with possible truncation
    font = load_font(int(absolute_min_size))
    if font:
        text_width = get_text_width_pil(text, font)
        if text_width > max_width:
            # Estimate character limit
            chars_per_pixel = len(text) / text_width if text_width > 0 else 0.1
            max_chars = int(max_width * chars_per_pixel) - 3
            text = text[:max(1, max_chars)] + "..."
        
        draw.text((x, y), text, font=font, fill="black")
    
    return absolute_min_size

def smart_text_scaling_png_with_wrapping(draw, x, y, text, max_width, max_lines=2,
                                        initial_size=12, min_size=7, max_shrink_ratio=0.6, 
                                        line_spacing=15):
    """Enhanced PNG text scaling with multi-line wrapping"""
    if not text:
        return
    
    absolute_min_size = max(min_size, initial_size * max_shrink_ratio)
    current_size = initial_size
    
    # Try single-line scaling first
    while current_size >= absolute_min_size:
        font = load_font(int(current_size))
        if font:
            text_width = get_text_width_pil(text, font)
            if text_width <= max_width:
                draw.text((x, y), text, font=font, fill="black")
                return current_size
        current_size -= 1
    
    # Single line doesn't fit - try multi-line
    if max_lines > 1:
        font = load_font(int(absolute_min_size))
        if font:
            wrapped_lines = wrap_text_to_lines_pil(text, max_width, font, max_lines)
            
            for i, line in enumerate(wrapped_lines):
                line_y = y + (i * line_spacing)
                draw.text((x, line_y), line, font=font, fill="black")
            
            return absolute_min_size
    
    # Fallback: truncate single line
    font = load_font(int(absolute_min_size))
    if font:
        text_width = get_text_width_pil(text, font)
        if text_width > max_width:
            chars_per_pixel = len(text) / text_width if text_width > 0 else 0.1
            max_chars = int(max_width * chars_per_pixel) - 3
            text = text[:max(1, max_chars)] + "..."
        
        draw.text((x, y), text, font=font, fill="black")
    
    return absolute_min_size

def wrap_text_to_lines_pil(text, max_width, font, max_lines):
    """Wrap text for PIL/PNG rendering"""
    if not text:
        return [""]
    
    words = text.split()
    lines = []
    current_line = []
    
    for word in words:
        test_line = " ".join(current_line + [word])
        test_width = get_text_width_pil(test_line, font)
        
        if test_width <= max_width:
            current_line.append(word)
        else:
            if current_line:
                lines.append(" ".join(current_line))
                if len(lines) >= max_lines:
                    break
                current_line = [word]
            else:
                # Single word too long - truncate
                truncated_word = truncate_text_by_width_pil(word, max_width, font)
                lines.append(truncated_word)
                if len(lines) >= max_lines:
                    break
    
    if current_line and len(lines) < max_lines:
        lines.append(" ".join(current_line))
    elif current_line and len(lines) == max_lines:
        remaining_text = " ".join(current_line)
        last_line = lines[-1] + " " + remaining_text
        truncated_last = truncate_text_by_width_pil(last_line, max_width, font)
        lines[-1] = truncated_last
    
    return lines[:max_lines]

def truncate_text_by_width_pil(text, max_width, font, suffix="..."):
    """Truncate text for PIL rendering"""
    if get_text_width_pil(text, font) <= max_width:
        return text
    
    # Simple approach: remove characters until it fits
    while len(text) > 1:
        test_text = text[:-1] + suffix
        if get_text_width_pil(test_text, font) <= max_width:
            return test_text
        text = text[:-1]
    
    return suffix

def generate_pdf(data):
    """Generate PDF with enhanced text handling including multi-line wrapping"""
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
        c.setFont(font_name, SPACING["font_size"])
        
        # PDF dimensions and coordinate system conversion
        width, height = letter
        x_scale = width / 15
        y_scale = height / 14
        
        def y_pos(latex_y):
            return height - (latex_y * y_scale)
        
        # Calculate context-aware field widths
        field_widths = calculate_field_widths()
        
        # Font scaling rules based on field importance
        scaling_rules = {
            "vendor_name": {"min_ratio": 0.7, "initial": 11},  # Vendor name is important
            "ptao": {"min_ratio": 0.6, "initial": 10},          # PTAO can be smaller
            "requestor": {"min_ratio": 0.7, "initial": 10},      # Name should be readable
            "description": {"min_ratio": 0.5, "initial": 9},     # Descriptions can be very small
            "catalog": {"min_ratio": 0.6, "initial": 8},         # Catalog numbers can be small
            "phone": {"min_ratio": 0.7, "initial": 9},
            "fax": {"min_ratio": 0.7, "initial": 9},
            "notes": {"min_ratio": 0.5, "initial": 8}
        }
        
        # Extract and format data with character limits
        vendor_name = data.get("vendor", "") or data.get("Vendor Name", "")
        ptao = data.get("purpose", "") or data.get("Purpose of Purchase", "")
        date = data.get("date", datetime.now().strftime("%m/%d/%Y"))
        req = data.get("name", "") or data.get("Name of Purchaser", "")
        
        # Calculate field-specific line limits
        field_line_limits = {
            "vendor_name": 2,     # Vendor names can have 2 lines
            "ptao": 2,            # PTAO can have 2 lines
            "requestor": 1,       # Names should stay single line
            "description": 3,     # Item descriptions can have up to 3 lines
            "catalog": 1,         # Catalog numbers should be single line
        }
        
        # Vendor name with multi-line wrapping
        vendor_x, vendor_y = POSITIONS["vendor_name"]
        rules = scaling_rules["vendor_name"]
        smart_text_scaling_pdf_with_wrapping(
            c, 
            vendor_x * x_scale, 
            y_pos(vendor_y),
            vendor_name,
            field_widths["vendor_name"],
            max_lines=field_line_limits["vendor_name"],
            initial_size=rules["initial"],
            max_shrink_ratio=rules["min_ratio"],
            line_spacing=10  # Tighter spacing for form fields
        )
        
        # Date
        date_x, date_y = POSITIONS["date"]
        c.setFont(font_name, SPACING["font_size"])
        c.drawString(date_x * x_scale, y_pos(date_y), date)
        
        # PTAO with multi-line support
        ptao_x, ptao_y = POSITIONS["ptao"]
        rules = scaling_rules["ptao"]
        smart_text_scaling_pdf_with_wrapping(
            c,
            ptao_x * x_scale,
            y_pos(ptao_y),
            ptao,
            field_widths["ptao"],
            max_lines=field_line_limits["ptao"],
            initial_size=rules["initial"],
            max_shrink_ratio=rules["min_ratio"],
            line_spacing=10
        )
        
        # Process items with multi-line descriptions
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
                c.setFont(font_name, SPACING["font_size"])
                c.drawString(TABLE_COLUMNS["quantity"] * x_scale, row_y, str(int(quantity)))
                
                # Catalog number - can be quite small
                rules = scaling_rules["catalog"]
                smart_text_scaling_pdf(
                    c,
                    TABLE_COLUMNS["catalog"] * x_scale,
                    row_y,
                    catalog_num,
                    field_widths["catalog"],
                    initial_size=rules["initial"],
                    max_shrink_ratio=rules["min_ratio"]
                )
                
                # Description - most permissive scaling
                rules = scaling_rules["description"]
                smart_text_scaling_pdf_with_wrapping(
                    c,
                    TABLE_COLUMNS["description"] * x_scale,
                    row_y,
                    description,
                    field_widths["description"],
                    max_lines=field_line_limits["description"],  # Allow up to 3 lines
                    initial_size=rules["initial"],
                    max_shrink_ratio=rules["min_ratio"],
                    line_spacing=8  # Very tight for table rows
                )
                
                # Prices (simple)
                c.setFont(font_name, SPACING["font_size"])
                c.drawString(TABLE_COLUMNS["unit_price"] * x_scale, row_y, f"${unit_price:.2f}")
                c.drawString(TABLE_COLUMNS["total_price"] * x_scale, row_y, f"${total_price:.2f}")
                
            except (ValueError, TypeError) as e:
                app.logger.warning(f"Error processing item {i}: {e}")
                continue
        
        # Total amount
        total_x, total_y = POSITIONS["total"]
        c.setFont(font_name, SPACING["font_size"])
        c.drawString(total_x * x_scale, y_pos(total_y), f"${total_amount:.2f}")
        
        # Requestor with truncation
        req_x, req_y = POSITIONS["requestor"]
        rules = scaling_rules["requestor"]
        smart_text_scaling_pdf(
            c,
            req_x * x_scale,
            y_pos(req_y),
            req,
            field_widths["requestor"],
            initial_size=rules["initial"],
            max_shrink_ratio=rules["min_ratio"]
        )
        
        # Approver
        approver_x, approver_y = POSITIONS["approver"]
        c.setFont(font_name, SPACING["font_size"])
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
    """Enhanced PNG overlay with multi-line text support"""
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
        
        # Calculate field widths for PNG (slightly larger than PDF)
        field_widths = {
            "vendor_name": 400,
            "ptao": 250,
            "requestor": 200,
            "description": 450,
            "catalog": 130
        }
        
        # Extract data with correct keys
        vendor_name = data.get("vendor", "")
        ptao = data.get("purpose", "")
        date = data.get("date", datetime.now().strftime("%m/%d/%Y"))
        req = data.get("name", "")
        
        # Vendor name with smart placement
        vendor_name_x, vendor_name_y = POSITIONS["vendor_name"]
        smart_text_scaling_png_with_wrapping(
            draw,
            vendor_name_x * x_scale,
            vendor_name_y * y_scale,
            vendor_name,
            field_widths["vendor_name"],
            max_lines=2,
            initial_size=12,
            max_shrink_ratio=0.7,
            line_spacing=15
        )
        
        # PTAO with multi-line support
        ptao_x, ptao_y = POSITIONS["ptao"]
        smart_text_scaling_png_with_wrapping(
            draw,
            ptao_x * x_scale,
            ptao_y * y_scale,
            ptao,
            field_widths["ptao"],
            max_lines=2,
            initial_size=11,
            max_shrink_ratio=0.6,
            line_spacing=15
        )
        
        # Process items with multi-line descriptions
        for i, item in enumerate(items[:9]):
            try:
                quantity = float(item.get("quantity", 0))
                catalog_num = item.get("catalog_number", "")
                description = item.get("description", "") or item.get("name", "")
                unit_price = float(item.get("unit_price", 0)) or float(item.get("price", 0))
                
                total_price = quantity * unit_price
                total_amount += total_price
                
                row_y = y_start + (i * SPACING["table_row"])
                
                # Quantity (simple)
                draw.text((TABLE_COLUMNS["quantity"] * x_scale, row_y), str(int(quantity)), font=font, fill="black")
                
                # Catalog number with scaling
                smart_text_scaling_png(
                    draw,
                    TABLE_COLUMNS["catalog"] * x_scale,
                    row_y,
                    catalog_num,
                    field_widths["catalog"],
                    initial_size=10,
                    max_shrink_ratio=0.6
                )
                
                # Description with multi-line wrapping
                smart_text_scaling_png_with_wrapping(
                    draw,
                    TABLE_COLUMNS["description"] * x_scale,
                    row_y,
                    description,
                    field_widths["description"],
                    max_lines=3,  # Allow up to 3 lines for descriptions
                    initial_size=10,
                    max_shrink_ratio=0.5,
                    line_spacing=12
                )
                
                # Prices (simple)
                draw.text((TABLE_COLUMNS["unit_price"] * x_scale, row_y), f"${unit_price:.2f}", font=font, fill="black")
                draw.text((TABLE_COLUMNS["total_price"] * x_scale, row_y), f"${total_price:.2f}", font=font, fill="black")
                
            except (ValueError, TypeError) as e:
                app.logger.warning(f"Error processing PNG item {i}: {e}")
                continue
        
        # Total and signatures
        draw.text((POSITIONS["total"][0] * x_scale, POSITIONS["total"][1] * y_scale), 
                  f"${total_amount:.2f}", font=font, fill="black")
        
        smart_text_scaling_png(
            draw,
            POSITIONS["requestor"][0] * x_scale,
            POSITIONS["requestor"][1] * y_scale,
            req,
            field_widths["requestor"],
            initial_size=12,
            max_shrink_ratio=0.7
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

        # Use consistent data keys
        subject = f"Purchase Order Submission - {data.get('vendor', 'Unknown Vendor')}"
        body = f"""
        Dear {EMAIL_RECEIVER},

        Please find the attached purchase order details.

        **Purchase Details:**
        - Name of Purchaser: {data.get('name', 'Unknown')}
        - Vendor Name: {data.get('vendor', 'Unknown')}
        - Date of Purchase: {data.get('date', 'Unknown')}
        - Purpose of Purchase: {data.get('purpose', 'Unknown')}

        The completed purchase order is attached as a PDF and PNG.

        Best,
        {data.get('name', 'Unknown')}
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

# Google Drive service (add this missing function)
def get_drive_service():
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
            scopes=['https://www.googleapis.com/auth/drive']
        )
        
        # Build the Drive API service
        service = build('drive', 'v3', credentials=credentials)
        return service
    except Exception as e:
        app.logger.error(f"Error setting up Google Drive service: {str(e)}")
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
                price = float(item.get("unit_price", 0))
                total_cost += quantity * price
            except (ValueError, TypeError):
                app.logger.warning(f"Could not calculate cost for item: {item}")
        
        # Format items as a list
        items_text = []
        for item in data.get("items", []):
            item_text = f"{item['description']} - Qty: {item['quantity']} - Price: ${item['unit_price']} - Catalog: {item['catalog_number']}"
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
            data["name"],                                  # Name of Purchaser
            data["vendor"],                                # Vendor
            items_combined,                                # Item(s)
            data["date"],                                  # Date of Purchase
            f"${total_cost:.2f}",                          # Cost
            data["purpose"],                               # Purpose
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
                    "description": item_names[i],  # Always use "description"
                    "quantity": quantity,
                    "unit_price": price,          # Always use "unit_price"  
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

if __name__ == '__main__':
    app.run(debug=True)
