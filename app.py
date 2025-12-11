from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from barcode import Code128
from barcode.writer import ImageWriter
import win32print
import win32ui
from win32con import FW_NORMAL, FW_BOLD, DEFAULT_CHARSET
from PIL import Image, ImageWin
import os
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, SimpleDocTemplate, Paragraph, Spacer
import subprocess
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
import textwrap
from datetime import datetime
import math
from reportlab.lib.enums import TA_CENTER
import tempfile

pdfmetrics.registerFont(TTFont('Calibri', 'C:/Windows/Fonts/calibri.ttf')) 
pdfmetrics.registerFont(TTFont("Calibri-Bold", 'C:/Windows/Fonts/calibrib.ttf'))
pdfmetrics.registerFont(TTFont("Calibri-Italic", 'C:/Windows/Fonts/calibrii.ttf'))
pdfmetrics.registerFont(TTFont("Book Antiqua", r"C:\Windows\Fonts\ANTQUAI.TTF"))

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "assets", "uprising_logo.png")

# Page limits for packing slips
FIRST_PAGE_LIMIT = 27       # max rows (incl. blanks) on page 1
OTHER_PAGE_LIMIT = 40       # max rows (incl. blanks) on subsequent pages



app = Flask(__name__)
CORS(app) 

ROLL_PRINTER = "ZDesigner GX430t"
SHEET_PRINTER = "RICOH P 501"
CURRENT_USER = os.getlogin()
SUMATRA_PATH = r"C:\Users\seedy\AppData\Local\SumatraPDF\SumatraPDF.exe"


@app.route('/health', methods=['GET'])
def health_check():
    return {'status': 'ok'}, 200

def create_font(name, size, bold=False, italic=False):
    weight = FW_BOLD if bold else FW_NORMAL
    return win32ui.CreateFont({
        "name": name,
        "height": -size,  # Negative for point size
        "weight": weight,
        "italic": italic,
        "charset": DEFAULT_CHARSET,
    })

def format_variety_name_with_quotes(variety_name):
       """
       Format variety name with quotes only around the part outside parentheses.
       Example: "Cocaigne (Haricot di Pamplano)" -> "'Cocaigne' (Haricot di Pamplano)"
       Example: "Simple Variety" -> "'Simple Variety'"
       """
       if '(' in variety_name:
           parts = variety_name.split('(', 1)
           quoted_part = f"'{parts[0].strip()}'"
           return f"{quoted_part} ({parts[1]}"
       else:
           return f"'{variety_name}'"

@app.route('/print-germ-label', methods=['POST'])
def print_germ_label():
    try:
        data = request.get_json()

        if CURRENT_USER.lower() == "ndefe":
            print("=== GERM SAMPLE PRINT REQUEST ON NDEFE===")
            print(f"Variety Name: {data.get('variety_name')}")
            print(f"SKU Prefix: {data.get('sku_prefix')}")
            print(f"Species: {data.get('species')}")
            print(f"Lot Code: {data.get('lot_code')}")
            print(f"Germ Year: {data.get('germ_year')}")
            print("================================")
        
        else: 

            # === Construct label text ===
            variety = data.get('variety_name')
            sku_prefix = data.get('sku_prefix')
            species = data.get('species')
            lot_code = data.get('lot_code')

            lot_number = f"{sku_prefix}-{lot_code}"
            lot_text = f"Lot: {lot_number}"
            var_name = f"'{variety}'"

            # === Generate barcode image and save to file ===
            barcode = Code128(lot_number, writer=ImageWriter())
            barcode_file = barcode.save("barcode_temp", options={"write_text": False})

            # === Open barcode image safely ===
            with Image.open(barcode_file) as img:
                barcode_img = img.convert("RGB")

                # === Setup printer ===
                printer_name = ROLL_PRINTER
                dc = win32ui.CreateDC()
                dc.CreatePrinterDC(printer_name)
                dc.StartDoc("Seed Label")
                dc.StartPage()

                # === Label dimensions ===
                dpi = dc.GetDeviceCaps(88)  # LOGPIXELSX
                label_width = int(2.625 * dpi)
                label_height = int(1.0 * dpi)
                x_center = label_width // 2

                # === Text drawing ===
                font = create_font("Courier New", 44)
                dc.SelectObject(font)

                line_height = 45
                y_text = 25
                lines = [var_name, species, lot_text]
                for line in lines:
                    text_width = dc.GetTextExtent(line)[0]
                    dc.TextOut(x_center - text_width // 2, y_text, line)
                    y_text += line_height

                # === Resize barcode to fit ===
                target_width = int(label_width * 0.9)
                aspect_ratio = barcode_img.height / barcode_img.width
                target_height = int(target_width * aspect_ratio * 0.6)
                resized_barcode = barcode_img.resize((target_width, target_height))

                # === Draw barcode ===
                x_barcode = (label_width - target_width) // 2
                y_barcode = y_text + 5
                dib = ImageWin.Dib(resized_barcode)
                dib.draw(dc.GetHandleOutput(), (x_barcode, y_barcode, x_barcode + target_width, y_barcode + target_height))

                # === Finalize print job ===
                dc.EndPage()
                dc.EndDoc()
                dc.DeleteDC()

            # === Cleanup barcode file ===
            if os.path.exists(barcode_file):
                os.remove(barcode_file)

        return jsonify({
            'success': True,
            'message': 'Label printed successfully'
        })
        
    except Exception as e:
        print(f"Error printing germ label: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500





def print_single_front_label_logic(data):
    """Extract the core front label printing logic"""
    try:
        quantity = int(data.get('quantity', 1))
        env_multiplier = int(data.get('env_multiplier', 1))
        print(f"Environmental Multiplier: {env_multiplier}")
        quantity *= env_multiplier

        if CURRENT_USER.lower() == "ndefe":
            print(f"Printing copy {quantity} single front labels on Ndefe's printer")
            print(f"Variety Name: {data.get('variety_name')}")
            print(f"Crop: {data.get('crop')}")
            print(f"Days: {data.get('days')}")
            print(f"SKU Suffix: {data.get('sku_suffix')}")
            print(f"Common name: {data.get('common_name')}")
            print(f"Pkg. size: {data.get('pkg_size')}")
            print(f"Env type: {data.get('env_type')}")
            print(f"Lot Code: {data.get('lot_code')}")
            print(f"Germination: {data.get('germination')}")
            print(f"For Year: {data.get('for_year')}")
            print(f"Quantity: {data.get('quantity')}")
            print(f"Desc1: {data.get('desc1')}")
            print(f"Desc2: {data.get('desc2')}")
            print(f"Desc3: {data.get('desc3')}")
            print(f"Rad type: {data.get('rad_type')}")
            print("================================")
            return {'success': True, 'message': f'Front Single Label printed successfully ({quantity} copies)'}
        else:
            # Gather label content (shared across copies)
            # variety_name = f"'{data.get('variety_name')}'"
            variety_name = format_variety_name_with_quotes(data.get('variety_name'))
           
            # Check for common_name first, fall back to crop if empty
            common_name = data.get('common_name', '').strip()
            variety_crop = common_name if common_name else data.get('crop')

            days = data.get('days')
            env_type = data.get('env_type')
            year = data.get('for_year')
            days_year = f"{days}    Packed for 20{year}"

            desc_line1 = data.get('desc1')
            desc_line2 = data.get('desc2')
            desc_line3 = data.get('desc3')
            lot_code = data.get('lot_code')
            germination = data.get('germination')
            rad_type = data.get('rad_type')

            if env_type == "LG Coffee":
                pkg_size = f"{data.get('pkg_size')} ••"
            elif env_type == "SM Coffee":
                pkg_size = f"{data.get('pkg_size')} •"
            else:
                pkg_size = data.get('pkg_size')

            pkg_lot_germ = f"{pkg_size}    Lot: {lot_code}    Germ: {germination}%"
            sku_suffix = data.get('sku_suffix')

            # Fonts (reusable)
            bold_12 = create_font("Times New Roman", 48, bold=True)
            italic_9 = create_font("Times New Roman", 36, italic=True)
            normal_8 = create_font("Times New Roman", 32)
            bold_16 = create_font("Times New Roman", 60, bold=True)
            normal_12 = create_font("Times New Roman", 40)
            italic_12 = create_font("Times New Roman", 40, italic=True)

            printer_name = ROLL_PRINTER

            # Loop through each copy
            for i in range(quantity):
                dc = win32ui.CreateDC()
                dc.CreatePrinterDC(printer_name)

                dc.StartDoc("Seed Label")
                dc.StartPage()

                # Label size
                dpi = dc.GetDeviceCaps(88)
                label_width = int(2.625 * dpi)
                label_height = int(1.0 * dpi)
                x_center = label_width // 2
                y_start = 20

                if "pkt" in sku_suffix:
                    if not desc_line3:  # only 2 description lines
                        dc.SelectObject(bold_12)
                        dc.TextOut(x_center - dc.GetTextExtent(variety_name)[0] // 2, y_start, variety_name)
                        y_start += 55

                        dc.TextOut(x_center - dc.GetTextExtent(variety_crop)[0] // 2, y_start, variety_crop)
                        y_start += 58

                        dc.SelectObject(italic_9)
                        dc.TextOut(x_center - dc.GetTextExtent(desc_line1)[0] // 2, y_start, desc_line1)
                        y_start += 43

                        dc.TextOut(x_center - dc.GetTextExtent(desc_line2)[0] // 2, y_start, desc_line2)
                        y_start += 50

                        dc.SelectObject(normal_8)
                        dc.TextOut(x_center - dc.GetTextExtent(pkg_lot_germ)[0] // 2, y_start, pkg_lot_germ)
                        y_start += 40

                        dc.TextOut(x_center - dc.GetTextExtent(days_year)[0] // 2, y_start, days_year)
                    else:  # 3 description lines
                        dc.SelectObject(bold_12)
                        dc.TextOut(x_center - dc.GetTextExtent(variety_name)[0] // 2, y_start, variety_crop)
                        y_start += 55

                        dc.SelectObject(italic_9)
                        dc.TextOut(x_center - dc.GetTextExtent(desc_line1)[0] // 2, y_start, desc_line1)
                        y_start += 43

                        dc.TextOut(x_center - dc.GetTextExtent(desc_line2)[0] // 2, y_start, desc_line2)
                        y_start += 43

                        dc.TextOut(x_center - dc.GetTextExtent(desc_line3)[0] // 2, y_start, desc_line3)
                        y_start += 50

                        dc.SelectObject(normal_8)
                        dc.TextOut(x_center - dc.GetTextExtent(pkg_lot_germ)[0] // 2, y_start, pkg_lot_germ)
                        y_start += 40

                        dc.TextOut(x_center - dc.GetTextExtent(days_year)[0] // 2, y_start, days_year)
                else:
                    lot_germ = f"Lot: {lot_code}    Germ: {germination}%"
                    if not desc_line3:
                        if not rad_type:
                            dc.SelectObject(bold_16)
                            dc.TextOut(x_center - dc.GetTextExtent(variety_name)[0] // 2, y_start, variety_name)
                            y_start += 69

                            dc.SelectObject(normal_12)
                            dc.TextOut(x_center - dc.GetTextExtent(variety_crop)[0] // 2, y_start, variety_crop)
                            y_start += 54

                            dc.SelectObject(bold_12)
                            dc.TextOut(x_center - dc.GetTextExtent(pkg_size)[0] // 2, y_start, pkg_size)
                            y_start += 65

                            dc.SelectObject(normal_12)
                            dc.TextOut(x_center - dc.GetTextExtent(lot_germ)[0] // 2, y_start, lot_germ)
                            y_start += 48

                            dc.TextOut(x_center - dc.GetTextExtent(days_year)[0] // 2, y_start, days_year)
                        else:
                            pkg_days = f"{pkg_size} -- {days}"
                            lot_germ_year = f"Lot: {lot_code}    Germ: {germination}%    Packed for: {year}"

                            dc.SelectObject(bold_16)
                            dc.TextOut(x_center - dc.GetTextExtent(variety_name)[0] // 2, y_start, variety_name)
                            y_start += 64

                            dc.SelectObject(italic_12)
                            dc.TextOut(x_center - dc.GetTextExtent(rad_type)[0] // 2, y_start, rad_type)
                            y_start += 55

                            dc.SelectObject(normal_12)
                            dc.TextOut(x_center - dc.GetTextExtent(variety_crop)[0] // 2, y_start, variety_crop)
                            y_start += 50

                            dc.SelectObject(bold_12)
                            dc.TextOut(x_center - dc.GetTextExtent(pkg_days)[0] // 2, y_start, pkg_days)
                            y_start += 60

                            dc.SelectObject(normal_12)
                            dc.TextOut(x_center - dc.GetTextExtent(lot_germ_year)[0] // 2, y_start, lot_germ_year)
                    else:
                        dc.SelectObject(bold_16)
                        dc.TextOut(x_center - dc.GetTextExtent(variety_crop)[0] // 2, y_start, variety_crop)
                        y_start += 80

                        dc.SelectObject(bold_12)
                        dc.TextOut(x_center - dc.GetTextExtent(pkg_size)[0] // 2, y_start, pkg_size)
                        y_start += 75

                        dc.SelectObject(normal_12)
                        dc.TextOut(x_center - dc.GetTextExtent(lot_germ)[0] // 2, y_start, lot_germ)
                        y_start += 60

                        dc.TextOut(x_center - dc.GetTextExtent(days_year)[0] // 2, y_start, days_year)

                dc.EndPage()
                dc.EndDoc()
                dc.DeleteDC()

            return {'success': True, 'message': f'Front Single Label printed successfully ({quantity} copies)'}

    except Exception as e:
        print(f"Error printing front label: {str(e)}")
        return {'success': False, 'error': str(e)}


def print_single_back_label_logic(data):
    """Extract the core back label printing logic"""
    try:
        quantity = int(data.get('quantity', 1))
        env_multiplier = int(data.get('env_multiplier', 1))
        quantity *= env_multiplier

        if CURRENT_USER == "ndefe":
            print(f"Printing {quantity} back single labels on Ndefe's printer")
            print(f"Back1 {data.get('back1')}")
            print(f"Back2 {data.get('back2')}")
            print(f"Back3 {data.get('back3')}")
            print(f"Back4 {data.get('back4')}")
            print(f"Back5 {data.get('back5')}")
            print(f"Back6 {data.get('back6')}")
            print(f"Back7 {data.get('back7')}")
            return {'success': True, 'message': f'Back Single Label printed successfully ({quantity} copies)'}
        else:
            back_lines = [
                data.get('back1'),
                data.get('back2'),
                data.get('back3'),
                data.get('back4'),
                data.get('back5'),
                data.get('back6'),
                data.get('back7')
            ]

            # Remove empty lines (None or "")
            back_lines = [line for line in back_lines if line]

            if not back_lines:
                return {'success': False, 'message': 'No back lines provided'}

            # Printer setup
            printer_name = ROLL_PRINTER
            font = create_font("Book Antiqua", 32, italic=True)

            for i in range(quantity):
                dc = win32ui.CreateDC()
                dc.CreatePrinterDC(printer_name)

                dc.StartDoc("Seed Label")
                dc.StartPage()

                # Label size: 1" x 2.625" at 300 DPI
                dpi = dc.GetDeviceCaps(88)
                print(f"[DEBUG] Printer DPI: {dpi}")
                label_width = int(2.625 * dpi)
                label_height = int(1.0 * dpi)
                x_center = label_width // 2

                dc.SelectObject(font)

                # Spacing logic
                num_lines = len(back_lines)
                line_height = 39
                total_text_height = line_height * num_lines
                remaining_space = label_height - total_text_height
                y_start = (remaining_space // 2) + 12

                for line in back_lines:
                    text_width = dc.GetTextExtent(line)[0]
                    dc.TextOut(x_center - text_width // 2, y_start, line)
                    y_start += line_height

                dc.EndPage()
                dc.EndDoc()
                dc.DeleteDC()

            return {'success': True, 'message': f'Back Single Label printed successfully ({quantity} copies)'}

    except Exception as e:
        print(f"Error printing back label: {str(e)}")
        return {'success': False, 'error': str(e)}


@app.route('/print-single-front', methods=['POST'])
def print_single_front_label():
    """Route handler for single front label printing"""
    try:
        data = request.get_json()
        result = print_single_front_label_logic(data)
        
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 500
            
    except Exception as e:
        print(f"Error in front label route: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/print-single-back', methods=['POST'])
def print_single_back_label():
    """Route handler for single back label printing"""
    try:
        data = request.get_json()
        result = print_single_back_label_logic(data)
        
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 500
            
    except Exception as e:
        print(f"Error in back label route: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500



# [CONFIRMED WORKING]

def print_sheet_front_logic(data):
    """Extract the core front sheet printing logic"""
    try:
        quantity = int(data.get('quantity', 1))
        env_multiplier = int(data.get('env_multiplier', 1))
        print(f"Environmental Multiplier: {env_multiplier}")
        quantity *= env_multiplier

        if CURRENT_USER.lower() == "ndefe":
            print(f"Printing {quantity} front sheet labels on Ndefe's printer")
            print(f"Variety Name: {data.get('variety_name')}")
            print(f"Crop: {data.get('crop')}")
            print(f"Days: {data.get('days')}")
            print(f"SKU Suffix: {data.get('sku_suffix')}")
            print(f"Common name: {data.get('common_name')}")
            print(f"Pkg. size: {data.get('pkg_size')}")
            print(f"Env type: {data.get('env_type')}")
            print(f"Lot Code: {data.get('lot_code')}")
            print(f"Germination: {data.get('germination')}")
            print(f"For Year: {data.get('for_year')}")
            print(f"Quantity: {data.get('quantity')}")
            print(f"Desc1: {data.get('desc1')}")
            print(f"Desc2: {data.get('desc2')}")
            print(f"Desc3: {data.get('desc3')}")
            print(f"Rad type: {data.get('rad_type')}")
            print("================================")
            return {'success': True, 'message': f'Front Sheet Label printed successfully ({quantity} copies)'}
        else:
            # Gather label content (shared across copies) - same as single label logic
            # variety_name = f"'{data.get('variety_name')}'"
            variety_name = format_variety_name_with_quotes(data.get('variety_name'))
            
            # Check for common_name first, fall back to crop if empty
            common_name = data.get('common_name', '').strip()
            variety_crop = common_name if common_name else data.get('crop')
            
            days = data.get('days')
            env_type = data.get('env_type')
            year = data.get('for_year')
            days_year = f"{days}    Packed for 20{year}"

            desc_line1 = data.get('desc1')
            desc_line2 = data.get('desc2')
            desc_line3 = data.get('desc3')
            lot_code = data.get('lot_code')
            germination = data.get('germination')
            rad_type = data.get('rad_type')

            if env_type == "LG Coffee":
                pkg_size = f"{data.get('pkg_size')} ••"
            elif env_type == "SM Coffee":
                pkg_size = f"{data.get('pkg_size')} •"
            else:
                pkg_size = data.get('pkg_size')

            pkg_lot_germ = f"{pkg_size}    Lot: {lot_code}    Germ: {germination}%"
            sku_suffix = data.get('sku_suffix')

            # Fonts (doubled for 600 DPI vs 300 DPI)
            bold_12 = create_font("Times New Roman", 96, bold=True)     # 48 * 2
            italic_9 = create_font("Times New Roman", 72, italic=True)  # 36 * 2
            normal_8 = create_font("Times New Roman", 64)              # 32 * 2
            bold_16 = create_font("Times New Roman", 120, bold=True)   # 60 * 2
            normal_12 = create_font("Times New Roman", 80)            # 40 * 2
            italic_12 = create_font("Times New Roman", 80, italic=True) # 40 * 2

            printer_name = SHEET_PRINTER

            # Loop through each copy (same as single label approach)
            for i in range(quantity):
                dc = win32ui.CreateDC()
                dc.CreatePrinterDC(printer_name)

                dc.StartDoc("Seed Label Sheet")
                dc.StartPage()

                # Get printer DPI and calculate sheet dimensions
                dpi = dc.GetDeviceCaps(88)
                page_width = dc.GetDeviceCaps(8)
                page_height = dc.GetDeviceCaps(10)
                
                # Sheet layout: 3 columns x 10 rows = 30 labels
                margin_y = int(0.5 * dpi)  # 0.5 inch top margin
                label_width = page_width // 3
                # label_height = (page_height - margin_y) // 10 - 6
                label_height = int(1.05 * dpi)  # Exactly 1 inch per label (Avery 5960 spec)

                # Column adjustments for better alignment
                left_col_offset = -30
                middle_col_offset = 0
                right_col_offset = 30

                col_offsets = [left_col_offset, middle_col_offset, right_col_offset]

                # Draw 30 labels (3 columns x 10 rows)
                for row in range(10):
                    y_base = margin_y + (row * label_height)
                    
                    for col in range(3):
                        x_center = (col * label_width) + (label_width // 2) + col_offsets[col]
                        y_start = y_base - 70

                        # Use same conditional logic as single front label
                        if "pkt" in sku_suffix:
                            if not desc_line3:  # only 2 description lines
                                dc.SelectObject(bold_12)
                                dc.TextOut(x_center - dc.GetTextExtent(variety_name)[0] // 2, y_start, variety_name)
                                y_start += 110  # was 55

                                dc.TextOut(x_center - dc.GetTextExtent(variety_crop)[0] // 2, y_start, variety_crop)
                                y_start += 116  # was 58

                                dc.SelectObject(italic_9)
                                dc.TextOut(x_center - dc.GetTextExtent(desc_line1)[0] // 2, y_start, desc_line1)
                                y_start += 86  # was 43

                                dc.TextOut(x_center - dc.GetTextExtent(desc_line2)[0] // 2, y_start, desc_line2)
                                y_start += 100  # was 50

                                dc.SelectObject(normal_8)
                                dc.TextOut(x_center - dc.GetTextExtent(pkg_lot_germ)[0] // 2, y_start, pkg_lot_germ)
                                y_start += 80  # was 40

                                dc.TextOut(x_center - dc.GetTextExtent(days_year)[0] // 2, y_start, days_year)
                            else:  # 3 description lines
                                y_start += 10  # small adjustment for extra line
                                dc.SelectObject(bold_12)
                                dc.TextOut(x_center - dc.GetTextExtent(variety_crop)[0] // 2, y_start, variety_crop)
                                y_start += 110  # was 55

                                dc.SelectObject(italic_9)
                                dc.TextOut(x_center - dc.GetTextExtent(desc_line1)[0] // 2, y_start, desc_line1)
                                y_start += 86  # was 43

                                dc.TextOut(x_center - dc.GetTextExtent(desc_line2)[0] // 2, y_start, desc_line2)
                                y_start += 86  # was 43

                                dc.TextOut(x_center - dc.GetTextExtent(desc_line3)[0] // 2, y_start, desc_line3)
                                y_start += 100  # was 50

                                dc.SelectObject(normal_8)
                                dc.TextOut(x_center - dc.GetTextExtent(pkg_lot_germ)[0] // 2, y_start, pkg_lot_germ)
                                y_start += 80  # was 40

                                dc.TextOut(x_center - dc.GetTextExtent(days_year)[0] // 2, y_start, days_year)
                        else:
                            lot_germ = f"Lot: {lot_code}    Germ: {germination}%"
                            if not desc_line3:
                                if not rad_type:
                                    dc.SelectObject(bold_16)
                                    dc.TextOut(x_center - dc.GetTextExtent(variety_name)[0] // 2, y_start, variety_name)
                                    y_start += 138  # was 69

                                    dc.SelectObject(normal_12)
                                    dc.TextOut(x_center - dc.GetTextExtent(variety_crop)[0] // 2, y_start, variety_crop)
                                    y_start += 108  # was 54

                                    dc.SelectObject(bold_12)
                                    dc.TextOut(x_center - dc.GetTextExtent(pkg_size)[0] // 2, y_start, pkg_size)
                                    y_start += 130  # was 65

                                    dc.SelectObject(normal_12)
                                    dc.TextOut(x_center - dc.GetTextExtent(lot_germ)[0] // 2, y_start, lot_germ)
                                    y_start += 96  # was 48

                                    dc.TextOut(x_center - dc.GetTextExtent(days_year)[0] // 2, y_start, days_year)
                                else:
                                    pkg_days = f"{pkg_size} -- {days}"
                                    lot_germ_year = f"Lot: {lot_code}    Germ: {germination}%    Packed for: {year}"

                                    dc.SelectObject(bold_16)
                                    dc.TextOut(x_center - dc.GetTextExtent(variety_name)[0] // 2, y_start, variety_name)
                                    y_start += 128  # was 64

                                    dc.SelectObject(italic_12)
                                    dc.TextOut(x_center - dc.GetTextExtent(rad_type)[0] // 2, y_start, rad_type)
                                    y_start += 110  # was 55

                                    dc.SelectObject(normal_12)
                                    dc.TextOut(x_center - dc.GetTextExtent(variety_crop)[0] // 2, y_start, variety_crop)
                                    y_start += 100  # was 50

                                    dc.SelectObject(bold_12)
                                    dc.TextOut(x_center - dc.GetTextExtent(pkg_days)[0] // 2, y_start, pkg_days)
                                    y_start += 120  # was 60

                                    dc.SelectObject(normal_12)
                                    dc.TextOut(x_center - dc.GetTextExtent(lot_germ_year)[0] // 2, y_start, lot_germ_year)
                            else:
                                dc.SelectObject(bold_16)
                                dc.TextOut(x_center - dc.GetTextExtent(variety_crop)[0] // 2, y_start, variety_crop)
                                y_start += 160  # was 80

                                dc.SelectObject(bold_12)
                                dc.TextOut(x_center - dc.GetTextExtent(pkg_size)[0] // 2, y_start, pkg_size)
                                y_start += 150  # was 75

                                dc.SelectObject(normal_12)
                                dc.TextOut(x_center - dc.GetTextExtent(lot_germ)[0] // 2, y_start, lot_germ)
                                y_start += 120  # was 60

                                dc.TextOut(x_center - dc.GetTextExtent(days_year)[0] // 2, y_start, days_year)

                # Add envelope info at bottom of sheet
                
                envelope = f"Envelope: {env_type}"
                envelope_font = create_font("Times New Roman", 96, bold=True)  # Doubled from 48
                dc.SelectObject(envelope_font)
                envelope_x = int(0.5 * dpi)
                envelope_y = page_height - int(0.2 * dpi)
                dc.TextOut(envelope_x, envelope_y, envelope)

                dc.EndPage()
                dc.EndDoc()
                dc.DeleteDC()

            return {'success': True, 'message': f'Front Sheet Label printed successfully ({quantity} copies)'}

    except Exception as e:
        print(f"Error printing front sheet: {str(e)}")
        return {'success': False, 'error': str(e)}


def print_sheet_back_logic(data):
    """Extract the core back sheet printing logic"""
    try:
        quantity = int(data.get('quantity', 1))
        env_multiplier = int(data.get('env_multiplier', 1))
        print(f"Environmental Multiplier: {env_multiplier}")
        quantity *= env_multiplier
        variety_name = f"'{data.get('variety_name')}'"

        if CURRENT_USER.lower() == "ndefe":
            print(f"Printing {quantity} back sheet labels for {variety_name} on Ndefe's printer")
            print(f"Back1 {data.get('back1')}")
            print(f"Back2 {data.get('back2')}")
            print(f"Back3 {data.get('back3')}")
            print(f"Back4 {data.get('back4')}")
            print(f"Back5 {data.get('back5')}")
            print(f"Back6 {data.get('back6')}")
            print(f"Back7 {data.get('back7')}")
            print("================================")
            return {'success': True, 'message': f'Back Sheet Label printed successfully ({quantity} copies)'}
        else:
            # Gather back label content (same as single back label logic)
            back_lines = [
                data.get('back1'),
                data.get('back2'),
                data.get('back3'),
                data.get('back4'),
                data.get('back5'),
                data.get('back6'),
                data.get('back7')
            ]
            
            # Remove empty lines (same as single back label)
            back_lines = [line for line in back_lines if line]
            
            if not back_lines:
                return {'success': False, 'message': 'No back lines provided'}

            # Font (same as single back label)
            font = create_font("Book Antiqua", 66, italic=True)
            footer_font = create_font("Calibri", 80)

            printer_name = SHEET_PRINTER

            # Loop through each copy (same as single label approach)
            for i in range(quantity):
                dc = win32ui.CreateDC()
                dc.CreatePrinterDC(printer_name)

                dc.StartDoc("Seed Label Back Sheet")
                dc.StartPage()

                # Get printer DPI and calculate sheet dimensions
                dpi = dc.GetDeviceCaps(88)
                # print(f"Printer DPI: {dpi}")
                page_width = dc.GetDeviceCaps(8)
                page_height = dc.GetDeviceCaps(10)

                
                # Sheet layout: 3 columns x 10 rows = 30 labels
                margin_y = int(0.5 * dpi)
                label_width = page_width // 3
                label_height = (page_height - margin_y) // 10 - 7
                
                # Column adjustments for better alignment
                left_col_offset = -35
                middle_col_offset = 0
                right_col_offset = 35
                col_offsets = [left_col_offset, middle_col_offset, right_col_offset]

                dc.SelectObject(font)

                # Spacing logic (same as single back label)
                num_lines = len(back_lines)
                # if back line 7 is not present, increase line height to spread out
                if len(back_lines) < 7:
                    line_height = 90
                else:
                    line_height = 80  # Exact same as single back label
                total_text_height = line_height * num_lines

                # Draw 30 labels (3 columns x 10 rows)
                for row in range(10):
                    y_base = margin_y + (row * label_height)
                    
                    for col in range(3):
                        x_center = (col * label_width) + (label_width // 2) + col_offsets[col]
                        
                        # Calculate y_start (same logic as single back label)
                        remaining_space = label_height - total_text_height
                        y_start = y_base + (remaining_space // 2) - 80

                        # Draw each back line (same as single back label)
                        for line in back_lines:
                            text_width = dc.GetTextExtent(line)[0]
                            dc.TextOut(x_center - text_width // 2, y_start, line)
                            y_start += line_height

                # Footer with variety name
                dc.SelectObject(footer_font)
                footer_text = f"Variety: {variety_name}"
                footer_x = int(0.5 * dpi)
                footer_y = page_height - int(0.2 * dpi)
                dc.TextOut(footer_x, footer_y, footer_text)

                dc.EndPage()
                dc.EndDoc()
                dc.DeleteDC()

            return {'success': True, 'message': f'Back Sheet Label printed successfully ({quantity} copies)'}

    except Exception as e:
        print(f"Error printing back sheet: {str(e)}")
        return {'success': False, 'error': str(e)}


@app.route('/print-sheet-front', methods=['POST'])
def print_sheet_front():
    """Route handler for front sheet label printing"""
    try:
        data = request.get_json()
        result = print_sheet_front_logic(data)
       
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 500
           
    except Exception as e:
        print(f"Error in front sheet route: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/print-sheet-back', methods=['POST'])
def print_sheet_back():
    """Route handler for back sheet label printing"""
    try:
        data = request.get_json()
        result = print_sheet_back_logic(data)
       
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 500
           
    except Exception as e:
        print(f"Error in back sheet route: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/print-orders', methods=['POST'])
def print_orders():
    try:
        data = request.get_json()
        customer_orders = data.get('customer_orders')
        missing_orders = data.get('missing_orders')
        bulk_orders = data.get('bulk_orders')
        misc_orders = data.get('misc_orders')
        order_data = data.get('order_data')

        # look through customer_orders dict and extract those with duplicate orders
        duplicate_orders = {customer: orders for customer, orders in customer_orders.items() if len(orders) > 1}
        # print(f"Duplicate Orders: {duplicate_orders}")

        # Handle printing of duplicate orders
        handled_orders = set()

        for customer, orders in duplicate_orders.items():
            for order_number in orders:
                if order_number in order_data:
                    order = order_data[order_number]
                    print(f"Printing duplicate order {order_number} for customer {customer}")
                    generate_pdf(order_number, order, action="print")
                    handled_orders.add(order_number)

        # Now remove them in one go
        for order_number in handled_orders:
            order_data.pop(order_number, None)  # safe remove

        # Handle printing of packet-only orders
        pkt_only_orders = [
            order_number for order_number, order in order_data.items()
            if not order.get("bulk_items") and not order.get("misc_items")
        ]

        for order_number in pkt_only_orders:
            order = order_data[order_number]
            print(f"Printing packet-only order {order_number}")
            generate_pdf(order_number, order, action="print")

        # Remove them afterward
        for order_number in pkt_only_orders:
            order_data.pop(order_number, None) 

        # Print remaining orders
        for order_number, order in order_data.items():
            print(f"Printing bulk/misc order {order_number}")
            generate_pdf(order_number, order, action="print")

        return jsonify({
            'success': True,
            'message': f'Orders printed successfully',
            'multiple_order_customers': duplicate_orders
        })

    except Exception as e:
        print(f"Error printing orders: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500




@app.route('/print-items-to-pull', methods=['POST'])
def print_items_to_pull():
    try:
        data = request.get_json()
        items = data.get('items', [])
        batch_date = data.get('batch_date', 'Unknown')
        
        if not items:
            return jsonify({
                'success': False,
                'error': 'No items provided'
            }), 400
        
        # Create pdfs directory if it doesn't exist
        pdf_dir = 'packing_slips'
        os.makedirs(pdf_dir, exist_ok=True)
        
        # Generate filename with current date
        current_date = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'pull_{current_date}.pdf'
        file_path = os.path.join(pdf_dir, filename)
        
        # Create PDF
        create_pull_items_pdf(file_path, items, batch_date)
        
        # Print using Sumatra (skip if user is ndefe)
        if CURRENT_USER.lower() != "ndefe":
            try:
                command = f'"{SUMATRA_PATH}" -print-to "{SHEET_PRINTER}" -print-settings "fit,portrait" -silent "{file_path}"'
                subprocess.run(command, check=True, shell=True)
                print(f"Successfully printed {filename}")
            except Exception as e:
                print(f"Failed to print {filename}: {e}")
                return jsonify({
                    'success': False,
                    'error': f'Failed to print: {str(e)}'
                }), 500
            finally:
                # Clean up the file
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f"Temporary file {file_path} deleted.")
        else:
            print(f"PDF created at {file_path} (printing skipped for ndefe)")
        
        return jsonify({
            'success': True,
            'message': f'Successfully printed {len(items)} items to pull for batch {batch_date}'
        })
        
    except Exception as e:
        print(f"Error printing items to pull: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500



def create_pull_items_pdf(file_path, items, batch_date):
    """Create a PDF with items to pull table - black and white version"""
    doc = SimpleDocTemplate(file_path, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    
    # Centered title with date
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading2'],
        fontSize=14,
        spaceAfter=20,
        alignment=1  # Center alignment
    )
    title = Paragraph(f"Bulk items to pull -- {batch_date}", title_style)
    story.append(title)
    
    # Table data
    table_data = [
        ['#', 'Variety Name', 'Type', 'SKU Suffix', 'Qty']
    ]
    
    for i, item in enumerate(items, 1):
        table_data.append([
            str(i),
            item.get('variety_name', ''),
            item.get('type', ''),
            item.get('sku_suffix', ''),
            str(item.get('quantity', 0))
        ])
    
    # Create table
    table = Table(table_data, colWidths=[0.5*inch, 3*inch, 1.5*inch, 1*inch, 0.7*inch])
    
    # Black and white table style
    table.setStyle(TableStyle([
        # Header row - bold and larger font
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        
        # Data rows - regular font
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        
        # Alignment
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        
        # Grid lines
        ('GRID', (0, 0), (-1, -1), 1, 'black'),
        
        # Padding
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    
    story.append(table)
    
    # Footer
    story.append(Spacer(1, 20))
    footer_text = f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    footer = Paragraph(footer_text, styles['Normal'])
    story.append(footer)
    
    doc.build(story)
    print(f"PDF created: {file_path}")



@app.route('/generate-packing-slip', methods=['POST'])
def generate_packing_slip():
    if request.method == 'OPTIONS':
        return '', 200
    try:
        # print("Generating packing slip PDF in flask...")
        data = request.get_json()
        order = data.get('order')
        
        if not order:
            return jsonify({'success': False, 'error': 'No order data provided'})
        
        order_number = order.get('order_number', 'unknown')
        
        # Call your existing generate_pdf function with action="view"
        generate_pdf(order_number, order, action="view")
        
        # Return success response to the browser
        return jsonify({
            'success': True, 
            'message': f'Packing slip for order {order_number} opened locally'
        })
       
    except Exception as e:
        print(f"Error generating packing slip PDF: {e}")
        return jsonify({'success': False, 'error': str(e)})



@app.route('/reprocess-order', methods=['POST', 'OPTIONS'])
def reprocess_order():
    if request.method == 'OPTIONS':
        return '', 200
    
    try:
        print("Reprocessing order in Flask...")
        data = request.get_json()
        order = data.get('order')
        bulk_to_print = data.get('bulk_to_print', {})
        
    
        if not order:
            return jsonify({'success': False, 'error': 'No order data provided'})
        
        order_number = order.get('order_number', 'unknown')
        
        # Generate and print the packing slip
        generate_pdf(order_number, order, action="print")
        

        # TODO: Process bulk_to_print items here if needed
        # by calling print_single_front_label_logic, and print_single_back_label_logic (if back lines exist)
        for sku, item in bulk_to_print.items():
            print_single_front_label_logic(item)
            if item.get('back1'):  # Assuming back1 is mandatory for back label
                print_single_back_label_logic(item)
        
        
        return jsonify({
            'success': True,
            'message': f'Order {order_number} reprocessed and sent to printer'
        })
       
    except Exception as e:
        print(f"Error reprocessing order: {e}")
        return jsonify({'success': False, 'error': str(e)})
    

# GENERATES PACKING SLIPS
def generate_pdf(order_number, order, action):
    # from process_orders import separate_pkts_and_bulk, sort_lineitems
    filename = f"{order_number}.pdf"
    file_path = f"packing_slips/{filename}"    

    sorted_misc_list = order.get("misc_items", [])
    sorted_bulk_list = order.get("bulk_items", [])
    sorted_pkt_list  = order.get("pkt_items", [])

    customer_name = order.get('customer_name')
    address = order.get('address')
    address2 = order.get('address2') or ""
    postal_code = order.get('postal_code')
    city = order.get('city')
    state = order.get('state')
    country = order.get('country')
    note = order.get('note') or ""
    shipping = order.get('shipping', 0)
    tax = order.get('tax', 0)
    subtotal = order.get('subtotal', 0)
    total = order.get('total', 0)
   
    order_date = order.get('date')
    order_dt = datetime.fromisoformat(order_date)
    order_date = order_dt.strftime("%m/%d/%Y")

    # packet only
    if not sorted_bulk_list and not sorted_misc_list:
        num_items = len(sorted_pkt_list)
    # misc only
    elif not sorted_pkt_list and not sorted_bulk_list:
        num_items = len(sorted_misc_list)
    # bulk only
    elif not sorted_pkt_list and not sorted_misc_list:
        num_items = len(sorted_bulk_list)
    # packets and bulk
    elif not sorted_misc_list:
        num_items = len(sorted_pkt_list) + len(sorted_bulk_list) + 1
    # packets and misc
    elif not sorted_bulk_list:
        num_items = len(sorted_pkt_list) + len(sorted_misc_list) + 1
    # bulk and misc
    elif not sorted_pkt_list:
        num_items = len(sorted_bulk_list) + len(sorted_misc_list) + 1   
    # packets, bulk, and misc
    elif sorted_pkt_list and sorted_bulk_list and sorted_misc_list:
        num_items = len(sorted_pkt_list) + len(sorted_bulk_list) + len(sorted_misc_list) + 2
    
    if num_items <= 27:
        num_pages = 1
    elif num_items <=  70:
        num_pages = 2
    elif num_items <=  113:  
        num_pages = 3
    elif num_items <=  156:
        num_pages = 4
    elif num_items <=  199:
        num_pages = 5
    else:
         num_pages = 6
    
    c = canvas.Canvas(file_path, pagesize=LETTER)
    width, height = LETTER

    # Add logo
    logo_width = 100
    logo_height = 50
    # Position in the upper-right corner
    logo_x = width - logo_width - 40  # 50 px from the right margin
    logo_y = height - logo_height - 35  # 30 px from the top margin
    # Draw the image
    c.drawImage(LOGO_PATH, logo_x, logo_y, width=logo_width, height=logo_height, mask='auto')
    # end logo

    c.setFont("Calibri-Bold", 10)
    c.drawString(460, height - 100, f"100% USDA Certified Organic")
    
    c.setFont("Calibri", 12)

    def draw_header(c, page_num):
        # Last 3 digits of order number
        last_digits = order_number[-3:]
        # Customer last name in uppercase
        try:
            last_name = customer_name.split()[-1].upper()
        except:
            last_name = customer_name.upper()
        # Total pages
        page_info = f"PAGE {page_num} OF {num_pages}"

        # Set font
        c.setFont("Calibri", 14)

        # Define positions for each section (adjust as needed)
        left_x = 30  # Left-aligned position
        center_x = width / 2  # Center of the page
        right_x = width - 100  # Right-aligned position
    
        # Draw each section separately
        c.drawString(left_x, height - 25, f"{last_name} - {last_digits}")  # Left
        c.drawCentredString(center_x, height - 25, "PACKING SLIP")  # Centered
        c.drawString(right_x, height - 25, page_info) 
        c.line(0, height - 30, width - 0, height - 30)

    draw_header(c, 1)

    c.setFont("Calibri", 10)

    # Draw the note if it exists
    if note:
        note = f"Note: {note}"
        # text wrap
        wrapped_note = textwrap.wrap(note, width=55)
        y = height - 120  # Starting y-position
        i = 0
        for line in wrapped_note:
            if i <= 4:
                c.drawString(335, y, line)
                y -= 14
            i += 1 
        print(f"order number {order_number} has a note")

    c.setFont("Calibri-Bold", 14)
    c.drawString(50, height - 60, "Uprising Seeds")
    c.setFont("Calibri", 12)
    c.drawString(50, height - 75, "1501 Fraser St")
    c.drawString(50, height - 90, "Suite 105")
    c.drawString(50, height - 105, "Bellingham, WA 98229")
    c.drawString(50, height - 120, "360-778-3749")
    c.drawString(50, height - 135, "info@uprisingorganics.com")
    
    # # if canadian order in italics
    if country == "CA":
        def draw_centered_text(c, text, x, y, font="Calibri-Italic", font_size=10):
            c.setFont(font, font_size)
            text_width = c.stringWidth(text, font, font_size)
            centered_x = x - (text_width / 2)  # Center the text based on the X coordinate
            c.drawString(centered_x, y, text)

        # c.setFont("Calibri-Italic", 12)
        y = height - 54  # Starting Y position
        draw_centered_text(c, "Certified in compliance with", 310, y)
        y -= 15  # Adjust Y position for the next line
        draw_centered_text(c, "the terms of the US-Canada", 310, y)
        y -= 15
        draw_centered_text(c, "Organic Equivalency Arrangement", 310, y)

    c.line(50, height - 150, width - 300, height - 150)
    c.setFont("Calibri-Bold", 12)
    c.drawString(60, height - 164, "SHIP TO:")
    
    c.line(50, height - 150, 50, height - 280)
    c.line(width - 300, height - 150, width - 300, height - 280)

    c.line(50, height - 170, width - 300, height - 170)
    
    c.setFont("Calibri", 12)

    # Function to right-align text at x = 120
    def draw_right_aligned(c, text, y):
        text_width = c.stringWidth(text, 'Calibri', 12)
        c.drawString(130 - text_width, y, text)

    # Customer info
    draw_right_aligned(c, "Order #:", height - 185)
    draw_right_aligned(c, "Name:", height - 200)
    draw_right_aligned(c, "Date:", height - 215)
    c.drawString(140, height - 215, order_date)  
    draw_right_aligned(c, "Address:", height - 230)
    c.drawString(140, height - 230, address)

    if address2:
        draw_right_aligned(c, "Address 2:", height - 245)
        c.drawString(140, height - 245, address2)
        draw_right_aligned(c, "City/State/Zip:", height - 260)
        c.drawString(140, height - 260, f"{city}, {state}   {postal_code}")
        draw_right_aligned(c, "Country:", height - 275)
        c.drawString(140, height - 275, country)
    else:
        draw_right_aligned(c, "City/State/Zip:", height - 245)
        c.drawString(140, height - 245, f"{city}, {state}   {postal_code}")
        draw_right_aligned(c, "Country:", height - 260)
        c.drawString(140, height - 260, country)

    # Draw customer info
    c.setFont("Calibri-Bold", 12)
    c.drawString(140, height - 185, order_number)
    c.drawString(140, height - 200, customer_name)

    # c.drawString(140, height - 215, order.order_number)
    c.line(50, height - 280, width - 300, height - 280)

    # Define column X positions
    qty_x = 65  # Centered quantity
    product_x = 90  # Left-aligned description
    price_x = 450  # Price column
    ext_price_x = 550  # Extended price column

    # Draw header lines
    c.line(0, height - 290, width, height - 290)  # Top header line

    # Draw column headers
    c.drawCentredString(qty_x, height - 305, "QTY")  # Centered over quantity
    c.drawString(product_x, height - 305, "Description")  # Left-aligned for product
    c.drawRightString(price_x, height - 305, "Price")  # Right-aligned price
    c.drawRightString(555, height - 305, "Ext. Price")  # Right-aligned extended price

    # Bottom header line
    c.line(0, height - 310, width, height - 310)  

    # Ensure shipping, tax, and total are treated as floats
    def format_currency(value):
        try:
            return f"${float(value):.2f}" 
        except ValueError:
            return "$0.00"

    # Calculate right-aligned positions for Shipping, Tax, Total
    right_x = 550 # Starting point for the right-aligned numbers (near the right edge)
    label_x = 435 
    # Function to draw right-aligned numbers
    def draw_right_aligned_label_value(label, value, y_position):
        c.drawString(label_x, y_position, label)
        value_str = format_currency(value)  # Format the value as currency
        value_width = c.stringWidth(value_str, "Calibri", 12)
        c.drawString(right_x - value_width, y_position, value_str)

    c.setFont("Calibri", 12)
    # Draw Shipping, Tax, and Total right-aligned
    draw_right_aligned_label_value("Shipping:", shipping, height - 225)
    draw_right_aligned_label_value("Tax:", tax, height - 240)
    draw_right_aligned_label_value("Subtotal:", subtotal, height - 255)
    c.setFont("Calibri-Bold", 12)
    draw_right_aligned_label_value("Total:", total, height - 270)

    # box in the order summary
    c.drawString(452, height - 205, "Order Summary")
    c.line(428, height - 190, 428, height - 278)
    c.line(558, height - 190, 558, height - 278)
    c.line(428, height - 190, 558, height - 190)
    c.line(428, height - 210, 558, height - 210)
    c.line(428, height - 278, 558, height - 278)
   
    def draw_lineitem(c, lineitem, lineitem_height, counter):

        qty = str(lineitem['qty'])  # Convert qty to string for centering
        lineitem_name = lineitem['lineitem']
        price = lineitem['price']

        ext_price = f"${float(qty) * float(price):.2f}"
        price = f"${float(price):.2f}"
        qty_x = 65  # Centered quantity
        product_x = 90  # Left-aligned description
        price_x = 450  # Price column
        ext_price_x = 550  # Extended price column

        line_y = height - lineitem_height

        # Highlight qty background in gray if qty > 1
        if int(qty) > 1:
            c.setFillGray(0.9)  # Light gray fill
            c.setStrokeColorRGB(0.9, 0.9, 0.9)  # Match stroke to fill
            c.rect(55, line_y - 4, 20, 17, fill=1, stroke=0)  # stroke=0 removes border
            c.setFillColorRGB(0, 0, 0)  # Reset text color
            c.setStrokeColorRGB(0, 0, 0)  

        # Draw line item data
        c.drawString(30, line_y, "___")
        c.drawCentredString(qty_x, line_y, qty)  # Center qty
        c.drawString(product_x, line_y, lineitem_name)  # Left-align description
        c.drawRightString(price_x, line_y, price)  # Right-align price from OOIncludes
        c.drawRightString(ext_price_x, line_y, ext_price)  # Right-align extended price

        # Move to next line
        lineitem_height += 17
        counter += 1
        return lineitem_height, counter


    # def draw_misc_items(c, sorted_misc, lineitem_height, counter):
    def draw_lineitems(c, item_list, lineitem_height, counter):
        if counter != 1:
            lineitem_height += 17  # extra space between sections
        # print(f"New draw_lineitems: {item_list}")
        for lineitem in item_list:
            if counter < 28:
                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)
            elif counter == 28:
                c.showPage()
                draw_header(c, 2)

                c.setFont("Calibri", 11)
                lineitem_height = 47

                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)
            elif counter < 71:
                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)
            elif counter == 71:
                c.showPage()
                draw_header(c, 3)

                c.setFont("Calibri", 11)
                lineitem_height = 47

                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)
            elif counter < 114:
                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)

            elif counter == 114:
                c.showPage()
                draw_header(c, 4)

                c.setFont("Calibri", 11)
                lineitem_height = 47

                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)

            elif counter < 157:
                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)
            elif counter == 157:
                c.showPage()
                draw_header(c, 5)

                c.setFont("Calibri", 11)
                lineitem_height = 47

                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)

            elif counter < 200:
                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)
            elif counter == 200:
                c.showPage()
                draw_header(c, 6)

                c.setFont("Calibri", 11)
                lineitem_height = 47

                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)

            elif counter == 200:
                c.showPage()
                draw_header(c, 6)

                c.setFont("Calibri", 11)
                lineitem_height = 47

                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)
            elif counter < 243:
                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)
            elif counter == 243:
                c.showPage()
                draw_header(c, 7)

                c.setFont("Calibri", 11)
                lineitem_height = 47

                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)

            elif counter < 285:
                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)
            elif counter == 285:
                c.showPage()
                draw_header(c, 8)

                c.setFont("Calibri", 11)
                lineitem_height = 47

                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)
            elif counter == 285:
                c.showPage()
                draw_header(c, 8)

                c.setFont("Calibri", 11)
                lineitem_height = 47

                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)
            elif counter < 285:
                lineitem_height, counter = draw_lineitem(c, lineitem, lineitem_height, counter)

        return lineitem_height, counter
    
    # make font smaller and not bold
    c.setFont("Calibri", 11)
    lineitem_height = 325
    counter = 1

    if len(sorted_pkt_list) > 0:
        # packet only
        if len(sorted_bulk_list) < 1 and len(sorted_misc_list) < 1:
            lineitem_height, counter = draw_lineitems(c, sorted_pkt_list, lineitem_height, counter)
        # packets and bulk
        elif len(sorted_bulk_list) > 0 and len(sorted_misc_list) < 1:
            lineitem_height, counter = draw_lineitems(c, sorted_bulk_list, lineitem_height, counter)
            lineitem_height, counter = draw_lineitems(c, sorted_pkt_list, lineitem_height, counter)
        # packets and misc
        elif len(sorted_bulk_list) < 1 and len(sorted_misc_list) > 0:
            lineitem_height, counter = draw_lineitems(c, sorted_misc_list, lineitem_height, counter)
            lineitem_height, counter = draw_lineitems(c, sorted_bulk_list, lineitem_height, counter)
        # pkts, bulks, and misc
        elif len(sorted_bulk_list) > 0 and len(sorted_misc_list) > 0:
            lineitem_height, counter = draw_lineitems(c, sorted_misc_list, lineitem_height, counter) 
            lineitem_height, counter = draw_lineitems(c, sorted_bulk_list, lineitem_height, counter)
            lineitem_height, counter = draw_lineitems(c, sorted_pkt_list, lineitem_height, counter)

    elif len(sorted_bulk_list) > 0:
        if len(sorted_misc_list) < 1:
            lineitem_height, counter = draw_lineitems(c, sorted_bulk_list, lineitem_height, counter)
        else:
            lineitem_height, counter = draw_lineitems(c, sorted_misc_list, lineitem_height, counter)
            lineitem_height, counter = draw_lineitems(c, sorted_bulk_list, lineitem_height, counter)
    elif len(sorted_misc_list) > 0:
        lineitem_height, counter = draw_lineitems(c, sorted_misc_list, lineitem_height, counter)

    c.save()
    
    if action == "print":
        if CURRENT_USER.lower() != "ndefe":
            try:
                command = f'"{SUMATRA_PATH}" -print-to "{SHEET_PRINTER}" -print-settings "fit,portrait" -silent "{file_path}"'
                subprocess.run(command, check=True, shell=True)

            except Exception as e:
                print(f"Failed to print: {e}")
            finally:
                # Clean up the file
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f"Temporary file {file_path} deleted.")

    elif action == "view":
    
        file_path = os.path.abspath(f"packing_slips/{order['order_number']}.pdf")  # Fixed
        os.startfile(file_path)  # This works on Windows only

    return

    

# @app.route('/view-invoice', methods=['POST'])
# def view_invoice():
#     order_number = request.form.get("order_data")
#     # call generate_pdf with action "view"
#     # generate_pdf(order_number, order_data[order_number], action="view")
#     return

# Handles printing bulk items from the process order page
@app.route('/print-range', methods=['POST'])
def print_range():

    items_missing_data = []

    try:

        data = request.get_json()
        items = data.get("items")
        current_order_year = data.get("current_order_year")
        
        if not items:
            return jsonify({
                'success': False,
                'error': 'No items provided'
            }), 400
        
        total_printed = 0
        
        for item in items:
            quantity = int(item.get('quantity', 1))
            env_multiplier = int(item.get('env_multiplier', 1))
            quantity *= env_multiplier
            sku = item.get('sku', '')
            
            # Extract sku_suffix from full sku (everything after the dash)
            sku_parts = sku.split('-')
            sku_suffix = sku_parts[-1] if len(sku_parts) > 1 else ''
            
            lot_code = item.get('lot', '')
            germination = item.get('germination', '')
            for_year = item.get('for_year', '')
            if not lot_code or not germination or not for_year:
                print(f"Item {sku} is missing lot, germination, or for_year")
                items_missing_data.append(sku)
                continue  # Skip this item and move to the next


            try:
                for_year_int = int(for_year)
                current_year_int = int(current_order_year) if current_order_year else 0
                
                if for_year_int < current_year_int:
                    print(f"Item {sku} germination for_year ({for_year_int}) is less than current_order_year ({current_year_int})")
                    items_missing_data.append(sku)
                    continue  # Skip this item
                    
            except (ValueError, TypeError):
                print(f"Item {sku} has invalid for_year ({for_year}) or current_order_year ({current_order_year})")
                items_missing_data.append(sku)
                continue  # Skip this item


            # Prepare data for printing functions
            print_data = {
                'variety_name': item.get('variety_name'),
                'crop': item.get('crop'),
                'days': item.get('days'),
                'sku_suffix': sku_suffix,
                'pkg_size': item.get('pkg_size'),
                'env_type': item.get('env_type'),
                'lot_code': item.get('lot', 'N/A'),
                'germination': item.get('germination', 'N/A'), 
                'for_year': item.get('for_year', 'N/A'),  
                'quantity': quantity,
                'desc1': item.get('desc1'),
                'desc2': item.get('desc2'),
                'desc3': item.get('desc3'),
                'rad_type': item.get('rad_type')
            }
   
            # Print back labels first if needed
            if item.get('print_back', False):
                back_data = {
                    'quantity': quantity,
                    'back1': item.get('back1'),
                    'back2': item.get('back2'),
                    'back3': item.get('back3'),
                    'back4': item.get('back4'),
                    'back5': item.get('back5'),
                    'back6': item.get('back6'),
                    'back7': item.get('back7')
                }
                
                # Call the back label printing logic directly
                result = print_single_back_label_logic(back_data)
                if not result.get('success'):
                    return jsonify(result), 500
            
            # Print front labels
            result = print_single_front_label_logic(print_data)
            if not result.get('success'):
                return jsonify(result), 500
                
            total_printed += quantity
        
        # ONLY CHANGE: Include items_missing_data in response
        response = {
            'success': True,
            'message': f'Printed {total_printed} bulk labels successfully',
        }
        
        if items_missing_data:
            response['items_missing_data'] = items_missing_data
            
        return jsonify(response)
        
    except Exception as e:
        print(f"Error printing bulk range: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500











@app.route('/print-envelope-table', methods=['POST'])
def print_envelope_table():
    """
    Receive envelope data and either print to console or create PDF and print
    """
    try:
        # Get JSON data from request
        data = request.get_json()
        
        if not data:
            return jsonify({
                'success': False,
                'error': 'No data received'
            }), 400
        
        print(f"Received envelope data for printing from user: {CURRENT_USER}")
        print(f"Data keys: {list(data.keys())}")
        
        # Extract envelope data
        envelope_data_by_year = data.get('envelope_data_by_year', {})
        years = data.get('years', [])
        grand_total = data.get('grand_total', 0)
        envelope_types = data.get('envelope_types', [])
        report_title = data.get('report_title', 'Envelope Usage Report')
        
        if not envelope_data_by_year:
            return jsonify({
                'success': False,
                'error': 'No envelope data found'
            }), 400
        
        # Check current user and handle accordingly
        if CURRENT_USER.lower() == "ndefe":
            print_console_table(envelope_data_by_year, years, grand_total, report_title)
            return jsonify({
                'success': True,
                'message': 'Table printed to console',
                'user': CURRENT_USER
            }), 200
        else:
            return create_and_print_pdf(envelope_data_by_year, years, grand_total, envelope_types, report_title)
            
    except Exception as e:
        print(f"Error in print_envelope_table: {str(e)}")
        return jsonify({
            'success': False,
            'error': f'Server error: {str(e)}'
        }), 500

def print_console_table(envelope_data_by_year, years, grand_total, report_title):
    """
    Print a nicely formatted table to the console for ndefe user
    """
    print("\n" + "="*80)
    print(f"{report_title:^80}")
    print("="*80)
    print(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Grand Total Envelopes: {grand_total:,}")
    print("="*80)
    
    # Get all unique envelope types across all years
    all_envelope_types = set()
    for year_data in envelope_data_by_year.values():
        all_envelope_types.update(year_data['envelope_counts'].keys())
    all_envelope_types = sorted(all_envelope_types)
    
    # Create header
    header = ["Envelope Type"] + [str(year) for year in sorted(years)] + ["Total"]
    
    # Calculate column widths
    max_type_width = max(len("Envelope Type"), max(len(env_type) for env_type in all_envelope_types) if all_envelope_types else 0)
    year_width = 10
    total_width = 12
    
    # Print header
    print(f"{'Envelope Type':<{max_type_width}} ", end="")
    for year in sorted(years):
        print(f"{str(year):>{year_width}} ", end="")
    print(f"{'Total':>{total_width}}")
    
    # Print separator
    print("-" * (max_type_width + len(years) * (year_width + 1) + total_width + 1))
    
    # Print data rows
    for env_type in all_envelope_types:
        row_total = 0
        print(f"{env_type:<{max_type_width}} ", end="")
        
        for year in sorted(years):
            year_str = str(year)
            count = envelope_data_by_year.get(year_str, {}).get('envelope_counts', {}).get(env_type, 0)
            row_total += count
            print(f"{count:>{year_width},} ", end="")
        
        print(f"{row_total:>{total_width},}")
    
    # Print totals row
    print("-" * (max_type_width + len(years) * (year_width + 1) + total_width + 1))
    print(f"{'TOTAL':<{max_type_width}} ", end="")
    
    for year in sorted(years):
        year_str = str(year)
        year_total = envelope_data_by_year.get(year_str, {}).get('total', 0)
        print(f"{year_total:>{year_width},} ", end="")
    
    print(f"{grand_total:>{total_width},}")
    print("="*80)
    print()

def create_and_print_pdf(envelope_data_by_year, years, grand_total, envelope_types, report_title):
    """
    Create a PDF report and print it
    """
    # Create temporary file
    temp_file = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    file_path = temp_file.name
    temp_file.close()
    
    try:
        # Create PDF document
        doc = SimpleDocTemplate(file_path, pagesize=letter,
                              rightMargin=72, leftMargin=72,
                              topMargin=72, bottomMargin=18)
        
        # Container for the 'Flowable' objects
        elements = []
        
        # Styles
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            spaceAfter=30,
            alignment=TA_CENTER
        )
        
        # Title
        title = Paragraph(report_title, title_style)
        elements.append(title)
        
        # Subtitle with generation info
        subtitle_text = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Grand Total: {grand_total:,} Envelopes"
        subtitle = Paragraph(subtitle_text, styles['Normal'])
        elements.append(subtitle)
        elements.append(Spacer(1, 20))
        
        # Prepare table data
        sorted_years = sorted(years)
        sorted_envelope_types = sorted(envelope_types)
        
        # Table header
        table_data = [["Envelope Type"] + [str(year) for year in sorted_years] + ["Total"]]
        
        # Table rows
        for env_type in sorted_envelope_types:
            row = [env_type]
            row_total = 0
            
            for year in sorted_years:
                year_str = str(year)
                count = envelope_data_by_year.get(year_str, {}).get('envelope_counts', {}).get(env_type, 0)
                row.append(f"{count:,}")
                row_total += count
            
            row.append(f"{row_total:,}")
            table_data.append(row)
        
        # Totals row
        totals_row = ["TOTAL"]
        for year in sorted_years:
            year_str = str(year)
            year_total = envelope_data_by_year.get(year_str, {}).get('total', 0)
            totals_row.append(f"{year_total:,}")
        totals_row.append(f"{grand_total:,}")
        table_data.append(totals_row)
        
        # Create table
        table = Table(table_data)
        
        # Table styling
        table_style = TableStyle([
            # Header row
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            
            # Data rows
            ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -2), 9),
            ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),  # Right align numbers
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),   # Left align envelope types
            
            # Totals row
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, -1), (-1, -1), 10),
            
            # Grid
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            
            # Alternating row colors (excluding header and total)
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.beige, colors.white])
        ])
        
        table.setStyle(table_style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        print(f"PDF created successfully: {file_path}")
        
        # Print the PDF
        if CURRENT_USER.lower() != "ndefe":
            try:
                command = f'"{SUMATRA_PATH}" -print-to "{SHEET_PRINTER}" -print-settings "fit,portrait" -silent "{file_path}"'
                subprocess.run(command, check=True, shell=True)
                print(f"Successfully printed envelope report")
                
                return jsonify({
                    'success': True,
                    'message': 'PDF created and sent to printer successfully',
                    'filename': os.path.basename(file_path)
                }), 200
                
            except subprocess.CalledProcessError as e:
                print(f"Failed to print envelope report: {e}")
                return jsonify({
                    'success': False,
                    'error': f'Failed to print: {str(e)}'
                }), 500
            except Exception as e:
                print(f"Failed to print envelope report: {e}")
                return jsonify({
                    'success': False,
                    'error': f'Failed to print: {str(e)}'
                }), 500
        
    except Exception as e:
        print(f"Error creating PDF: {str(e)}")
        return jsonify({
            'success': False,
            'error': f'Failed to create PDF: {str(e)}'
        }), 500
    
    finally:
        # Clean up the temporary file
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                print(f"Temporary file {file_path} deleted.")
            except Exception as e:
                print(f"Warning: Could not delete temporary file {file_path}: {e}")




@app.route('/print-address-labels', methods=['POST'])
def print_address_labels():
    """
    Print address labels PDF
    """
    try:
        print(f"Address labels print request from user: {CURRENT_USER}")
        
        # Check current user and handle accordingly
        if CURRENT_USER.lower() == "ndefe":
            print("Printing address labels")
            return jsonify({
                'success': True,
                'message': 'Address labels printed to console',
                'user': CURRENT_USER
            }), 200
        else:
            # Print the PDF using Sumatra
            pdf_path = "assets/address_labels.pdf"
            
            # Check if PDF exists
            if not os.path.exists(pdf_path):
                return jsonify({
                    'success': False,
                    'error': f'Address labels PDF not found at {pdf_path}'
                }), 404
            
            try:
                command = f'"{SUMATRA_PATH}" -print-to "{SHEET_PRINTER}" -print-settings "fit,portrait" -silent "{pdf_path}"'
                subprocess.run(command, check=True, shell=True)
                print(f"Successfully printed address labels from {pdf_path}")
                
                return jsonify({
                    'success': True,
                    'message': 'Address labels sent to printer successfully',
                    'user': CURRENT_USER
                }), 200
                
            except subprocess.CalledProcessError as e:
                print(f"Failed to print address labels: {e}")
                return jsonify({
                    'success': False,
                    'error': f'Failed to print: {str(e)}'
                }), 500
            except Exception as e:
                print(f"Failed to print address labels: {e}")
                return jsonify({
                    'success': False,
                    'error': f'Print error: {str(e)}'
                }), 500
            
    except Exception as e:
        print(f"Error in print_address_labels: {str(e)}")
        return jsonify({
            'success': False,
            'error': f'Server error: {str(e)}'
        }), 500



@app.route('/print-stock-seed-label', methods=['POST'])
def print_stock_seed_label():
    """
    Handle stock seed label printing requests
    """
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        # Extract the data
        variety = data.get('variety', 'Unknown')
        veg_type = data.get('veg_type', 'Unknown') 
        lot_number = data.get('lot_number', 'Unknown')
        quantity = data.get('quantity', 'Unknown')
        
        if CURRENT_USER.lower() == "ndefe":
            # Print to console for ndefe user
            print("\n" + "="*50)
            print("STOCK SEED LABEL PRINT REQUEST")
            print("="*50)
            print(f"Variety: {variety}")
            print(f"Vegetable Type: {veg_type}")
            print(f"Lot Number: {lot_number}")
            print(f"Quantity Saved: {quantity}")
            print("="*50 + "\n")
            
            # Log the full data for debugging
            print(f"Full data received: {data}")
            
            return jsonify({
                'success': True,
                'message': 'Stock seed label data received and printed to console'
            })
        else:
            # Print actual label for other users
            try:
                # Format the variety name with single quotes
                variety_formatted = f"'{variety}'"
                
                # Setup printer
                printer_name = ROLL_PRINTER
                dc = win32ui.CreateDC()
                dc.CreatePrinterDC(printer_name)
                
                dc.StartDoc("Stock Seed Label")
                dc.StartPage()
                
                # Label dimensions - same as your other labels
                dpi = dc.GetDeviceCaps(88)
                label_width = int(2.625 * dpi)
                label_height = int(1.0 * dpi)
                x_center = label_width // 2
                
                # Create bold font at size 54
                bold_font = create_font("Times New Roman", 54, bold=True)
                dc.SelectObject(bold_font)
                
                # Starting Y position
                y_start = 20
                line_height = 70  # Increased spacing between rows
                
                # Line 1: "* STOCK SEED *"
                header_text = "* STOCK SEED *"
                text_width = dc.GetTextExtent(header_text)[0]
                dc.TextOut(x_center - text_width // 2, y_start, header_text)
                y_start += line_height
                
                # Line 2: Variety name in single quotes
                text_width = dc.GetTextExtent(variety_formatted)[0]
                dc.TextOut(x_center - text_width // 2, y_start, variety_formatted)
                y_start += line_height
                
                # Line 3: Vegetable type
                text_width = dc.GetTextExtent(veg_type)[0]
                dc.TextOut(x_center - text_width // 2, y_start, veg_type)
                y_start += line_height
                
                # Line 4: "Lot: " + lot number
                lot_text = f"Lot: {lot_number}"
                text_width = dc.GetTextExtent(lot_text)[0]
                dc.TextOut(x_center - text_width // 2, y_start, lot_text)
                
                # Finalize print job
                dc.EndPage()
                dc.EndDoc()
                dc.DeleteDC()
                
                return jsonify({
                    'success': True,
                    'message': 'Stock seed label printed successfully'
                })
                
            except Exception as print_error:
                print(f"Error printing stock seed label: {str(print_error)}")
                return jsonify({
                    'success': False,
                    'error': f'Printing failed: {str(print_error)}'
                }), 500
        
    except Exception as e:
        print(f"Error processing stock seed label request: {str(e)}")
        return jsonify({'error': str(e)}), 500



@app.route('/print-pick-list', methods=['POST'])
def print_pick_list():
    """
    Handle pick list printing for store orders
    """
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'success': False, 'error': 'No data provided'}), 400
        
        order_id = data.get('order_id')
        order_number = data.get('order_number', 'Unknown')
        store_name = data.get('store_name', 'Unknown')
        items = data.get('items', [])
        
        if not items:
            return jsonify({'success': False, 'error': 'No items provided'}), 400
        
        # Generate the PDF (regardless of user)
        filename = f"pick_list_{order_number}.pdf"
        filepath = os.path.join("store_pick_lists", filename)
        
        # Ensure directory exists
        os.makedirs("store_pick_lists", exist_ok=True)
        
        # Create the pick list PDF
        generate_pick_list_pdf(filepath, order_number, store_name, items)
        
        if CURRENT_USER.lower() == "ndefe":
            print(f"\nPick list saved to {filepath}")
            return jsonify({
                'success': True,
                'message': f'Pick list PDF created at {filepath} (printing skipped for ndefe)'
            })
        else:
            # Print using Sumatra
            try:
                command = f'"{SUMATRA_PATH}" -print-to "{SHEET_PRINTER}" -print-settings "fit,portrait" -silent "{filepath}"'
                subprocess.run(command, check=True, shell=True)
                print(f"Successfully printed pick list {filename}")
                
                return jsonify({
                    'success': True,
                    'message': f'Pick list for {len(items)} items sent to printer'
                })
            except Exception as e:
                print(f"Failed to print pick list {filename}: {e}")
                return jsonify({
                    'success': False,
                    'error': f'Failed to print: {str(e)}'
                }), 500
            finally:
                # Clean up the file
                if os.path.exists(filepath):
                    os.remove(filepath)
                    print(f"Temporary file {filename} deleted.")
        
    except Exception as e:
        print(f"Error printing pick list: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


def generate_pick_list_pdf(filepath, order_number, store_name, items):
    """
    Generate a pick list PDF using ReportLab Platypus
    """
    from reportlab.lib.styles import ParagraphStyle, TA_RIGHT
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    
    # Create PDF document
    doc = SimpleDocTemplate(filepath, pagesize=LETTER,
                            leftMargin=40, rightMargin=40,
                            topMargin=40, bottomMargin=40)
    
    styles = getSampleStyleSheet()
    elements = []
    
    # Custom styles
    header_style = ParagraphStyle(
        name="HeaderBigBold",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=16,
        spaceAfter=6,
    )
    
    small_right = ParagraphStyle(
        name="RightSmall",
        parent=styles["Normal"],
        fontSize=12,
        alignment=TA_RIGHT,
    )
    
    # First row: Order Number + Store Name (big and bold), Date (right-aligned)
    left_header = Paragraph(f"{order_number} &nbsp;&nbsp;&nbsp; {store_name}", header_style)
    right_header = Paragraph(
        f"{datetime.now().strftime('%Y-%m-%d')}",
        small_right
    )
    
    # Header table to align them on the same row
    header_table = Table(
        [[left_header, right_header]],
        colWidths=[400, 130]
    )
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    
    elements = [header_table, Spacer(1, 12)]
    
    # Check if any items have photos
    has_photos = any(item.get('has_photo', False) for item in items)
    
    # Build table data
    headers = ["Seed", "Qty", "Variety", "Type"]
    if has_photos:
        headers.insert(1, "Photo")
    
    data = [headers]
    
    for item in items:
        row = [""]  # Seed checkbox column
        
        if has_photos:
            # Show checkmark if photo is NOT needed, empty if photo IS needed
            row.append("✓" if not item.get('has_photo', False) else "")
        
        row.append(str(item.get('quantity', 0)))
        row.append(item.get('variety_name', 'Unknown'))
        row.append(item.get('veg_type', 'Unknown'))
        
        data.append(row)
    
    # Table column widths
    if has_photos:
        col_widths = [50, 50, 50, 230, 150]
    else:
        col_widths = [50, 50, 230, 150]
    
    # Create the table
    table = Table(data, colWidths=col_widths, repeatRows=1, hAlign='LEFT')
    
    # Table style
    style = TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),  # Center seed checkbox column
        ('ALIGN', (1 if not has_photos else 2, 0), (1 if not has_photos else 2, -1), 'CENTER'),  # Center Qty column
        ('GRID', (0, 0), (-1, -1), 0.25, colors.grey)
    ])
    
    if has_photos:
        style.add('ALIGN', (1, 0), (1, -1), 'CENTER')  # Center Photo column
    
    # Shade every other row
    for i in range(1, len(data)):
        if i % 2 == 1:
            style.add('BACKGROUND', (0, i), (-1, i), colors.whitesmoke)
    
    table.setStyle(style)
    elements.append(table)
    
    # Build the PDF
    doc.build(elements)
    print(f"Pick list PDF created: {filepath}")



@app.route('/print-store-order-invoice', methods=['POST'])
def print_store_order_invoice():
    """
    Handle invoice printing for finalized store orders
    """
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'success': False, 'error': 'No data provided'}), 400
        
        order = data.get('order', {})
        store = data.get('store', {})
        items = data.get('items', [])
        
        print("\n=== FLASK INVOICE DEBUG START ===")
        print(f"Order data received: {order}")
        print(f"Credit in order dict: {order.get('credit', 'NOT FOUND')}")
        print("=== FLASK INVOICE DEBUG END ===\n")
        
        if not order or not store or not items:
            return jsonify({'success': False, 'error': 'Incomplete order data'}), 400
        
        # Generate the PDF (regardless of user)
        order_number = order.get('order_number', 'Unknown')
        
        # Generate the invoice
        generate_store_invoice_pdf(order, store, items)
        
        if CURRENT_USER.lower() == "ndefe":
            print(f"\nStore invoice generated for {store.get('store_name')} - Order {order_number}")
            return jsonify({
                'success': True,
                'message': f'Invoice PDF created (printing skipped for ndefe)'
            })
        else:
            return jsonify({
                'success': True,
                'message': f'Invoice for order {order_number} sent to printer'
            })
        
    except Exception as e:
        print(f"Error printing store order invoice: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


def generate_store_invoice_pdf(order, store, items):
    """
    Generate store invoice PDF using BaseDocTemplate approach
    """
    from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, NextPageTemplate
    from datetime import datetime, timedelta 

    print("\n=== PDF GENERATION DEBUG START ===")
    print(f"Order dict in PDF generation: {order}")
    
    # Calculate order financials
    shipping = float(order.get('shipping', 0))
    credit = float(order.get('credit', 0))
    
    print(f"Shipping: {shipping}")
    print(f"Credit extracted: {credit}")
    print("=== PDF GENERATION DEBUG END ===\n")

    # Dictionary access (correct way)
    store_address = store.get('address', '')      
    store_address2 = store.get('address2', '')    
    store_city = store.get('city', '')            
    store_state = store.get('state', '')          
    store_zip = store.get('zip', '')              

    # Setup file path
    order_number = order.get('order_number', 'Unknown')
    store_name = store.get('store_name', 'Unknown').replace('/', '-')
    file_path = f"store_invoices/{store_name}_{order_number}.pdf"
    
    # Ensure directory exists
    os.makedirs("store_invoices", exist_ok=True)

    width, height = LETTER
    styles = getSampleStyleSheet()
    
    # Calculate pagination based on item count
    item_count = len(items)
    if item_count <= 25:
        num_pages = 1
    elif item_count <= 62:
        num_pages = 2
    elif item_count <= 99:
        num_pages = 3
    elif item_count <= 136:
        num_pages = 4
    elif item_count <= 173:
        num_pages = 5
    elif item_count <= 210:
        num_pages = 6
    elif item_count <= 247:
        num_pages = 7
    else:
        num_pages = 8

    # Build table data - NEW COLUMN ORDER
    data = [["Qty", "Variety", "Type", "Unit Price", "Extended"]]
    subtotal = 0
    for item in items:
        variety = item.get('variety_name', 'Unknown')
        crop = item.get('veg_type', 'Unknown')
        quantity = item.get('quantity', 0)
        price = float(item.get('price', 0))
        line_total = quantity * price
        subtotal += line_total
        
        data.append([
            str(quantity),
            variety,
            crop,
            f"${price:.2f}",
            f"${line_total:.2f}"
        ])
    
    # Create table with adjusted column widths
    table = Table(data, colWidths=[40, 193, 135, 70, 70], repeatRows=1, hAlign='LEFT')
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (1, 1), (2, -1), "LEFT"),
        ("ALIGN", (3, 0), (-1, -1), "RIGHT"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
    ]))

    # Frames with different top margins
    first_page_frame = Frame(
        45,
        30,
        width - 60,
        height - 300
    )
    
    later_pages_frame = Frame(
        45,
        30,
        width - 60,
        height - 80
    )
    
    total_due = subtotal + shipping - credit
    
    print(f"\n=== INVOICE TOTALS DEBUG ===")
    print(f"Subtotal: {subtotal}")
    print(f"Shipping: {shipping}")
    print(f"Credit: {credit}")
    print(f"Total Due: {total_due}")
    print("=== INVOICE TOTALS DEBUG END ===\n")

    # Calculate due date (Net 30)
    order_date_str = order.get('date', '')
    try:
        order_date = datetime.fromisoformat(order_date_str.replace('Z', '+00:00'))
        order_date_formatted = order_date.strftime("%m/%d/%Y")
        due_date = order_date + timedelta(days=30)
        due_date_str = due_date.strftime("%m/%d/%Y")
    except:
        order_date_formatted = 'N/A'
        due_date_str = 'N/A'
    
    # Page header functions
    def on_first_page(canvas, doc):
        canvas.setFont("Helvetica", 13)
        canvas.drawString(30, height - 30, f"Order #: {order_number}")
        canvas.drawCentredString(width / 2, height - 30, "INVOICE")
        canvas.drawRightString(width - 40, height - 30, f"PAGE 1 of {num_pages}")
        canvas.line(0, height - 40, width - 0, height - 40)
        
        # Logo
        logo_width = 100
        logo_height = 50
        logo_x = width - logo_width - 40
        logo_y = height - logo_height - 50
        canvas.drawImage(LOGO_PATH, logo_x, logo_y, width=logo_width, height=logo_height, mask='auto')
        
        # Company info
        canvas.setFont("Calibri-Bold", 14)
        canvas.drawString(50, height - 60, "Uprising Seeds")
        canvas.setFont("Calibri", 12)
        canvas.drawString(50, height - 75, "1501 Fraser St")
        canvas.drawString(50, height - 90, "Suite 105")
        canvas.drawString(50, height - 105, "Bellingham, WA 98229")
        canvas.drawString(50, height - 120, "360-778-3749")
        canvas.drawString(50, height - 135, "info@uprisingorganics.com")
        
        # Ship To box
        canvas.line(50, height - 150, width - 300, height - 150)
        canvas.setFont("Calibri-Bold", 12)
        canvas.drawString(60, height - 164, "SHIP TO:")
        canvas.line(50, height - 150, 50, height - 265)
        canvas.line(width - 300, height - 150, width - 300, height - 265)
        canvas.line(50, height - 170, width - 300, height - 170)
        
        # Right-aligned label helper
        def draw_right_label(label, y):
            text_width = canvas.stringWidth(label, "Calibri", 12)
            canvas.drawString(130 - text_width, y, label)
        
        canvas.setFont("Calibri", 12)
        
        # Ship to info
        draw_right_label("Order #:", height - 185)
        draw_right_label("Name:", height - 200)
        draw_right_label("Date:", height - 215)
        canvas.drawString(140, height - 215, order_date_formatted)
        draw_right_label("Address:", height - 230)
        
        address = store_address
        address2 = store_address2

        canvas.drawString(140, height - 230, address)
        
        if address2:
            draw_right_label("Address 2:", height - 245)
            canvas.drawString(140, height - 245, address2)
            draw_right_label("City/State/Zip:", height - 260)
            city = store_city
            state = store_state
            zipcode = store_zip
            canvas.drawString(140, height - 260, f"{city}, {state}   {zipcode}")
        else:
            draw_right_label("City/State/Zip:", height - 245)
            city = store_city
            state = store_state
            zipcode = store_zip
            canvas.drawString(140, height - 245, f"{city}, {state}   {zipcode}")
        
        # Order info
        canvas.setFont("Calibri-Bold", 12)
        canvas.drawString(140, height - 185, order_number)
        canvas.drawString(140, height - 200, store.get('store_name', 'N/A'))
        
        # Order summary box
        right_x = 550
        label_x = 435
        
        def format_currency(value):
            try:
                return f"${float(value):.2f}"
            except ValueError:
                return "$0.00"
        
        def draw_right_aligned_label_value(label, value, y_position):
            try:
                float(value)
                value_str = format_currency(value)
            except (ValueError, TypeError):
                value_str = str(value)
            
            canvas.drawString(label_x, y_position, label)
            value_width = canvas.stringWidth(value_str, "Calibri", 12)
            canvas.drawString(right_x - value_width, y_position, value_str)
        
        canvas.setFont("Calibri", 12)
        draw_right_aligned_label_value("Subtotal:", subtotal, height - 180)
        draw_right_aligned_label_value("Shipping:", shipping, height - 195)
        draw_right_aligned_label_value("Credit:", credit, height - 210)
        
        print(f"\n=== DRAWING CREDIT ON PDF ===")
        print(f"Credit value being drawn: {credit}")
        print("=== DRAWING CREDIT ON PDF END ===\n")
        
        canvas.setFont("Calibri-Bold", 12)
        draw_right_aligned_label_value("Total Due:", total_due, height - 225)
        canvas.setFont("Calibri", 12)
        draw_right_aligned_label_value("Term:", "Net 30", height - 245)
        draw_right_aligned_label_value("Due Date:", due_date_str, height - 260)
        
        # Box around order summary
        canvas.drawString(452, height - 159, "Order Summary")
        canvas.line(428, height - 145, 428, height - 265)
        canvas.line(558, height - 145, 558, height - 265)
        canvas.line(428, height - 145, 558, height - 145)
        canvas.line(428, height - 166, 558, height - 166)
        canvas.line(428, height - 265, 558, height - 265)
        canvas.line(428, height - 232, 558, height - 232)
        
        canvas.line(50, height - 265, width - 300, height - 265)
    
    def on_later_pages(canvas, doc):
        canvas.setFont("Helvetica", 13)
        canvas.drawString(30, height - 30, f"Order #: {order_number}")
        canvas.drawCentredString(width / 2, height - 30, "INVOICE")
        canvas.drawRightString(width - 40, height - 30, f"PAGE {doc.page} of {num_pages}")
        canvas.line(0, height - 40, width - 0, height - 40)
    
    # Set up the document
    doc = BaseDocTemplate(file_path, pagesize=LETTER)
    doc.addPageTemplates([
        PageTemplate(id='FirstPage', frames=first_page_frame, onPage=on_first_page),
        PageTemplate(id='LaterPages', frames=later_pages_frame, onPage=on_later_pages)
    ])
    
    # Build elements list with first page handling
    elements = [
        NextPageTemplate('LaterPages'),
        table
    ]
    
    # Build the PDF
    doc.build(elements)
    print(f"Store invoice PDF created: {file_path}")
    
    if CURRENT_USER.lower() != "ndefe":
        try:
            # Print the invoice PDF
            command = f'"{SUMATRA_PATH}" -print-to "{SHEET_PRINTER}" -print-settings "fit,portrait" -silent "{file_path}"'
            subprocess.run(command, check=True, shell=True)
            print(f"Successfully printed invoice {file_path}")
            
            # Print two labels on roll printer
            order_label = f"Order #: {order_number}"
            store_label = store.get('store_name', 'Unknown')
            
            # First label (smaller font - for inside)
            print_order_label(order_label, store_label, font_size_order=56, font_size_store=48, y_start=20)
            
            # Second label (larger font - for outside)
            print_order_label(order_label, store_label, font_size_order=64, font_size_store=54, y_start=80)
            
        except Exception as e:
            print(f"Failed to print invoice: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"{store.get('store_name')}_{order_number}.pdf saved in store_invoices/ dir")




def print_order_label(order_text, store_text, font_size_order=56, font_size_store=48, y_start=20):
    """
    Print a single order label on roll printer
    """
    printer_name = ROLL_PRINTER
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer_name)
    
    dc.StartDoc("Order Label")
    dc.StartPage()
    
    # Label dimensions
    dpi = dc.GetDeviceCaps(88)
    label_width = int(2.625 * dpi)
    label_height = int(1.0 * dpi)
    x_center = label_width // 2
    
    # Fonts
    bold_font = create_font("Times New Roman", font_size_order, bold=True)
    norm_font = create_font("Times New Roman", font_size_store)
    
    # Draw order number
    dc.SelectObject(bold_font)
    dc.TextOut(x_center - dc.GetTextExtent(order_text)[0] // 2, y_start, order_text)
    y_start += 75
    
    # Draw store name
    dc.SelectObject(norm_font)
    dc.TextOut(x_center - dc.GetTextExtent(store_text)[0] // 2, y_start, store_text)
    
    dc.EndPage()
    dc.EndDoc()
    dc.DeleteDC()


if __name__ == "__main__":
    app.run(port=5000, debug=True)  # Debug=True helps while testing
