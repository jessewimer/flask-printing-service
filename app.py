from flask import Flask, request, jsonify
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
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
import subprocess


app = Flask(__name__)
CORS(app) 

ROLL_PRINTER = "ZDesigner GX430t"
SHEET_PRINTER = "RICOH P 501"
CURRENT_USER = os.getlogin()
SUMATRA_PATH = r"C:\Users\seedy\AppData\Local\SumatraPDF\SumatraPDF.exe"


# @app.route("/")
# def home():
#     return "Flask is running locally!"

# @app.route("/print-test")
# def print_test():
#     print("✅ Print request received!")  # This will show up in your terminal
#     return "Check your terminal — message was printed!"


def create_font(name, size, bold=False, italic=False):
    weight = FW_BOLD if bold else FW_NORMAL
    return win32ui.CreateFont({
        "name": name,
        "height": -size,  # Negative for point size
        "weight": weight,
        "italic": italic,
        "charset": DEFAULT_CHARSET,
    })



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


@app.route('/print-single-front', methods=['POST'])
def print_single_front_label():
    try:
        data = request.get_json()
        quantity = int(data.get('quantity', 1))  # default to 1 if not provided

        if CURRENT_USER.lower() == "ndefe":
            # for i in range(quantity):
            print(f"Printing copy {quantity} single front labels on Ndefe's printer")
            print(f"Variety Name: {data.get('variety_name')}")
            print(f"Crop: {data.get('crop')}")
            print(f"Days: {data.get('days')}")
            print(f"SKU Suffix: {data.get('sku_suffix')}")
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
        else:
            # Gather label content (shared across copies)
            variety_name = f"'{data.get('variety_name')}'"
            variety_crop = data.get('crop')
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

        return jsonify({
            'success': True,
            'message': f'Front Single Label printed successfully ({quantity} copies)'
        })

    except Exception as e:
        print(f"Error printing germ label: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500



@app.route('/print-single-back', methods=['POST'])
def print_single_back_label():
    try:
        data = request.get_json()
        quantity = int(data.get('quantity', 1))

        if CURRENT_USER == "ndefe":
            print(f"Printing {quantity} back single labels on Ndefe's printer")
            print(f"Back1 {data.get('back1')}")
            print(f"Back2 {data.get('back2')}")
            print(f"Back3 {data.get('back3')}")
            print(f"Back4 {data.get('back4')}")
            print(f"Back5 {data.get('back5')}")
            print(f"Back6 {data.get('back6')}")
            print(f"Back7 {data.get('back7')}")

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
                return jsonify({
                    'success': False,
                    'message': 'No back lines provided'
                }), 400

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

        return jsonify({
            'success': True,
            'message': f'Back Single Label printed successfully ({quantity} copies)'
        })

    except Exception as e:
        print(f"Error printing back label: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

    


@app.route('/print-front-sheet', methods=['POST'])
def print_front_sheet():
    try:
        data = request.json
        quantity = int(data.get('quantity', 1))
        if CURRENT_USER == "ndefe":
            print(f"Printing {quantity} front sheet on Ndefe's printer")

        else:

            variety_name = f"'{data.get('variety_name')}'"
            variety_crop = data.get('crop')
            days = data.get('days')
            env_type = data.get('env_type')
            year = data.get('for_year')
            for_year = f"20{year}"
            days_year = f"{days}    Packed for 20{year}"

            desc_line1 = data.get('desc1')
            desc_line2 = data.get('desc2')
            desc_line3 = data.get('desc3')
            lot_code = data.get('lot_code')
            germination = data.get('germination')
            # rad_type = data.get('rad_type')

            if env_type == "LG Coffee":
                pkg_size = f"{data.get('pkg_size')} ••"
            elif env_type == "SM Coffee":
                pkg_size = f"{data.get('pkg_size')} •"
            else:
                pkg_size = data.get('pkg_size')

            pkg_lot_germ = f"{pkg_size}    Lot: {lot_code}    Germ: {germination}%"
            sku_suffix = data.get('sku_suffix')


            # if TRANSITION:
                
            #     # get current germination year
            #     current_germ_obj = session.query(Germination).filter(Germination.lot_id == lot.lot_id).order_by(Germination.year.desc()).first()
            #     if current_germ_obj.year == YEAR + 1:
            #         print("current_germ_obj == YEAR + 1")
            #         if current_germ_obj.germination != 0 and current_germ_obj.germination != None:
            #             print("current_germ_obj != 0 or None")
            #             # if the germination year is the same as the current year, ask if they want to print for next year
            #             result = messagebox.askyesno(
            #                 "Choose Year",
            #                 "Should the label be printed for the next sales year?",
            #             )

            #             if result:
            #                 # get current year
            #                 current_year = datetime.now().year
            #                 year = current_year + 1
            #                 product.num_printed_next_year += (num_copies * 30)
            #             else:
            #                 # year = datetime.now().year
            #                 product.num_printed += (num_copies * 30)
            #         else:
            #             product.num_printed += (num_copies * 30)
            #     else:
            #         product.num_printed += (num_copies * 30)
            # else:
            #     # year = datetime.now().year
            #     product.num_printed += (num_copies * 30)



            # # check to see if has already been printed today
            # today = datetime.now().date()
            # last_printed = session.query(LabelPrint).filter(
            #     LabelPrint.product_id == product.product_id,
            #     LabelPrint.date >= today
            # ).first()
            # if last_printed:
            #     print(f"Last printed: {last_printed.date} for product {product.product_id}")
            #     last_printed.num_printed += (num_copies * 30)
            # else:
            #     print(f"No previous print record found for product {product.product_id} today.")
            #     # if not, create a new LabelPrint object
            #     last_printed = LabelPrint(
            #         num_printed=(num_copies * 30),
            #         date=datetime.now(),
            #         product_id=product.product_id,
            #         lot=lot
            #     )
            #     session.add(last_printed)

            # session.commit()

            envelope = f"Envelope: {env_type}"

            # File name
            filename = f"{variety_name}_Labels.pdf"
            full_path = os.path.abspath(filename)
            c = canvas.Canvas(filename, pagesize=LETTER)
            
            # Label sheet dimensions
            page_width, page_height = LETTER
            margin_y = 35  # Vertical margin, slightly adjusted
            label_width = page_width / 3  # Divide the page into 3 columns
            label_height = (page_height - margin_y) / 10.44  # Divide the page into 10 rows

            # Column adjustments
            left_col_offset = 5  # Move the left column slightly left
            middle_col_offset = 0  # Keep the middle column centered
            right_col_offset = -5   # Move the right column slightly right
            col_offsets = [left_col_offset, middle_col_offset, right_col_offset]

            # Top row adjustment
            row_offset = -5  # Move rows slightly down

            # if envenlope == LG Coffee, add two ••, if SM Coffee, add one •
            if env_type == "LG Coffee":
                pkg_size = f"{pkg_size} ••"
            elif env_type == "SM Coffee":
                pkg_size = f"{pkg_size} •"
            else:
                pkg_size = pkg_size

            # Label content
            variety_name = f"'{variety_name}'"
            variety_crop = variety_crop.upper()

            # year = f"20{germ.year}"
            pkg_lot_germ = f"{pkg_size}    Lot: {lot_code}    Germ: {germination}%"
            days_year = f"Days: {days}    Packed for {for_year}"

            # Calculate the position for the envelope
            envelope_x = 40  # Adjust to position the envelope to the left
            envelope_y = page_height - margin_y - 10 * label_height - 18  # Position below the last row of labels

            # Draw labels (3 columns x 10 rows)
            for row in range(10):
                y = page_height - margin_y - (row * label_height) + row_offset
                for col in range(3):
                    x = (col * label_width) + (label_width / 2) + col_offsets[col]

                    # Draw each label
                    c.setFont("Times-Bold", 12)
                    c.drawCentredString(x, y - 10, variety_name)

                    c.setFont("Times-Bold", 12)
                    c.drawCentredString(x, y - 23, variety_crop)

                    c.setFont("Times-Italic", 9)
                    c.drawCentredString(x, y - 33, desc_line1) 
                    c.drawCentredString(x, y - 43, desc_line2)  

                    c.setFont("Times-Roman", 8)
                    c.drawCentredString(x, y - 54, pkg_lot_germ)  
                    c.drawCentredString(x, y - 63, days_year)  

            # Set font and draw the envelope text
            c.setFont("Times-Bold", 12)
            c.drawString(envelope_x, envelope_y, envelope)
            c.save() 

            # Handle printing
            try:
                for _ in range(quantity):
                    command = f'"{SUMATRA_PATH}" -print-to "{SHEET_PRINTER}" -print-settings "fit,portrait" -silent "{full_path}"'
                    subprocess.run(command, check=True, shell=True)

            except Exception as e:
                print(f"Failed to print: {e}")
            finally:
                # Clean up the file
                if os.path.exists(filename):
                    os.remove(filename)
                                        
    except Exception as e:
        print(f"Error printing back label: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500
    



if __name__ == "__main__":
    app.run(port=5000, debug=True)  # Debug=True helps while testing
