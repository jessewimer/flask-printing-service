from flask import Flask, request, jsonify
from flask_cors import CORS
from barcode import Code128
from barcode.writer import ImageWriter
import win32print
import win32ui
from win32con import FW_NORMAL, FW_BOLD, DEFAULT_CHARSET
from PIL import Image, ImageWin
import os

app = Flask(__name__)
CORS(app) 

ROLL_PRINTER = "ZDesigner GX430t"
SHEET_PRINTER = "RICOH P 501"
CURRENT_USER = os.getlogin()

def create_font(name, size, bold=False, italic=False):
    weight = FW_BOLD if bold else FW_NORMAL
    return win32ui.CreateFont({
        "name": name,
        "height": -size,  # Negative for point size
        "weight": weight,
        "italic": italic,
        "charset": DEFAULT_CHARSET,
    })


@app.route("/")
def home():
    return "Flask is running locally!"

@app.route("/print-test")
def print_test():
    print("✅ Print request received!")  # This will show up in your terminal
    return "Check your terminal — message was printed!"


@app.route('/print-germ-label', methods=['POST'])
def print_germ_label():
    try:
        data = request.get_json()

        if CURRENT_USER.lower() == "ndefe":
            print("Printing on Ndefe's printer")
            print("=== GERM SAMPLE PRINT REQUEST ===")
            print(f"Variety Name: {data.get('variety_name')}")
            print(f"SKU Prefix: {data.get('sku_prefix')}")
            print(f"Species: {data.get('species')}")
            print(f"Lot Code: {data.get('lot_code')}")
            print(f"Germ Year: {data.get('germ_year')}")
            print("================================")
        
        else: 

            # === Construct label text ===
            variety = data.get('variety')
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

        if CURRENT_USER.lower() == "ndefe":
            print("Printing on Ndefe's printer")
            print("=== PRINTING FRONT SINGLE LABEL(S) ===")
            print(f"Variety Name: {data.get('variety_name')}")
            # print(f"SKU Prefix: {data.get('sku_prefix')}")
            print(f"Species: {data.get('species')}")
            print(f"Lot Code: {data.get('lot_code')}")
            print(f"Germ Year: {data.get('germ_year')}")
            print("================================")
        
        else: 
            pass

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

if __name__ == "__main__":
    app.run(port=5000, debug=True)  # Debug=True helps while testing
