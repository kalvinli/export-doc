import os, re
import barcode
from barcode.writer import ImageWriter
import qrcode


## 生成条形码
def generate_barcode(content: str, barcode_class: str, options: object|None, file_path: str):

    if barcode_class == "":
        barcode_class = 'code128'

    if options is None:
        options = {}

    barcode_object = barcode.get_barcode_class(barcode_class)

    make_barcode = barcode_object(content, writer=ImageWriter())

    file_name = "".join(re.findall(r'[a-zA-Z0-9]', content)) + '_' + barcode_class + '_barcode'
    file_path = os.path.join(file_path, file_name)
    result = make_barcode.save(file_path, options=options)
    file_name = os.path.basename(result)

    return file_name



## 生成二维码
def generate_qrcode(content: str, options: object|None, file_path: str):

    if options is None:
        options = {}

    if options.get("error_correction", "ERROR_CORRECT_H") == "ERROR_CORRECT_L":
        error_correction=qrcode.constants.ERROR_CORRECT_L
    elif options.get("error_correction", "ERROR_CORRECT_H") == "ERROR_CORRECT_M":
        error_correction=qrcode.constants.ERROR_CORRECT_M
    elif options.get("error_correction", "ERROR_CORRECT_H") == "ERROR_CORRECT_Q":
        error_correction=qrcode.constants.ERROR_CORRECT_Q
    elif options.get("error_correction", "ERROR_CORRECT_H") == "ERROR_CORRECT_H":
        error_correction=qrcode.constants.ERROR_CORRECT_H

    make_qrcode = qrcode.QRCode(
        version=options.get("version", 2),
        error_correction=error_correction,
        box_size=options.get("box_size", 10),
        border=options.get("border", 2),
    )

    make_qrcode.add_data(content)
    make_qrcode.make(fit=True)

    img = make_qrcode.make_image(fill_color=options.get("fill_color", "black"), back_color=options.get("back_color", "white"))

    file_name = "".join(re.findall(r'[a-zA-Z0-9]', content)) + '_qrcode.png'
    file_path = os.path.join(file_path, file_name)

    img.save(file_path)

    return file_name
