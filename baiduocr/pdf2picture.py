from pdf2image import convert_from_path
from pdf2image.exceptions import (
 PDFInfoNotInstalledError,
 PDFPageCountError,
 PDFSyntaxError
)

pdf_path = r"D:\learnpython\ocr_yindeng\浙银租赁2020年审计报告.pdf"
images = convert_from_path(pdf_path)
for i, image in enumerate(images):
    fname = "page" + str(i) + ".png"
    image.save(fname, "PNG")