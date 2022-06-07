from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
import io, openpyxl
from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

#takes inputs from excel input file, organizes it into dictionary d
def openpyxl():
    path = r'C:\Users\Lenovo\PycharmProjects\YucaipaCityPermit\input.xlsx'
    file = load_workbook(path)
    sheet = file.active

    d = {}
    for row in range(sheet.max_row):
        num = sheet.cell(row=row+1, column=2).value
        field = sheet.cell(row=row+1, column=3).value
        if type(num)==int  and field !=None:
            raw_input = sheet.cell(row=row + 1, column=4).value
            d[field] = raw_input

    #append a bunch of (hard coded) stuff to the back of the dictionary
    d = append_dic(d)
    print(d)
    return d

#append (hard coding) stuff to the excel-made-dictionary (above)
#ex)from "9_date", we can deduce that the month is "April" and day is "15"; add this stuff to the back of the dictionary
def append_dic(d):
    # todays_month = d['9_date'][0:4]
    return d

def llc(parkname):
    d = {}
    return None

#ink autocad PDF drawing with those sweet, sweet labels on bottom left hand corner
def ink_drawing():
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.drawString(10, 100, "Hitching Post Mobile Home Park, LLC")
    can.drawString(10, 85, "34642 Yucaipa Blvd, Sp #10, Yucaipa CA 92399")
    can.drawString(10, 70, "Owner - Jian Chen")
    can.drawString(10, 55, "Phone - 909 210 1491")
    can.drawString(10, 40, "Hello world")
    can.drawString(10, 25, "Hello world")
    can.drawString(10, 10, "Hello world")
    can.save()

    #move to the beginning of the StringIO buffer
    packet.seek(0)

    # create a new PDF with Reportlab
    new_pdf = PdfFileReader(packet)
    # read your existing PDF
    input_path = r'C:\Users\Lenovo\PycharmProjects\YucaipaCityPermit\input\drawing.pdf'
    output_path = r'C:\Users\Lenovo\PycharmProjects\YucaipaCityPermit\output\drawingoutput.pdf'

    existing_pdf = PdfFileReader(open(input_path, "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    # finally, write "output" to a real file
    outputStream = open(output_path, "wb")
    output.write(outputStream)
    outputStream.close()

def alterpdf():
    emptypath = 'C:\\Users\\Lenovo\\PycharmProjects\\YucaipaCityPermit\\input\\main_empty.pdf'
    filledpath = 'C:\\Users\\Lenovo\\PycharmProjects\\YucaipaCityPermit\\output\\main_filled.pdf'
    reader = PdfFileReader(emptypath)
    writer = PdfFileWriter()

    page = reader.pages[0]
    fields = reader.getFields()

    writer.addPage(page)

    # Now you add your data to the forms!
    writer.updatePageFormFieldValues(
        writer.getPage(0), {"Address": "test"}
    )

    # write "output" to PyPDF2-output.pdf
    with open(filledpath, "wb") as output_stream:
        writer.write(output_stream)

alterpdf()