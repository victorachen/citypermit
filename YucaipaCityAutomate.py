from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

packet = io.BytesIO()
can = canvas.Canvas(packet, pagesize=letter)
can.drawString(10, 100, "Hello world")
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
