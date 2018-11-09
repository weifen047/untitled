from docx import Document
from docx.shared import Inches

document = Document()


table = document.add_table(rows=1, cols=3)
hdr_cells = table.rows[0].cells
hdr_cells[0].text.format() = 'Qty'
hdr_cells[1].text = 'Id'
hdr_cells[2].text = 'Desc'
ghnj = 'll'

document.add_page_break()

document.save('demo.docx')
