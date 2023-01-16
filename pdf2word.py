from pdf2docx import Converter
from pathlib import Path

pdf_file = Path.home() / Path('Desktop', 'book.pdf')
docx_file = Path('D:\_Python_Projects\Python-Docx\word_files', 'doc_book.docx')

# convert pdf to docx
cv = Converter(pdf_file)
cv.convert(docx_file)  # all pages by default
cv.close()
