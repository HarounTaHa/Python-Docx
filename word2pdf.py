from docx2pdf import convert
from pathlib import Path

docx_file = Path('D:\_Python_Projects\Python-Docx\word_files', 'doc_book.docx')
pdf_file = Path('D:\_Python_Projects\Python-Docx\word_files', 'pdf_book.pdf')


convert(docx_file, pdf_file)
