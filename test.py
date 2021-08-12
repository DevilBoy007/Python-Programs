from io import StringIO
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
import os


# finds pdf files and prints metadata associated including original Author, date of creation, etc.
# working to find metadata associated with text contents

path = 'C:\\Users\\C-Dylan.Bakr\\Documents\\Top Secret\\tools\\Formularies\\'
files = os.listdir(path)
output_string = StringIO()
for file in files:
    pdf_file = path + file
    # with pike.open(pdf_file, password='', allow_overwriting_input=True) as pdf:
    #     pdf.save(pdf_file)
    with open(pdf_file, 'rb') as f:
         parser = PDFParser(f)
         doc = PDFDocument(parser)
         rsrcmgr = PDFResourceManager()
         device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
         interpreter = PDFPageInterpreter(rsrcmgr, device)
         for page in PDFPage.create_pages(doc):
             interpreter.process_page(page)
    print(output_string.getvalue())
