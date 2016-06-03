# made with <3

import sys
import os
import xlrd
from pyPdf import PdfFileWriter, PdfFileReader


def chunk_pdf(pdf_path):
    inputpdf = PdfFileReader(open("assets/document.pdf", "rb"))

    for i in xrange(inputpdf.numPages):
        output = PdfFileWriter()
        output.addPage(inputpdf.getPage(i))
        with open("_tmp%s.pdf" % i, "wb") as outputStream:
            output.write(outputStream)


def read_sheet(sheet_path):
    book = xlrd.open_workbook(sheet_path)
    sheet = book.sheet_by_index(0)    # get the first worksheet, can safely assume only one
    for row_index in range(1, sheet.nrows):
        pdf_ptr = row_index - 1
        product_id = sheet.cell(row_index, 0).value

        try:
            os.rename("_tmp" + str(pdf_ptr) + ".pdf", product_id + ".pdf")
            print "[OK] Product " + product_id
        except EnvironmentError:
            print "[ERR] No pdf to rename for product: " + product_id


def main():
    pdf_path = sys.argv[1]
    sheet_path = sys.argv[2]
    chunk_pdf(pdf_path)
    read_sheet(sheet_path)

if __name__ == "__main__":
    main()
