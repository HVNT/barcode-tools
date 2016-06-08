import sys
import os
import xlrd

# from PIL import Image
#
# img = Image.new('RGB', (width, height))
# img.putdata(my_list)
# img.save('image.png')


def extract_images(sheet_path, img_col, id_col, row_offset):
    book = xlrd.open_workbook(sheet_path)
    sheet = book.sheet_by_index(0)    # get the first worksheet, can safely assume only one

    for row_index in range(row_offset, sheet.nrows):
        pdf_ptr = row_index - 1
        img = sheet.cell(row_index, img_col).value
        id = sheet.cell(row_index, id_col).value
        print img
        print id


def main():
    sheet_path = sys.argv[1]
    img_col = sys.argv[2]
    id_col = sys.argv[3]
    row_offset = sys.arg[4]
    extract_images(sheet_path, img_col, id_col, row_offset)

if __name__ == "__main__":
    main()
