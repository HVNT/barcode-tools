import string
import sys
import os
import xlrd


# from PIL import Image
#
# img = Image.new('RGB', (width, height))
# img.putdata(my_list)
# img.save('image.png')


def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


def extract_images(sheet_path, img_col, id_col, row_offset):
    book = xlrd.open_workbook(sheet_path)
    sheet = book.sheet_by_index(0)    # get the first worksheet, can safely assume only one

    for row_index in range(row_offset, sheet.nrows):
        img = sheet.cell(row_index, img_col).value
        id = sheet.cell(row_index, id_col).value
        print img
        print id


def main():
    sheet_path = sys.argv[1]
    img_col = col2num(sys.argv[2])
    id_col = col2num(sys.argv[3])
    row_offset = int(sys.argv[4])
    extract_images(sheet_path, img_col, id_col, row_offset)

if __name__ == "__main__":
    main()
