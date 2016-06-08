import string
import sys
import xlrd
import os


def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


def get_target_files(dir_path, prefix):
    directory = os.listdir(dir_path)
    target_files = []
    for file_name in directory:
        if prefix in file_name:
            target_files.append(file_name)
    return target_files


def rename_target_files(dir_path, target_files, sheet_path, id_col, row_offset):
    book = xlrd.open_workbook(sheet_path)
    sheet = book.sheet_by_index(0)  # get the first worksheet, can safely assume only one

    img_ptr = 0
    for row_index in range(row_offset, sheet.nrows):
        product_id = str(sheet.cell(row_index, id_col).value)

        try:
            os.rename(dir_path + '/' + target_files[img_ptr], product_id + ".png")
            print "[OK] Product " + str(product_id)
        except (EnvironmentError, IndexError) as e:
            print "[ERR] No file to rename for product: " + product_id

        img_ptr += 1


def main():
    rename_dir_path = sys.argv[1]
    rename_prefix = sys.argv[2]
    sheet_path = sys.argv[3]
    id_col = col2num(sys.argv[4])
    row_offset = int(sys.argv[5])

    target_files = get_target_files(rename_dir_path, rename_prefix)
    rename_target_files(rename_dir_path, target_files, sheet_path, id_col, row_offset)


if __name__ == "__main__":
    main()
