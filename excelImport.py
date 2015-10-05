__author__ = 'Mixibird'

from xlrd import open_workbook
import csv


def open_file(path):
    """
    Open and read an Excel file
    """
    try:
        book = open_workbook(path)
        return book

    except ValueError:
        print("Error opening the file")


def find_index(libro):
    # print sheet names
    names = libro.sheet_names()
    index = names.index('1s Trend')
    return index


def add_headline(bog, index):
    # Open output csv data file
    od = open('outdata.csv', 'a', encoding="ascii")
    outData = csv.writer(od, delimiter=';', lineterminator='\n')

    # Retrieve sheets
    sheet_a = bog.sheet_by_index(index)

    # Write all rows from sheet_a to csv-file
    i = 0
    row = sheet_a.row_values(i)
    row = [a.replace('\u03c6', 'phi') for a in row]
    row = [a.replace('\u03a6', 'lambda') for a in row]

    outData.writerow(row)

    od.close()


def add_to_csv(bog, index):
    # Open output csv data file
    od = open('outdata.csv', 'a', encoding="utf-8")
    outData = csv.writer(od, delimiter=';', lineterminator='\n')

    # Retrieve sheets
    sheet_a = bog.sheet_by_index(index)
    sheet_b = bog.sheet_by_index(index + 1)

    # number of rows in sheet
    count_a = sheet_a.nrows
    count_b = sheet_b.nrows
    print('count_a', count_a, 'count_b', count_b)

    # Write all rows from sheet_a to csv-file
    i = 2
    while i < count_a:
        row = sheet_a.row_values(i)
        outData.writerow(row)
        i += 1

    # Write all rows from sheet_a to csv-file
    i = 0
    while i < count_b:
        row = sheet_b.row_values(i)
        outData.writerow(row)
        i += 1

    od.close()


# ----------------------------------------------------------------------
if __name__ == "__main__":
    fileNo = 1
    while fileNo < 36:  # 36 in total
        filename = 'Omborgen ' + str(fileNo) + '.xls'
        print(filename)

        # Open first or next woorkbook (according to fileNo)
        libro = open_file(filename)

        # Find sheet number of '1s Trend' sheet
        index = find_index(libro)
        if fileNo == 1:
            add_headline(libro, index)

        add_to_csv(libro, index)

        fileNo += 1
