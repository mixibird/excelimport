__author__ = 'Mixibird'

from xlrd import open_workbook
import csv
import datetime
import os

def open_file(path):
    """
    Open and read an Excel file
    """
    try:
        book = open_workbook(path)
        return book

    except ValueError:
        print("Error opening the file")


def find_index(bog):
    # print sheet names
    names = bog.sheet_names()
    indeks = names.index('1s Trend')
    return indeks


def add_headline(bog, indeks):
    # Open output csv data file
    od = open('outdata.csv', 'w', encoding="utf-8")
    outData = csv.writer(od, delimiter=';', lineterminator='\n')

    # Retrieve sheets
    sheet_a = bog.sheet_by_index(indeks)

    # Write all rows from sheet_a to csv-file
    i = 0
    row = sheet_a.row_values(i)
    row = [a.replace('\u03c6', 'phi') for a in row]
    row = [a.replace('\u03a6', 'lambda') for a in row]

    outData.writerow(row)
    print(row)

    od.close()


def add_to_csv(bog, indeks):
    # Open output csv data file
    od = open('outdata.csv', 'a', encoding="utf-8")
    outData = csv.writer(od, delimiter=';', lineterminator='\n')

    # Retrieve sheets
    sheet_a = bog.sheet_by_index(indeks)
    sheet_b = bog.sheet_by_index(indeks + 1)

    # number of rows in sheet
    count_a = sheet_a.nrows
    count_b = sheet_b.nrows
    print('count_a', count_a, 'count_b', count_b)

    # Write all rows from sheet_a to csv-file
    i = 2
    while i < count_a:
        row = sheet_a.row_values(i)

        # inserts
        dato = row[0]
        tid = row[1]
        date = dato + ' ' + tid
        py_date = datetime.datetime.strptime(date, '%d-%m-%Y %H:%M:%S')

        # Remove the original date
        row.pop(0)

        # Remove the original time
        row.pop(0)

        # Insert inverted and combined date and time
        position = 0
        row.insert(position, str(py_date))

        outData.writerow(row)
        i += 1

    # Write all rows from sheet_a to csv-file
    i = 0
    while i < count_b:
        row = sheet_b.row_values(i)

        # inserts
        dato = row[0]
        tid = row[1]
        date = dato + ' ' + tid
        py_date = datetime.datetime.strptime(date, '%d-%m-%Y %H:%M:%S')

        # Remove the original date
        row.pop(0)

        # Remove the original time
        row.pop(0)

        # Insert inverted and combined date and time
        position = 0
        row.insert(position, str(py_date))

        outData.writerow(row)
        i += 1

    od.close()


# ----------------------------------------------------------------------
if __name__ == "__main__":
    fileNo = 1
    path = '/home/mixibird/PycharmProjects/appendix'
    while fileNo < 36:  # 36 in total
        filename = 'Omborgen ' + str(fileNo) + '.xls'
        print(filename)
        syspath = os.path.join(path, filename)
        print(syspath)

        # Open first or next woorkbook (according to fileNo)
        libro = open_file(syspath)

        # Find sheet number of '1s Trend' sheet
        index = find_index(libro)
        if fileNo == 1:
            add_headline(libro, index)

        add_to_csv(libro, index)

        fileNo += 1
