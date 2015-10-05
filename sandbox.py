__author__ = 'Rosa'
from xlrd import open_workbook
import csv
import time
import datetime

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
    print(index)
    return index


#----------------------------------------------------------------------
if __name__ == "__main__":
    fileNo = 5
    while fileNo < 6:
        filename = 'Omborgen ' + str(fileNo) + '.xls'
        print(filename)
        fileNo += 1

        libro = open_file(filename)
        book = find_index(libro)
        print(book)



