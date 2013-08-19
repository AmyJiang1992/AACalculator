#!/usr/bin/env python

from xlutils.copy import copy
from xlwt import Workbook, easyxf
from xlrd import open_workbook, cellname

WELCOME = ("%s%s%s%s%s") % ('*'*50, '\n', ' '*12, 'Welcome to AACalculator\n', '*'*50)
CHOICE  = '\nPlease select from the following options:\n  1.Generate Template\n  2.Generate Result\n'
DEFAULT = 4
OFFSET = 4

def select_mode():
    choice = raw_input(CHOICE)
    if choice == '1':
        gen_template()
    elif choice == '2':
        gen_result()
    else:
        print '\nError: Wrong input. Please ONLY enter the number of option\n'
        select_mode()

def gen_template():
    file = 'template.xls'
    book = Workbook(encoding="utf-8")
    sheet1 = book.add_sheet('1st share')
    lineNo = OFFSET
    tmp_start = lineNo
    for i in range(DEFAULT):
        sheet1.write(lineNo, i+1, 'Person' + str(i+1))
    lineNo += 1
    for i in range(DEFAULT):
        sheet1.write(OFFSET+i+1, 0, 'Event' + str(i+1))
    book.save(file)

def gen_result():
    people = []
    read_book = open_workbook('template.xls')
    r_sheet = read_book.sheet_by_index(0)
    write_book = copy(read_book)
    w_sheet = write_book.get_sheet(0)
    # Copy read_book to write_book
    for row_index in range(r_sheet.nrows):
        for col_index in range(r_sheet.ncols):
            print w_sheet.write(row_index, col_index, r_sheet.cell(row_index, col_index).value)

    # Generate Summasion Row
    for row_index in range(r_sheet.nrows - OFFSET-1):
        for col_index in range(1, r_sheet.ncols):
            if r_sheet.cell(row_index, col_index).value != '':
                print r_sheet.cell(row_index, col_index).value
        w_sheet.write(row_index + OFFSET + 1, r_sheet.ncols, 'test')
    write_book.save('template.xls')

    '''
    for col_index in range(sheet.ncols-1):
        people.append(sheet.cell(0, col_index+1).value)
    '''


def main():
    print WELCOME
    select_mode()

if __name__ == '__main__':
    main()
