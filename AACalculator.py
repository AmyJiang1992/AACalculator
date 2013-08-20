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
            w_sheet.write(row_index, col_index, r_sheet.cell(row_index, col_index).value)

    Matrix = []
    sums = []
    no_shares = []

    # Generate Data Matrix
    for row_index in range(r_sheet.nrows - OFFSET-1):
        row = []
        no_share = 0
        cost = 0
        for col_index in range(1, r_sheet.ncols):
            row.append(r_sheet.cell(row_index + OFFSET +1, col_index).value)
            if r_sheet.cell(row_index + OFFSET +1, col_index).value == 'x':
                no_share += 1
            elif r_sheet.cell(row_index + OFFSET +1, col_index).value != '':
                cost += r_sheet.cell(row_index + OFFSET +1, col_index).value
        Matrix.append(row)
        no_shares.append(no_share)
        sums.append(cost)
        #w_sheet.write(row_index + OFFSET + 1, r_sheet.ncols, 'test')

    print no_shares
    print sums
    print Matrix

    # Initialize should_pay and paid list
    should_pays = [0] * (r_sheet.ncols-1)
    paids = [0] * (r_sheet.ncols-1)

    # Generate Result
    for j, row in enumerate(Matrix):
        for i, cost in enumerate(row):
            if cost!= 'x':
                should_pays[i] += sums[j] / (r_sheet.ncols-1 - no_shares[j])
                if cost!='':
                    paids[i] += cost
    diffs = [a-b for a, b in zip(paids, should_pays)]

    # Print Result
    lineNo = r_sheet.nrows + 1
    colNo = 0
    w_sheet.write(lineNo, colNo, 'Paid')
    for paid in paids:
        colNo += 1
        w_sheet.write(lineNo, colNo, round(paid,2))
    lineNo += 1
    colNo = 0
    w_sheet.write(lineNo, colNo, 'Should Pay')
    for should_pay in should_pays:
        colNo += 1
        w_sheet.write(lineNo, colNo, round(should_pay,2))
    lineNo += 1
    colNo =0
    w_sheet.write(lineNo, colNo, 'Difference')
    for diff in diffs:
        colNo +=1
        w_sheet.write(lineNo, colNo, round(diff,2))

    colNo = r_sheet.ncols
    lineNo = OFFSET
    w_sheet.write(lineNo, colNo, 'Sum')
    lineNo += 1
    for cost in sums:
        w_sheet.write(lineNo, colNo, cost)
        lineNo += 1
    lineNo = r_sheet.nrows +1
    w_sheet.write(lineNo, colNo, round(sum(paids),2))
    lineNo += 1
    w_sheet.write(lineNo, colNo, round(sum(should_pays),2))
    lineNo += 1
    w_sheet.write(lineNo, colNo, round(sum(diffs),2))

    write_book.save('template.xls')

def main():
    print WELCOME
    select_mode()

if __name__ == '__main__':
    main()
