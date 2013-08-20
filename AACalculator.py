#!/usr/bin/env python

from xlutils.copy import copy
from xlwt import Workbook, easyxf
from xlrd import open_workbook, cellname

WELCOME = ("%s%s%s%s%s") % ('*'*50, '\n', ' '*12, 'Welcome to AACalculator\n', '*'*50)
CHOICE  = '\nPlease select from the following options:\n  1.Generate Template\n  2.Generate Result\n'
DEFAULT = 5
OFFSET = 7
DESC_LEN = 15

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
    sheet = book.add_sheet('1st share')
    sheet.write_merge(0,0,0,DESC_LEN, 'How to use:')
    sheet.write_merge(1,1,1,DESC_LEN, '1.Rename the person\'s name and event name to its actual name')
    sheet.write_merge(2,2,1,DESC_LEN, '2.Add cost for each event under the name for who paid for it. Note: there can be more than one people paid for an event')
    sheet.write_merge(3,3,1,DESC_LEN, '3.For those who shouldn\'t share an event, put \'x\' there')
    sheet.write_merge(4,4,1,DESC_LEN, '4.Double check the data, Make sure there\'s no redundant data in the table. Then save and exit Excel. Re-run the script and choose the 2nd option')
    rowNo = OFFSET
    tmp_start = rowNo
    for i in range(DEFAULT):
        sheet.write(rowNo, i+1, 'Person' + str(i+1))
    rowNo += 1
    for i in range(DEFAULT):
        sheet.write(OFFSET+i+1, 0, 'Event' + str(i+1))
    book.save(file)
    print 'Template is generated, please open template.xls fill in AA data'

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
    rowNo = r_sheet.nrows + 1
    colNo = 0
    w_sheet.write(rowNo, colNo, 'Paid')
    for paid in paids:
        colNo += 1
        w_sheet.write(rowNo, colNo, round(paid,2))
    rowNo += 1
    colNo = 0
    w_sheet.write(rowNo, colNo, 'Should Pay')
    for should_pay in should_pays:
        colNo += 1
        w_sheet.write(rowNo, colNo, round(should_pay,2))
    rowNo += 1
    colNo =0
    w_sheet.write(rowNo, colNo, 'Difference')
    for diff in diffs:
        colNo +=1
        w_sheet.write(rowNo, colNo, round(diff,2))

    colNo = r_sheet.ncols
    rowNo = OFFSET
    w_sheet.write(rowNo, colNo, 'Sum')
    rowNo += 1
    for cost in sums:
        w_sheet.write(rowNo, colNo, cost)
        rowNo += 1
    rowNo = r_sheet.nrows +1
    w_sheet.write(rowNo, colNo, round(sum(paids),2))
    rowNo += 1
    w_sheet.write(rowNo, colNo, round(sum(should_pays),2))
    rowNo += 1
    w_sheet.write(rowNo, colNo, round(sum(diffs),2))

    write_book.save('template.xls')

def main():
    print WELCOME
    select_mode()

if __name__ == '__main__':
    main()
