# author: puru Panta (pmpanta@gmail.com)
# date created: 08/09/2023
# filename: X_RandomFromXLS.py
# 

import os
import pyodbc

import xlrd
import xlwt

import random


if __name__ == '__main__':

    # current_dir = os.getcwd()
    # print('Current Directory: ' + str(current_dir))

    # Reading Filename:
    rand_num = 500;
    rd_filename = 'note_splitter_output.xls'
    wt_filename = 'Rand' + str(rand_num) + str('_' + rd_filename)
    rd_sheet_index = 0
    wt_sheetname = 'sheet1'
    

    #Holding the random rows to write
    row_list = []

    # Read "xls" file:
    workbook_rd = xlrd.open_workbook(rd_filename)
    #Get the first sheet in the workbook by index
    worksheet_rd = workbook_rd.sheet_by_index(rd_sheet_index)

    # Write "xls" file:
    workbook_wt = xlwt.Workbook()
    worksheet_wt = workbook_wt.add_sheet(wt_sheetname, cell_overwrite_ok=True)

    # All row numbers in list
    all_rownumbers_list = list(range(1, worksheet_rd.nrows))

    #Pick the random row numbers
    random_rownumbers_list = random.choices(all_rownumbers_list, k = rand_num)

    #Read the headers
    row = worksheet_rd.row_values(0)
    row_list.append(row)

    # Read the rows and put in array
    for row_number in random_rownumbers_list:
        row = worksheet_rd.row_values(row_number)
        row_list.append(row)
        # print row:
        # print(sheet.row(random_index))
        # print(row)

    # Write the rows
    for j, row_item in enumerate(row_list):
            # print('Row Item: ' + str(row_item)); # For test
            # print('\t j = ' + str(j)) # For test
            for k, col_item in enumerate(row_item):
                # print('\t\t k = ' + str(k)) # For test
                # worksheet.write(j, k, col_item)
                worksheet_wt.row(j).write(k, col_item)

    workbook_wt.save(wt_filename)
