#!/usr/bin/env python
#-*- coding -*-


import os
import sys
import xlrd
import xlwt
from openpyxl.workbook import Workbook


__doc__ = '''Un-merge excel cell and auto fill with the first cell value in the merged cells

Usage: unMergeExcelCell.py <excel file>

excel file: the path for un-merging excel file
'''


def usage():
    print(__doc__)
    exit(-1)

def unMergeExcelCell(path):
    # if not os.path.exists(path):
    #     print(("Could not find the excel file: " % path))
    #     return

    # read merged cells for all sheets
    book = xlrd.open_workbook(path, formatting_info=False)

    writed_cells = []   # writed cell for merged cells
    # open excel file and write
    excel = xlwt.Workbook()
    for rd_sheet in book.sheets():
        # for each sheet
        wt_sheet = excel.add_sheet(rd_sheet.name)
        
        writed_cells = []

        # over write for merged cells
        for crange in rd_sheet.merged_cells:
            # for each merged_cell
            rlo, rhi, clo, chi = crange
            cell_value = rd_sheet.cell(rlo, clo).value
            for rowx in range(rlo, rhi):
                for colx in range(clo, chi):
                    wt_sheet.write(rowx, colx, cell_value)
                    writed_cells.append((rowx, colx))

        # write all un-merged cells
        for r in range(0, rd_sheet.nrows):
            for c in range(0, rd_sheet.ncols):
                if (r, c) in writed_cells:
                    continue
                cell_value = rd_sheet.cell(r, c).value
                wt_sheet.write(r, c, cell_value)

    # save the un-merged excel file
    (origin_file, ext) = os.path.splitext(path)
    ext = '.xls'
    unmerge_excel_file = origin_file + '_unmerged' + ext
    excel.save(unmerge_excel_file)
    
    return unmerge_excel_file

    #print(("Save un-merged excel file: %s" % unmerge_excel_file))


# def main():
#     if 2 != len(sys.argv):
#         usage()

#     excel_file = sys.argv[1]
#     unMergeExcelCell(excel_file)

# if __name__ == '__main__':
#     main()