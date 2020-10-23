#!/usr/bin/env python3
import xlrd
import xlwt
import sys

xlsinfile = sys.argv[1]
xlsoutfile = sys.argv[2]

wb = xlwt.Workbook()

rb = xlrd.open_workbook(filename=xlsinfile, encoding_override="cp1252")
for sn in rb.sheet_names():
	if (len(sn) > 31):
		ws = wb.add_sheet(sn[0:30])
	else:
		ws = wb.add_sheet(sn)
		
rs = rb.sheet_by_index(0)
		
ncol = rs.ncols
for row in range(0, rs.nrows):
	for col in range(0, ncol):
		cell = rs.cell(row, col)
		ws.write(row, col, cell.value)
								
wb.save(xlsoutfile)
