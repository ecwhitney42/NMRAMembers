#!/usr/bin/env python3
import pyexcel
import pyexcel_xls
import sys

xlsfile = sys.argv[1]
csvfile = sys.argv[2]

worksheet = pyexcel.get_sheet(file_name=xlsfile, auto_detect_datetime=False)

worksheet.save_as(csvfile)

