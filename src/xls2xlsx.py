#!/usr/bin/env python3
import pyexcel
import pyexcel_xls
import pyexcel_xlsx
import sys

xlsfile = sys.argv[1]
xlsxfile = sys.argv[2]

workbook = pyexcel.save_as(file_name=xlsfile, dest_file_name=xlsxfile, encoding_override="cp1252")


