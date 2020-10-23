#!/usr/bin/env python3
from openpyxl import Workbook
import csv
import sys

csvfile = sys.argv[1]
xlsxfile = sys.argv[2]

wb = Workbook()
ws = wb.active
with open(csvfile, 'r') as f:
    for row in csv.reader(f):
        ws.append(row)
wb.save(xlsxfile)
