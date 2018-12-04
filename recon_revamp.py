import pandas
import io
import os
import sys

os.chdir(".")

workbook = sys.argv[1]
alco_sheet = sys.argv[2]
FA_sheet = sys.argv[3]

alco = pandas.read_excel((workbook + ".xlsx"),alco_sheet)
FA = pandas.read_excel((workbook + ".xlsx"),FA_sheet)

print(alco["Trans Amt"])