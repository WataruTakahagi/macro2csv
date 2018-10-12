#Python3 support
#macro2csv 0.2.2 beta

#setting
class pycolor:
    BLACK = '\033[30m'
    RED = '\033[31m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    PURPLE = '\033[35m'
    CYAN = '\033[36m'
    WHITE = '\033[37m'
    END = '\033[0m'
    BOLD = '\038[1m'
    UNDERLINE = '\033[4m'
    INVISIBLE = '\033[08m'
    REVERCE = '\033[07m'

#import module
from sys import argv
import csv
import sys
try: import xlrd
except ModuleNotFoundError:
    print('xlrd module was not found.')
    print('You should run '+pycolor.RED +'\"pip install xlrd\"'+pycolor.END)
    sys.exit()

try: import xlwt
except ModuleNotFoundError:
    print('xlwt module was not found.')
    print('You should run '+pycolor.RED +'\"pip install xlwt\"'+pycolor.END)
    sys.exit()

try: import pprint
except ModuleNotFoundError:
    print('pprint module was not found.')
    print('You should run '+pycolor.RED +'\"pip install pprint\"'+pycolor.END)
    sys.exit()

#main
try: filename = argv[1]
except IndexError:
    print('retry '+pycolor.RED +'\"python macro2csv.py yourdata.xlsx\"'+pycolor.END)
    sys.exit()
wb = xlrd.open_workbook(filename)
sheet = wb.sheet_by_name('Ages ')
col = sheet.col(3)
frag,numindex,gainlist = 0,0,[]
for id in col:
    if str(id) == 'number:0.0': frag = 1
    if frag == 0:
        if str(id.value) == 'Primary standard': pass
        else: gainlist.append(numindex)
    elif frag == 1 and str(id) == 'number:0.0': gainlist.append(numindex)
    else: break
    numindex += 1
del gainlist[0:10]

#write csv
with open('ch_data.csv', 'w') as f:
    writer = csv.writer(f, lineterminator='\n')
    list = [sheet.cell_value(8,3)]+sheet.row_values(8)[13:22]
    writer.writerow(list)
    for line in gainlist:
        list = [sheet.cell_value(line,3)]+sheet.row_values(line)[13:22]
        writer.writerow(list)
print('EXPORT' +pycolor.GREEN+ ' -> ' + pycolor.BLUE+'ch_data.csv')
