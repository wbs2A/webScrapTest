from datetime import datetime

from xlwt import *

w = Workbook()
ws = w.add_sheet('Hey, Dude')

fmts = [
    'M/D/YY',
    'D-MMM-YY',
    'D-MMM',
    'MMM-YY',
    'h:mm AM/PM',
    'h:mm:ss AM/PM',
    'h:mm',
    'h:mm:ss',
    'M/D/YY h:mm',
    'mm:ss',
    '[h]:mm:ss',
    'mm:ss.0',
]

i = 0
for fmt in fmts:
    ws.write(i, 0, fmt)

    style = XFStyle()
    style.num_format_str = fmt

    ws.write(i, 1, datetime.now(), style)

    i += 1

w.save('dates.xls')
"""import os
with open(os.path.join(os.getcwd()+'\\tmp\\','links.txt'), 'r') as file:
    lines = [line.rstrip('\n') for line in file]
    for line in lines:
        print(line)

"""