from openpyxl import load_workbook
wb = load_workbook(filename = 'hello2.xlsx')
sheet_ranges = wb['Sheet1']
x = []
y = [1,2,3,4,5,6,7,8,9]
for j in range(2, 10):
    cell = 'B'+str(j)
    x.append(sheet_ranges[cell].value)
from pylab import *
hist(y)
show()