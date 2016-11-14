from openpyxl import load_workbook
import pygal
bar_chart = pygal.HorizontalBar()
bar_chart.title = 'Number of Smoker\n      Jobs'
wb = load_workbook(filename = 'Job.xlsx')
sheet_ranges = wb['Sheet1']
x = []
y = []
for j in range(2, 11):
    cell = 'B'+str(j)
    cell1 = 'A'+str(j)
    bar_chart.add(sheet_ranges[cell1].value, sheet_ranges[cell].value)
bar_chart.render_to_file('project.svg')
