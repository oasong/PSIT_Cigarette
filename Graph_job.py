from openpyxl import load_workbook
import pygal
bar_chart = pygal.HorizontalBar()
bar_chart.title = 'Number of Smoker\n      Jobs'
wb = load_workbook(filename = 'job.xlsx')
sheet_ranges = wb['Sheet1']
for j in range(2, 10):
    cell = 'B'+str(j)
    cell1 = 'A'+str(j)
    bar_chart.add(sheet_ranges[cell1].value, sheet_ranges[cell].value)
bar_chart.render_to_file('project_job.svg')