from openpyxl import load_workbook
import pygal
wb = load_workbook(filename = 'region.xlsx')
sheet_ranges = wb['Sheet1']
level = []
allr = []
ktm = []
mid = []
north = []
esan = []
south = []
for i in range(2, 12):
    cell_a = 'A'+str(i)
    cell_b = 'B'+str(i)
    cell_c = 'C'+str(i)
    cell_d = 'D'+str(i)
    cell_e = 'E'+str(i)
    cell_f = 'F'+str(i)
    cell_g = 'G'+str(i)
    level.append(sheet_ranges[cell_a].value)
    allr.append(sheet_ranges[cell_b].value)
    ktm.append(sheet_ranges[cell_c].value)
    mid.append(sheet_ranges[cell_d].value)
    north.append(sheet_ranges[cell_e].value)
    esan.append(sheet_ranges[cell_f].value)
    south.append(sheet_ranges[cell_g].value)


line_chart = pygal.Bar()
line_chart.title = 'Static of Smoker sorted by Region & Education'
line_chart.x_labels = map(str, level)

line_chart.add('ทั่วราชอาณาจักร', allr)
line_chart.add('กรุงเทพฯ', ktm)
line_chart.add('กลาง', mid)
line_chart.add('เหนือ', north)
line_chart.add('อิสาน', esan)
line_chart.add('ใต้', south)
line_chart.render_to_file('Graph_region.svg')