"""
    Group Name: Oasong
    Member 1.59070098 Paroot Satjawanit
           2.59070156 Vasanchai Prakobkij
    Program Name: Cigarette Analyze
"""
from openpyxl import load_workbook
import pygal
def region():
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
        ktm.append(sheet_ranges[cell_c].value)
        mid.append(sheet_ranges[cell_d].value)
        north.append(sheet_ranges[cell_e].value)
        esan.append(sheet_ranges[cell_f].value)
        south.append(sheet_ranges[cell_g].value)
    line_chart = pygal.Dot()
    line_chart.title = 'Static of Smoker sorted by Region & Education'
    line_chart.x_labels = map(str, level)
    line_chart.add('กรุงเทพฯ', ktm)
    line_chart.add('กลาง', mid)
    line_chart.add('เหนือ', north)
    line_chart.add('อิสาน', esan)
    line_chart.add('ใต้', south)
    line_chart.render_to_file('Graph_region.svg')
    age()

def age():
    wb = load_workbook(filename = 'ages.xlsx')
    sheet_ranges = wb['Sheet1']
    level = []
    never = []
    always = []
    sometimes = []
    for i in range(2, 8):
        cell_a = 'A'+str(i)
        cell_b = 'B'+str(i)
        cell_c = 'C'+str(i)
        cell_d = 'D'+str(i)
        level.append(sheet_ranges[cell_a].value)
        never.append(sheet_ranges[cell_b].value)
        always.append(sheet_ranges[cell_c].value)
        sometimes.append(sheet_ranges[cell_d].value)
    line_chart = pygal.HorizontalBar()
    line_chart.title = 'Static of Smoker sorted by Age&Sex'
    line_chart.x_labels = map(str, level)
    line_chart.add('ไม่เคยสูบ', never)
    line_chart.add('สูบประจำ', always)
    line_chart.add('สูบนานๆครั้ง', sometimes)
    line_chart.render_to_file('Graph_age_sex.svg')
    job()

def job():
    wb = load_workbook(filename = 'jobs.xlsx')
    sheet_ranges = wb['Sheet1']
    jobs = []
    all_r = []
    ktm = []
    mid = []
    north = []
    esan = []
    south = []
    for i in range(2, 11):
        cell_a = 'A'+str(i)
        cell_b = 'B'+str(i)
        cell_c = 'C'+str(i)
        cell_d = 'D'+str(i)
        cell_e = 'E'+str(i)
        cell_f = 'F'+str(i)
        cell_g = 'G'+str(i)
        jobs.append(sheet_ranges[cell_a].value)
        ktm.append(sheet_ranges[cell_c].value)
        mid.append(sheet_ranges[cell_d].value)
        north.append(sheet_ranges[cell_e].value)
        esan.append(sheet_ranges[cell_f].value)
        south.append(sheet_ranges[cell_g].value)
    line_chart = pygal.Bar()
    line_chart.title = 'Static of Smoker sorted by jobs'
    line_chart.x_labels = map(str, jobs)
    line_chart.add('กรุงเทพมหานคร', ktm)
    line_chart.add('กลาง', mid)
    line_chart.add('เหนือ', north)
    line_chart.add('อีสาน', esan)
    line_chart.add('ใต้', south)
    line_chart.render_to_file('Graph_jobs.svg')
region()