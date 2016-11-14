import pygal
bar_chart = pygal.HorizontalBar()
bar_chart.title = 'Number of Smoker\n      Jobs'
bar_chart.add('คนงาน', 1361986)
bar_chart.add('เกษตร, ประมง', 3715143)
bar_chart.add('ผู้ปฏิบัติงานด้านความสามารถทางฝีมือ', 1185974)
bar_chart.add('ผู้ปฏิบัติการเครื่องจักรโรงงาน', 859673)
bar_chart.add('ผู้บัญญัติกฏหมาย, ข้าราชการระดับผู้อาวุโส', 665292)
bar_chart.add('ช่างเทคนิค', 232035)
bar_chart.add('พนักงานบริการ', 535813)
bar_chart.add('เสมียน', 115913)
bar_chart.add('ผู้ประกอบวิชาชีพด้านต่างๆ', 106067)
bar_chart.render_to_file('hello111.svg')