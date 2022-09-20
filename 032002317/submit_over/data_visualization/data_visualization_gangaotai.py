from pyecharts.charts import *
from pyecharts import options as opts
import xlrd

path = "D:\\Python\\execl_test\\gang_ao_tai_new_confirmed.xls"
sheetName = 'test1'
data = xlrd.open_workbook(path)
table = data.sheet_by_name(sheetName)

rowAmount = table.nrows
colAmount = table.ncols

x_data = []
for rowIndex in range(1, rowAmount):
    x_data.append(table.cell_value(rowIndex, 0))
type_data = []
for colIndex in range(1, colAmount):
    type_data.append(table.cell_value(0, colIndex))
# y_data = []
# for rowIndex in range(1, rowAmount):
#     y_data.append(table.cell_value(rowIndex, 1))


def get_y_data(x):
    y_data = []
    for rowIndex in range(1, rowAmount):
        y_data.append(table.cell_value(rowIndex, x))
    return y_data


def per_province_line(cols):
    line = Line(init_opts=opts.InitOpts(theme='light',
                                        width='1000px',
                                        height='600px'))
    line.add_xaxis(x_data)
    line.extend_axis(xaxis_data=x_data,
                     xaxis=opts.AxisOpts(),
                     )
    # total = len(cols)
    count = 1
    for item in cols:
        y_data = get_y_data(count)
        line.add_yaxis(series_name=item,
                       y_axis=y_data,
                       is_smooth=True,
                       label_opts=opts.LabelOpts(is_show=False),
                       # markpoint_opts=['min', 'max'],
                       markline_opts=opts.MarkLineOpts(data=[opts.MarkLineItem(type_='average')]),
                       )
        count += 1
    return line


chart = per_province_line(type_data)
chart.render_notebook()
# chart.render('test.html')
chart.load_javascript()
# line.render_notebook()
# line.render_notebook('666.html')
chart.render(path='D:\\Python\\png_test\\04.html')
# chart.render('test3.html')

