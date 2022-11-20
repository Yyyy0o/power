import PySimpleGUI as sg
import pandas as pd

import main

header = ['区县', '停电数', '发电数']
statistic = '数据统计'
export = '导出PDF'


def get_table(data):
    list = []
    for key in data.groupby('区县').groups.keys():
        zone = data[data.区县 == key]
        power = zone[zone.开始发电时间 != '无']
        list.append([key, len(zone), len(power)])
    frame = pd.DataFrame(list, columns=header)

    return frame

def get_line_chart(data):
    print(data)
    data.cumsum()
    data.plot()

def show_detail(data):
    layout = [[sg.Text(header[0]), sg.Text(header[1]), sg.Text(header[2])]]
    df = get_table(data)
    for index, row in df.iterrows():
        layout.append([sg.Text(row[header[0]]), sg.Text(row[header[1]]), sg.Text(row[header[2]])])
    layout.append((sg.Text('合计'), sg.Text(df[header[1]].sum()), sg.Text(df[header[2]].sum())))
    layout.append([sg.Button(statistic), sg.Button(export), sg.Button('退出')])

    window_detail = sg.Window('文件详情', layout, resizable=True)

    while True:
        event, value = window_detail.read()
        if event == sg.WIN_CLOSED or event == '退出':
            break
        if event == statistic:
            print(statistic)
        if event == export:
            print(export)

    window_detail.close()


if __name__ == '__main__':
    path = '停电记录.xlsx'
    data = main.parse_file(path)
    get_line_chart(data)
