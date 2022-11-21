import os

import PySimpleGUI as sg
import pandas as pd

header = ['区县', '停电数', '断站数', '发电数', '停电时间', '发电时间', 'OLT中断数', '光缆中断数']
sheet_name = '停电记录'
data_key = ['区县', '停电时间', '动环来电时间', '开始发电时间', '结束发电时间']

statistic = '数据统计'
export = '导出PDF'
exit = '退出'


def parse_file(path):
    if not os.path.exists(path):
        sg.PopupOK('文件不存在')
        return None

    sheet_dic = pd.read_excel(path, sheet_name=None, index_col=0)
    if not sheet_dic.keys().__contains__(sheet_name):
        sg.PopupOK('不存在 停电记录 sheet页')
        return None

    sheet = sheet_dic[sheet_name]
    for key in data_key:
        if not sheet.keys().__contains__(key):
            sg.PopupOK('停电记录页 不存在 {} 列'.format(key))
            return None

    return sheet.get(data_key)


def import_excel(path, win_main):
    if path is None or path == '':
        sg.PopupOK('请选择文件！！！')
    else:
        value = []
        data = parse_file(path)
        total1 = 0
        total2 = 0
        for key in data.groupby('区县').groups.keys():
            zone = data[data.区县 == key]
            power = zone[zone.开始发电时间 != '无']
            value.append([key, len(zone), 0, len(power), 0, 0, 0, 0])
            total1 += len(zone)
            total2 += len(power)
        value.append(['总计', total1, 0, total2, 0, 0, 0, 0])
        win_main['_Table_'].update(value)


def show_main():
    # 界面布局，将会按照列表顺序从上往下依次排列，二级列表中，从左往右依此排列
    layout = [[sg.Text()],
              [sg.Text('文件路径：', size=(14, 1)), sg.InputText(),
               sg.FileBrowse('浏览', file_types=([("excel", "*.xls *.xlsx")])),
               sg.Button('导入')],
              [sg.Text()],
              [sg.Table(
                  values=[[]],
                  headings=header,
                  auto_size_columns=True,
                  display_row_numbers=False,
                  key='_Table_',
                  def_col_width=118,
                  justification='center',
                  right_click_selects=True,
                  num_rows=18,
                  row_height=18,
                  enable_events=True,
                  expand_x=True,
                  expand_y=True,
                  background_color='white',
                  text_color='black'
              )],
              [sg.Button(statistic), sg.Button(export), sg.Button('退出')]]

    # 创造窗口
    win_main = sg.Window('汛期发电小程序', layout, resizable=True)

    # 事件循环并获取输入值
    while True:
        event, value = win_main.read()
        # 如果用户关闭窗口或点击`Cancel`
        if event is None or event == exit:
            break
        if event == '导入':
            import_excel(value[0], win_main)
        if event == statistic:
            print(statistic)
        if event == export:
            print(export)

    win_main.close()


if __name__ == '__main__':
    show_main()
