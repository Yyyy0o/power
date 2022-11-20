import os

import PySimpleGUI as sg
import pandas

import detail

sheet_name = '停电记录'
data_key = ['区县', '停电时间', '动环来电时间', '开始发电时间', '结束发电时间']


def parse_file(path):
    if not os.path.exists(path):
        sg.PopupOK('文件不存在')

    sheet_dic = pandas.read_excel(path, sheet_name=None, index_col=0)
    if not sheet_dic.keys().__contains__(sheet_name):
        raise Exception('不存在 停电记录 sheet页')

    sheet = sheet_dic[sheet_name]
    for key in data_key:
        if not sheet.keys().__contains__(key):
            raise Exception('停电记录页 不存在 {} 列'.format(key))

    return sheet.get(data_key)


def show_main():
    # 界面布局，将会按照列表顺序从上往下依次排列，二级列表中，从左往右依此排列
    layout = [[sg.Text()],
              [sg.Text('文件路径：'), sg.InputText(),
               sg.FileBrowse('浏览', file_types=([("excel", "*.xls *.xlsx")])),
               sg.Button('导入')],
              [sg.Text()],
              [sg.Text()]]

    # 创造窗口
    win_main = sg.Window('汛期发电小程序', layout)
    detail_active = False

    # 事件循环并获取输入值
    while True:
        event, value = win_main.read()
        # 如果用户关闭窗口或点击`Cancel`
        if event is None:
            break
        if event == '导入':
            path = value[0]
            if path is None or path == '':
                sg.PopupOK('请选择文件！！！')
            elif not detail_active:
                data = None
                try:
                    data = parse_file(path)
                except Exception as err:
                    sg.PopupOK('读取Excel出错:{}'.format(err))
                if data is not None:
                    win_main.Hide()
                    detail.show_detail(data)
                    detail_active = False
                    win_main.UnHide()

    win_main.close()
