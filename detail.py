import PySimpleGUI as sg

import main


def trans_data(data):
    group = data.groupby('区县')
    print(group.groups.keys())


def show_detail(data):
    print(data)


    table_layout = []
    layout = [table_layout,
              [sg.Button('数据统计'), sg.Button('导出PDF'), sg.Button('退出')]]

    window_detail = sg.Window('文件详情', layout, resizable=True)

    while True:
        event, value = window_detail.read()
        if event == sg.WIN_CLOSED or event == '退出':
            break

    window_detail.close()


if __name__ == '__main__':
    path = '停电记录.xlsx'
    data = main.parse_file(path)
    trans_data(data)