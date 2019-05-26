import xlrd
import os
import csv


def main():
    try:
        there_is_csv = False
        os.mkdir("csv")
        list = os.listdir(path='.')
        for file_name in list:
            if file_name[-5:] == '.xlsx' or file_name[-4:] == '.xls':
                print('转换"' + file_name + '"')
                transcoding(file_name)
            else:
                continue
    except OSError as reason:
        print("错误：" + str(reason))
        there_is_csv = True
    finally:
        if not there_is_csv:
            print("\n转换结束，请查看csv目录")
        input("按任意键退出窗口\n")


def transcoding(file_name):
    table = xlrd.open_workbook(file_name).sheets()[0]
    nrows = table.nrows
    with open("csv/" + file_name + '.csv', 'w', newline='', encoding="utf-8") as csvfile:
        spamwriter = csv.writer(csvfile)
        for i in range(nrows):
            # print(table.row_values(i))
            spamwriter.writerow(table.row_values(i))


main()
