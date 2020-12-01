# !/usr/bin/python
# -*- coding: UTF-8 -*-
import openpyxl
from src import XmindtoDict

def toExcel(file_name, path_name,vban, user):
        try:
                # path_name.split("/")[-1].split(".")[:-1]
                tr = '.'.join(path_name.split("/")[-1].split(".")[:-1])
                # print(tr)
                luj = file_name + '/' + tr + '.xlsx'
                wb = openpyxl.Workbook()
                fileheadr = [['用例标题(*)', '类型', '用例分类', '前置条件', '测试步骤', '期望的测试结果', '创建人', '创建时间']]

                list_ = XmindtoDict.XmindTo(path_name, user, vban)
                list_1 = fileheadr+list_[0]
                ws = wb.active
                for n in range(len(list_1)):
                        ws.append(list_1[n])
                wb.save(luj)
                return luj

        except Exception as e:
                return  str(e)
if __name__ == '__main__':
     ri = toExcel('C:/Users/ex-fzk001/Desktop','C:/Users/ex-fzk001/Desktop/test.xmind', 'v1.6.4', 'fzk')
     print(ri)