# -*- coding:utf-8 -*-
import math
import xlwings as xw
from xlwings.main import Sheet


filePath = "data/test.xlsx"

class ExcelOpt(object):

    def __init__(self,filePath = filePath) -> None:
        super().__init__()
        self.filePath = filePath

    def copy(self):
        # visible 控制 Excel 打开是否显示界面
        # add_book 控制是否添加新的 workbook
        app = xw.App(visible=True, add_book=False)

        # 打开 data.xlsx 文件到 wookbook 中
        wb = app.books.open(filePath)
        wb.save("data/test2.xlsx")
        wb.close()
        app.quit()

    def modify(self):
        # visible 控制 Excel 打开是否显示界面
        # add_book 控制是否添加新的 workbook
        app = xw.App(visible=True, add_book=False)

        # 打开 data.xlsx 文件到 wookbook 中
        wb = app.books.open(filePath)
        # 切换到当前活动的 sheet 中
        sheet = wb.sheets.active

        # 选择 A1 所在的一列
        # 当 Excel 格式复杂的时候,不建议使用 expand
        # 可以这样选择
        # ARange = sheet.range("A1:A100")
        #ARange = sheet.range("A1").expand("down")
        DRange = sheet.range("D2:D10")
        bList = sheet.range("B2:B10").value
        cList = sheet.range("C2:C10").value
        i = 0
        for d in DRange:            
            d.value = int(bList[i]) * int(cList[i])
            i = i + 1

        wb.save()
        wb.close()
        app.quit()

    def sum(self):
         # visible 控制 Excel 打开是否显示界面
        # add_book 控制是否添加新的 workbook
        app = xw.App(visible=True, add_book=False)

        # 打开 data.xlsx 文件到 wookbook 中
        wb = app.books.open(filePath)
        # 切换到当前活动的 sheet 中
        sheet = wb.sheets.active
        #选取所有有效区域数据
        # valueList = sheet.range("B2").current_region.value
        # print(valueList)
        #用于记录列数据
        rowContent = []
        #读取excel数据
        listValue = sheet.range("A2:D10").value
        for i in range(len(listValue)):
            #使用Python的类库直接访问Excel的表单是很缓慢的，不要在Python的循环中引用sheet等Excel表单的单元格，
            #而是要用List一次性读取Excel里的数据，在List内存中计算好了，然后返回结果
            bv = listValue[i][1]
            cv = listValue[i][2]
            rowContent.append(bv*cv)

        total = sum(rowContent)
        sheet.range("D11").value = total
        print("-------- %s --------"%total)
        wb.save()
        wb.close()
        app.quit()

    def check(self,path1,path2):
         # visible 控制 Excel 打开是否显示界面
        # add_book 控制是否添加新的 workbook
        app = xw.App(visible=True, add_book=False)

        # 打开 data.xlsx 文件到 wookbook 中
        wb = app.books.open(path1)
        sheet = wb.sheets.active
        valueList = sheet[0,0].current_region.value
        print(valueList)

        