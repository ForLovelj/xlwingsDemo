# -*- coding:utf-8 -*-
import math
import xlwings as xw
from operator import itemgetter, attrgetter

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

    def extractData(self):
        data = {}
        # visible 控制 Excel 打开是否显示界面
        # add_book 控制是否添加新的 workbook
        app = xw.App(visible=True, add_book=False)

        # 打开 data.xlsx 文件到 wookbook 中
        wb = app.books.open("data/d.xls")
          # 切换到当前活动的 sheet 中
        sheet = wb.sheets[0]
        #总行数
        # rows = sheet.api.UsedRange.Rows.count 
        #总列数
        # cols = sheet.api.UsedRange.columns.count  
        valueList = sheet.range("A1").current_region.value
        # valueList = sheet.range("A1:AO500").value
        print("sheet1数据条数---- %s ----"%len(valueList))
        print("sheet1列名：",valueList[0],sep="\n")   
        BRAND = "红牛"
        i = 0
        for v in valueList:
            # print("key:%s value:%s"%key%value)
            if i>0:
                #取办事处、客户编码、订单量
                name = str(v[0]).strip()
                id = str(v[9]).strip()
                brand = str(v[29]).strip()
                num = int(v[31])
                if(BRAND != brand):
                    continue
                if data.get(id) != None:
                    #销量累加
                    data[id][1] += num
                else:
                    value = [name,num]
                    data[id] = value
                # print("key:%s value:%s"%key%value)
            i += 1
            
        

        #提取第二个sheet数据
        #引用第二个表单
        sheet = wb.sheets[1]
        #将所引用的表单设为活动表单
        sheet.activate
        #引用活动表单
        sheet = wb.sheets.active
        valueList2 = sheet.range("A1").current_region.value
        # valueList2 = sheet.range("A1:AI5").value
        print("sheet2数据条数---- %s ----"%len(valueList2))
        print("sheet2列名：",valueList2[0],sep="\n")   
        j = 0
        for v2 in valueList2:
            # print("key:%s value:%s"%key%value)
            if j>0:
                #取办事处、客户编码、订单量
                name2 = str(v2[2]).strip()
                id2 = str(v2[7]).strip()
                brand2 = str(v2[26]).strip()
                num2 = int(v2[27])
                if(BRAND != brand2):
                    continue
                if data.get(id2) != None:
                    #销量累加
                    data[id2][1] += num2
                    print("相同客户编码：%s 在第%s行  销量已累加---"%(id2,j+1))
                else:
                    value2 = [name2,num2]
                    data[id2] = value2
            j += 1

        # print(data)
        #创建新表
        wb.sheets.add("sheet5")
        # 引用
        sheet = wb.sheets["sheet5"]
        #将所引用的表单设为活动表单
        sheet.activate
        #引用活动表单
        sheet = wb.sheets.active
        #写入数据
        sheet.range("A1").value = ["办事处","客户编码","销量总计"]
        formatData = []
        for k,v in data.items():
            formatData.append([v[0],k,v[1]])

        # sorted(formatData,key=itemgetter(0))
        formatData.sort(key=itemgetter(0))
        sheet.range("A2").value = formatData
        print("---------------数据处理完成---------------")
        wb.save()
        wb.close()
        app.quit()
