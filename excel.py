# -*- coding:utf-8 -*-
import math
import xlwings as xw
from operator import itemgetter, attrgetter

filePath = "data/test.xlsx"
tempPath = "data/d.xls"


#见费用网点明细表
class RowData(object):
    def __init__(self,name,id,brand,data={}):
        #办事处
        self.name = name
        #客户编码
        self.id = id
        #品牌
        self.brand = brand
        #红牛(战马、果倍爽)供应商数据[红牛供应商编码1:红牛1，红牛供应商编码2:红牛2]
        self.data = data

    def __str__(self):
        return "[name:%s,id:%s,brand:%s,data:%s]" %(self.name,self.id,self.brand,self.data)
        

class ExcelOpt(object):

    def __init__(self,filePath = filePath) -> None:
        super().__init__()
        self.filePath = filePath
        self.app = xw.App(visible=True, add_book=False)

    def copy(self):
        # visible 控制 Excel 打开是否显示界面
        # add_book 控制是否添加新的 workbook
        # app = xw.App(visible=True, add_book=False)
        app = self.app
        # 打开 data.xlsx 文件到 wookbook 中
        wb = self.app.books.open(filePath)
        wb.save("data/test2.xlsx")
        wb.close()
        app.quit()

    def modify(self):
        # visible 控制 Excel 打开是否显示界面
        # add_book 控制是否添加新的 workbook
        # app = xw.App(visible=True, add_book=False)
        app = self.app
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
        # app = xw.App(visible=True, add_book=False)
        app = self.app
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
        # app = xw.App(visible=True, add_book=False)
        app = self.app
        # 打开 data.xlsx 文件到 wookbook 中
        wb = app.books.open(path1)
        sheet = wb.sheets.active
        valueList = sheet[0,0].current_region.value
        print(valueList)

    #基于订单明细表、MIT订单表 提取出办事处、客户编码、订单量
    def extractData(self,path=tempPath):
        data = {}
        # visible 控制 Excel 打开是否显示界面
        # add_book 控制是否添加新的 workbook
        # app = xw.App(visible=True, add_book=False)
        app = self.app

        # 打开 data.xlsx 文件到 wookbook 中
        wb = app.books.open(path)
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
        i = -1
        for v in valueList:
            i += 1
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
                    # print("相同客户编码：%s 在第%s行  销量已累加---"%(id,i+1))
                else:
                    value = [name,num]
                    data[id] = value
                # print("key:%s value:%s"%key%value)
            
            
        

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
        j = -1
        for v2 in valueList2:
            # print("key:%s value:%s"%key%value)
            j += 1
            if j>0:
                #取办事处、客户编码、订单量
                name2 = str(v2[2]).strip()
                id2 = str(v2[7]).strip()
                brand2 = str(v2[26]).strip()
                num2 = int(v2[27])
                # print("%s"%id2)
                if(BRAND != brand2):
                    continue
                if data.get(id2) != None:
                    #销量累加
                    data[id2][1] += num2
                    print("相同客户编码：%s 在第%s行  销量已累加---"%(id2,j+1))
                else:
                    value2 = [name2,num2]
                    data[id2] = value2
            

        # print(data)
        #创建新表
        wb.sheets.add("temp")
        # 引用
        sheet = wb.sheets["temp"]
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
        wb.save("data/d1.xls")
        wb.close()
        app.quit()

    #基于订单明细表、MIT订单表 提取出费用网点明细表
    def extractDataDetail(self,path=tempPath):
        app = self.app

        # 打开 data.xlsx 文件到 wookbook 中
        wb = app.books.open(path)
          # 切换到当前活动的 sheet 中
        sheet = wb.sheets[0]
        valueList = sheet.range("A1").current_region.value
        # valueList = sheet.range("A1:AO5").value
        print("sheet1数据条数---- %s ----"%len(valueList))
        print("sheet1列名：",valueList[0],sep="\n")
        BRAND_HN = "红牛"
        BRAND_ZM = "战马"
        BRAND_GBS = "果倍爽"
        #记录转换的所有行数据{id:rowData}
        data = {}
        #[name,id,brand,{supplierId:num}]
        # i = -1
        for i,v in enumerate(valueList):
            # i += 1
            # print("key:%s value:%s"%key%value)
            if i>0:
                #取办事处、客户编码、品牌、供应商编码、订单量、产品小计
                name = str(v[0]).strip()
                id = str(v[9]).strip()
                brand = str(v[29]).strip()
                supplierId = str(v[16]).strip()
                num = int(v[31])
                totalPrice = float(v[36])
                # print("name:%s id:%s brand:%s supplierId: %s num %s"%(name,id,brand,supplierId,num))
                #构造品牌唯一key
                key = brand+"-"+supplierId
                rowData = data.get(id)
                if rowData != None:                   
                    if rowData[2].get(key) != None:
                        #销量累加 PS:红牛取销量 战马和果倍爽取产品小计
                        if brand == BRAND_HN: 
                            rowData[2][key] += num
                        else:
                            rowData[2][key] += totalPrice
                        
                        # print("相同客户编码：%s 在第%s行  销量已累加---"%(id,i+1))
                    else:
                        #第N个供应商编码对应销量 PS:红牛取销量 战马和果倍爽取产品小计
                        if brand == BRAND_HN: 
                            rowData[2][key] = num
                        else:
                            rowData[2][key] = totalPrice
                        
                else:
                    #列数据 PS:红牛取销量 战马和果倍爽取产品小计
                    if brand == BRAND_HN: 
                        rowData = [name,id,{key:num}]
                    else:
                         rowData = [name,id,{key:totalPrice}]
                    data[id] = rowData

        # for d,v in data.items():
        #     print(v)
        #提取第二个sheet数据
        #引用第二个表单
        sheet = wb.sheets[1]
        #将所引用的表单设为活动表单
        sheet.activate
        #引用活动表单
        sheet = wb.sheets.active
        valueList2 = sheet.range("A1").current_region.value
        # valueList2 = sheet.range("A1:AI50").value
        print("sheet2数据条数---- %s ----"%len(valueList2))
        print("sheet2列名：",valueList2[0],sep="\n")   
        # j = -1
        for j,v in enumerate(valueList2):
            # print("key:%s value:%s"%key%value)
            # j += 1
         #取办事处、客户编码、品牌、供应商编码、订单量、产品小计
            if j > 0:
                name = str(v[2]).strip()
                id = str(v[7]).strip()
                brand = str(v[26]).strip()
                supplierId = str(v[12]).strip()
                num = int(v[27])
                totalPrice = float(v[30])
                #构造品牌唯一key
                key = brand+"-"+supplierId
                rowData = data.get(id)
                if rowData != None:                   
                    if rowData[2].get(key) != None:
                         #销量累加 PS:红牛取销量 战马和果倍爽取产品小计
                        if brand == BRAND_HN: 
                            rowData[2][key] += num
                        else:
                            rowData[2][key] += totalPrice
                        # print("相同客户编码：%s 在第%s行  销量已累加---"%(id,j+1))
                    else:
                        #第N个供应商编码对应销量 PS:红牛取销量 战马和果倍爽取产品小计
                        if brand == BRAND_HN: 
                            rowData[2][key] = num
                        else:
                            rowData[2][key] = totalPrice
                else:
                     #列数据 PS:红牛取销量 战马和果倍爽取产品小计
                    if brand == BRAND_HN: 
                        rowData = [name,id,{key:num}]
                    else:
                         rowData = [name,id,{key:totalPrice}]
                    data[id] = rowData


        # print(data)
        #创建新表
        wb.sheets.add("费用网点明细表")
        # 引用
        sheet = wb.sheets["费用网点明细表"]
        #将所引用的表单设为活动表单
        sheet.activate
        #引用活动表单
        sheet = wb.sheets.active
        formatData = []
        # 客户编码对应的红牛供应商编码最大数（动态的）
        hnSupplierNum = 1
        # 客户编码对应的战马供应商编码最大数（动态的）
        zmSupplierNum = 1
        # 客户编码对应的果倍爽供应商编码最大数（动态的）
        gbsSupplierNum = 1
        for k,v in data.items():
            hnSupplierNum2 = 0
            zmSupplierNum2 = 0
            gbsSupplierNum2 = 0
            for k1,v1 in v[2].items():
                brandSupplierId = k1.split("-")
                brand = brandSupplierId[0]
                if BRAND_HN == brand:
                    hnSupplierNum2 += 1
                elif BRAND_ZM == brand:
                   zmSupplierNum2 += 1
                elif BRAND_GBS == brand:
                   gbsSupplierNum2 += 1
                else:
                    print("未知品牌：%s 请校验表格数据-----"%brand)
            hnSupplierNum = max(hnSupplierNum,hnSupplierNum2)
            zmSupplierNum = max(zmSupplierNum,zmSupplierNum2)
            gbsSupplierNum = max(gbsSupplierNum,gbsSupplierNum2)
           
        print("红牛供应商最大数量：%s 战马供应商最大数量：%s 果倍爽供应商最大数量: %s"%(hnSupplierNum,zmSupplierNum,gbsSupplierNum))
        for k,v in data.items():
            row = [v[0],v[1]]
            x = 0
            y = 0
            z = 0
            for k1,v1 in v[2].items():
                brandSupplierId = k1.split("-")
                brand = brandSupplierId[0]
                supplierId = brandSupplierId[1]
                if(brand == BRAND_HN):
                    row.append(supplierId)
                    row.append(v1)
                    x+=1
            diff = hnSupplierNum-x
            if(diff>0):
                #补齐表格数据
                for i in range(diff):
                    row.append("") 
                    row.append("")      
            for k2,v2 in v[2].items():    
                brandSupplierId = k2.split("-")
                brand = brandSupplierId[0]
                supplierId = brandSupplierId[1]
                if(brand == BRAND_ZM):
                    row.append(supplierId)
                    row.append(v2)
                    y+=1
            diff = zmSupplierNum-y
            if(diff>0):
                #补齐表格数据
                for i in range(diff):
                    row.append("") 
                    row.append("")     
            for k3,v3 in v[2].items():    
                brandSupplierId = k3.split("-")
                brand = brandSupplierId[0]
                supplierId = brandSupplierId[1]
                if(brand == BRAND_GBS):
                    row.append(supplierId)
                    row.append(v3)
                    z+=1
            diff = gbsSupplierNum-z
            if(diff>0):
                #补齐表格数据
                for i in range(diff):
                    row.append("") 
                    row.append("") 
            formatData.append(row)
           
        
        # print(formatData)

        sheetHead = ["办事处","客户编码"]
        for m in range(hnSupplierNum):
            sheetHead.append("红牛供应商编码%s"%(m+1))
            sheetHead.append("红牛%s"%(m+1))
        for n in range(zmSupplierNum):
            sheetHead.append("战马供应商编码%s"%(n+1))
            sheetHead.append("战马%s"%(n+1))
        for o in range(gbsSupplierNum):
            sheetHead.append("果倍爽供应商编码%s"%(o+1))
            sheetHead.append("果倍爽%s"%(o+1))    
        

        print(sheetHead)
        #写表头
        sheet.range("A1").value = sheetHead
        #写内容
        sheet.range("A2").value = formatData
        print("---------------数据处理完成---------------")
        wb.save("data/d2.xls")
        wb.close()
        app.quit()


# 处理新的表格数据
class ExcelOptNew(object):

    def __init__(self,filePath = filePath) -> None:
        super().__init__()
        self.filePath = filePath
        self.app = xw.App(visible=True, add_book=False)

    def copy(self):
        # visible 控制 Excel 打开是否显示界面
        # add_book 控制是否添加新的 workbook
        # app = xw.App(visible=True, add_book=False)
        app = self.app
        # 打开 data.xlsx 文件到 wookbook 中
        wb = self.app.books.open(filePath)
        wb.save("data/test2.xlsx")
        wb.close()
        app.quit()

    def modify(self):
        # visible 控制 Excel 打开是否显示界面
        # add_book 控制是否添加新的 workbook
        # app = xw.App(visible=True, add_book=False)
        app = self.app
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
        # app = xw.App(visible=True, add_book=False)
        app = self.app
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
        # app = xw.App(visible=True, add_book=False)
        app = self.app
        # 打开 data.xlsx 文件到 wookbook 中
        wb = app.books.open(path1)
        sheet = wb.sheets.active
        valueList = sheet[0,0].current_region.value
        print(valueList)

    #基于订单明细表、MIT订单表 提取出办事处、客户编码、订单量
    def extractData(self,path=tempPath):
        data = {}
        # visible 控制 Excel 打开是否显示界面
        # add_book 控制是否添加新的 workbook
        # app = xw.App(visible=True, add_book=False)
        app = self.app

        # 打开 data.xlsx 文件到 wookbook 中
        wb = app.books.open(path)
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
        i = -1
        for v in valueList:
            i += 1
            # print("key:%s value:%s"%key%value)
            if i>0:
                #取办事处、客户编码、订单量
                name = str(v[0]).strip()
                id = str(v[9]).strip()
                brand = str(v[33]).strip()
                num = int(v[35])
                if(BRAND != brand):
                    continue
                if data.get(id) != None:
                    #销量累加
                    data[id][1] += num
                    # print("相同客户编码：%s 在第%s行  销量已累加---"%(id,i+1))
                else:
                    value = [name,num]
                    data[id] = value
                # print("key:%s value:%s"%key%value)
            
            
        

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
        j = -1
        for v2 in valueList2:
            # print("key:%s value:%s"%key%value)
            j += 1
            if j>0:
                #取办事处、客户编码、订单量
                name2 = str(v2[0]).strip()
                id2 = str(v2[9]).strip()
                brand2 = str(v2[32]).strip()
                num2 = int(v2[34])
                # print("%s"%id2)
                if(BRAND != brand2):
                    continue
                if data.get(id2) != None:
                    #销量累加
                    data[id2][1] += num2
                    print("相同客户编码：%s 在第%s行  销量已累加---"%(id2,j+1))
                else:
                    value2 = [name2,num2]
                    data[id2] = value2
            

        # print(data)
        #创建新表
        wb.sheets.add("temp")
        # 引用
        sheet = wb.sheets["temp"]
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
        wb.save("data/d1.xls")
        wb.close()
        app.quit()

    #基于订单明细表、MIT订单表 提取出费用网点明细表
    def extractDataDetail(self,path=tempPath):
        app = self.app

        # 打开 data.xlsx 文件到 wookbook 中
        wb = app.books.open(path)
          # 切换到当前活动的 sheet 中
        sheet = wb.sheets[0]
        valueList = sheet.range("A1").current_region.value
        # valueList = sheet.range("A1:AO5").value
        print("sheet1数据条数---- %s ----"%len(valueList))
        print("sheet1列名：",valueList[0],sep="\n")
        BRAND_HN = "红牛"
        BRAND_ZM = "战马"
        BRAND_GBS = "果倍爽"
        #记录转换的所有行数据{id:rowData}
        data = {}
        #[name,id,brand,{supplierId:num}]
        # i = -1
        for i,v in enumerate(valueList):
            # i += 1
            # print("key:%s value:%s"%key%value)
            if i>0:
                #取办事处、客户编码、品牌、供应商编码、订单量、产品小计
                name = str(v[0]).strip()
                id = str(v[9]).strip()
                brand = str(v[33]).strip()
                supplierId = str(v[16]).strip()
                num = int(v[35])
                totalPrice = float(v[40])
                # print("name:%s id:%s brand:%s supplierId: %s num %s"%(name,id,brand,supplierId,num))
                #构造品牌唯一key
                key = brand+"-"+supplierId
                rowData = data.get(id)
                if rowData != None:                   
                    if rowData[2].get(key) != None:
                        #销量累加 PS:红牛取销量 战马和果倍爽取产品小计
                        if brand == BRAND_HN: 
                            rowData[2][key] += num
                        else:
                            rowData[2][key] += totalPrice
                        
                        # print("相同客户编码：%s 在第%s行  销量已累加---"%(id,i+1))
                    else:
                        #第N个供应商编码对应销量 PS:红牛取销量 战马和果倍爽取产品小计
                        if brand == BRAND_HN: 
                            rowData[2][key] = num
                        else:
                            rowData[2][key] = totalPrice
                        
                else:
                    #列数据 PS:红牛取销量 战马和果倍爽取产品小计
                    if brand == BRAND_HN: 
                        rowData = [name,id,{key:num}]
                    else:
                         rowData = [name,id,{key:totalPrice}]
                    data[id] = rowData

        # for d,v in data.items():
        #     print(v)
        #提取第二个sheet数据
        #引用第二个表单
        sheet = wb.sheets[1]
        #将所引用的表单设为活动表单
        sheet.activate
        #引用活动表单
        sheet = wb.sheets.active
        valueList2 = sheet.range("A1").current_region.value
        # valueList2 = sheet.range("A1:AI50").value
        print("sheet2数据条数---- %s ----"%len(valueList2))
        print("sheet2列名：",valueList2[0],sep="\n")   
        # j = -1
        for j,v in enumerate(valueList2):
            # print("key:%s value:%s"%key%value)
            # j += 1
         #取办事处、客户编码、品牌、供应商编码、订单量、产品小计
            if j > 0:
                name = str(v[0]).strip()
                id = str(v[9]).strip()
                brand = str(v[32]).strip()
                supplierId = str(v[16]).strip()
                num = int(v[34])
                totalPrice = float(v[39])
                #构造品牌唯一key
                key = brand+"-"+supplierId
                rowData = data.get(id)
                if rowData != None:                   
                    if rowData[2].get(key) != None:
                         #销量累加 PS:红牛取销量 战马和果倍爽取产品小计
                        if brand == BRAND_HN: 
                            rowData[2][key] += num
                        else:
                            rowData[2][key] += totalPrice
                        # print("相同客户编码：%s 在第%s行  销量已累加---"%(id,j+1))
                    else:
                        #第N个供应商编码对应销量 PS:红牛取销量 战马和果倍爽取产品小计
                        if brand == BRAND_HN: 
                            rowData[2][key] = num
                        else:
                            rowData[2][key] = totalPrice
                else:
                     #列数据 PS:红牛取销量 战马和果倍爽取产品小计
                    if brand == BRAND_HN: 
                        rowData = [name,id,{key:num}]
                    else:
                         rowData = [name,id,{key:totalPrice}]
                    data[id] = rowData


        # print(data)
        #创建新表
        wb.sheets.add("费用网点明细表")
        # 引用
        sheet = wb.sheets["费用网点明细表"]
        #将所引用的表单设为活动表单
        sheet.activate
        #引用活动表单
        sheet = wb.sheets.active
        formatData = []
        # 客户编码对应的红牛供应商编码最大数（动态的）
        hnSupplierNum = 1
        # 客户编码对应的战马供应商编码最大数（动态的）
        zmSupplierNum = 1
        # 客户编码对应的果倍爽供应商编码最大数（动态的）
        gbsSupplierNum = 1
        for k,v in data.items():
            hnSupplierNum2 = 0
            zmSupplierNum2 = 0
            gbsSupplierNum2 = 0
            for k1,v1 in v[2].items():
                brandSupplierId = k1.split("-")
                brand = brandSupplierId[0]
                if BRAND_HN == brand:
                    hnSupplierNum2 += 1
                elif BRAND_ZM == brand:
                   zmSupplierNum2 += 1
                elif BRAND_GBS == brand:
                   gbsSupplierNum2 += 1
                else:
                    print("未知品牌：%s 请校验表格数据-----"%brand)
            hnSupplierNum = max(hnSupplierNum,hnSupplierNum2)
            zmSupplierNum = max(zmSupplierNum,zmSupplierNum2)
            gbsSupplierNum = max(gbsSupplierNum,gbsSupplierNum2)
           
        print("红牛供应商最大数量：%s 战马供应商最大数量：%s 果倍爽供应商最大数量: %s"%(hnSupplierNum,zmSupplierNum,gbsSupplierNum))
        for k,v in data.items():
            row = [v[0],v[1]]
            x = 0
            y = 0
            z = 0
            for k1,v1 in v[2].items():
                brandSupplierId = k1.split("-")
                brand = brandSupplierId[0]
                supplierId = brandSupplierId[1]
                if(brand == BRAND_HN):
                    row.append(supplierId)
                    row.append(v1)
                    x+=1
            diff = hnSupplierNum-x
            if(diff>0):
                #补齐表格数据
                for i in range(diff):
                    row.append("") 
                    row.append("")      
            for k2,v2 in v[2].items():    
                brandSupplierId = k2.split("-")
                brand = brandSupplierId[0]
                supplierId = brandSupplierId[1]
                if(brand == BRAND_ZM):
                    row.append(supplierId)
                    row.append(v2)
                    y+=1
            diff = zmSupplierNum-y
            if(diff>0):
                #补齐表格数据
                for i in range(diff):
                    row.append("") 
                    row.append("")     
            for k3,v3 in v[2].items():    
                brandSupplierId = k3.split("-")
                brand = brandSupplierId[0]
                supplierId = brandSupplierId[1]
                if(brand == BRAND_GBS):
                    row.append(supplierId)
                    row.append(v3)
                    z+=1
            diff = gbsSupplierNum-z
            if(diff>0):
                #补齐表格数据
                for i in range(diff):
                    row.append("") 
                    row.append("") 
            formatData.append(row)
           
        
        # print(formatData)

        sheetHead = ["办事处","客户编码"]
        for m in range(hnSupplierNum):
            sheetHead.append("红牛供应商编码%s"%(m+1))
            sheetHead.append("红牛%s"%(m+1))
        for n in range(zmSupplierNum):
            sheetHead.append("战马供应商编码%s"%(n+1))
            sheetHead.append("战马%s"%(n+1))
        for o in range(gbsSupplierNum):
            sheetHead.append("果倍爽供应商编码%s"%(o+1))
            sheetHead.append("果倍爽%s"%(o+1))    
        

        print(sheetHead)
        #写表头
        sheet.range("A1").value = sheetHead
        #写内容
        sheet.range("A2").value = formatData
        print("---------------数据处理完成---------------")
        wb.save("data/d2.xls")
        wb.close()
        app.quit()
