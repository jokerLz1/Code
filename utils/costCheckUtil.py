
import xlrd

#获取物资总金额
def itemCostCheck(itemUrl,pointName):
    wb = xlrd.open_workbook(itemUrl)
    sheet = wb.sheet_by_name("单项工程用料存量统计报表")
    rows = sheet.nrows
    res = rows - 1
    itemsCost = 0.0
    count = 0
    for x in range(rows):
        if sheet.row(x)[1].value == "":
            res = res - 1
        else:
            if sheet.row(x)[5].value.find(pointName[:-2])!=-1 :
                count += 1
                itemsCost = itemsCost + float(sheet.row(x)[13].value)
    print(pointName+"的器件条数为 ："+str(count))
    return itemsCost
#获取品牌信息总表
def getBrandList(itemUrl,pointName):
    wb = xlrd.open_workbook(itemUrl)
    sheet = wb.sheet_by_name("单项工程用料存量统计报表")
    rows = sheet.nrows
    res = rows - 1
    itemList = list()
    for x in range(rows):
        if sheet.row(x)[1].value == "":
            res = res - 1
        else:
            if sheet.row(x)[5].value == pointName:
                itemList.append((sheet.row(x)[9].value,sheet.row(x)[6].value))
    return itemList
#具体品牌搜索   itemKeyword：需求物品关键字,itemlist:品牌信息总表
def brandSerch(itemKeyword,itemList):
    resStr=""
    for x in itemList:
        for y in itemKeyword:
            if x[0].find(y) != -1:
                resStr=resStr+"\""+x[0] +"\""+ " 的品牌是：" + x[1]+"\n"
    return resStr

