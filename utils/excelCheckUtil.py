
import xlrd
from datetime import datetime, timedelta


class Point:
  def __init__(self, name,Id,date):
    self.name = name  #站点名称
    self.Id = Id      #站点号
    self.date=date    #开工日期

  def getPointDetail(self):
      return "该站点为 " + self.name+" ,站号: "+self.Id+",开始时间: "+self.date

class Sheet:
    def __init__(self, strAll, total,count):
        self.strAll = strAll    #所有条目详细
        self.total =total       #数量
        self.count=count        #条数
    def getDetail(self):
        return "共" + str(self.count) + "项" + ",总计" + self.total+"\n"
class Sheet3(Sheet):
    def getDetail(self):
        return "表三甲共" + str(self.count) + "项" + ",总计" + self.total+"\n"
class Sheet4(Sheet):
    def getDetail(self):
        return "表四甲共" + str(self.count) + "项" + ",总计" + self.total+"\n"
class Sheet5(Sheet):
    def getDetail(self):
        return "表四乙共" + str(self.count) + "项" + ",总计" + self.total+"\n"
class Sheet6(Sheet):
    def getDetail(self):
        return "完工平衡表共" + str(self.count) + "项" + ",总计" + self.total+"\n"

def init(xlsUrl):
    wb = xlrd.open_workbook(xlsUrl)
    return wb
def initSheet(wb, sheetObj: Sheet, num):
    if num==3:
        sheet=wb.sheet_by_name('表三甲')
        begin=num+1
    elif num==4:
        sheet=wb.sheet_by_name('表四甲供')
        begin = num + 1
    elif num==5:
        sheet=wb.sheet_by_name("表四乙供")
        begin = num


    rows=sheet.nrows
    resRows = rows - 3
    if num==3:
        resRows=resRows+1
    for x in range(rows):
        if sheet.row(x)[1].value == "":
            resRows = resRows - 1
    total = 0.00
    strAll=dict()
    for x in range(begin, resRows + begin):
        if num==3:
            strAll[sheet.row(x)[2].value]=sheet.row(x)[4].value
            total=total+sheet.row(x)[4].value
        elif num==4:
            strAll[sheet.row(x)[1].value] = sheet.row(x)[4].value
            total = total + sheet.row(x)[4].value
        else:
            strAll[sheet.row(x)[1].value]=sheet.row(x)[5].value
            total=total+sheet.row(x)[4].value
    sheetObj.strAll=strAll
    sheetObj.count=resRows
    totalStr="{:.2f}".format(total)
    sheetObj.total=totalStr
    return sheetObj

def dealSheet1(wb,point: Point):
    sheet1 = wb.sheet_by_name('表一')

    point.name =str(sheet1.row(1)[1].value)
    if(point.name.find(" ")!=-1):
        raise ValueError("工程名称存在空格")
    point.Id =str(sheet1.row(2)[1].value)
    if (point.Id.find(" ") != -1):
        raise ValueError("工程编号存在空格")
    date_num =sheet1.row(3)[1].value
    date_obj = datetime(1899, 12, 30) + timedelta(days=date_num)
    # 将日期对象转换为字符串
    date_str = date_obj.strftime('%Y-%m-%d')
    point.date=date_str
    return point
def dealOverSheet(overWb,sheetObj: Sheet):
    sheet=overWb.sheet_by_name("完工工程物资平衡表")
    rows = sheet.nrows
    resRows = rows - 1
    total = 0.00
    for x in range(1, resRows + 1):
        # print(sheet.row(x)[11].value)
        total = total + float(sheet.row(x)[11].value)
    totalStr = "{:.2f}".format(total)
    sheetObj.count=resRows
    sheetObj.total=totalStr
    return sheetObj
def dealDesignSheet(designWb,sheetObj: Sheet):
    sheet=designWb.sheet_by_name("表三(甲)")
    rows = sheet.nrows
    resRows = rows - 1
    strAll = dict()
    for x in range(1, resRows + 1):
        # print(sheet.row(x)[11].value)
        strAll[sheet.row(x)[2].value]=float(sheet.row(x)[4].value)
    sheetObj.count=resRows
    sheetObj.strAll=strAll
    return sheetObj


def checkSheet3(sheet3Obj: Sheet,designSheetObj: Sheet):
    errmsg=[]
    designDict=designSheetObj.strAll
    sheet3Dict=sheet3Obj.strAll
    for x in designDict:
        if x not in sheet3Dict:
            errmsg.append(x)
    resStr="与设计表比对共"+str(len(errmsg))+"条不同，分别是\n"
    resStr=resStr+errmsg.__str__()+"\n"
    return resStr

def checkSheet5(targetItems,sheetObj: Sheet):
    items=sheetObj.strAll
    resStr=""
    for x in targetItems:
        if x in items:
            resStr=resStr+"有"+x+", 价格是"+str(items[x])+"\n"
    return resStr
def initialize_objects(wb,point, sheet3obj, sheet4obj, sheet5obj):
    point=dealSheet1(wb,point)
    sheet3obj=initSheet(wb, sheet3obj, 3)
    sheet4obj=initSheet(wb, sheet4obj, 4)
    sheet5obj=initSheet(wb, sheet5obj, 5)










