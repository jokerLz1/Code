
from datetime import datetime

from enum import Enum
import ast
#文件路径
txtUrl=r"C:\Users\L\Desktop\pointInfo.txt"

#信息关键字枚举
class NameList(Enum):
    ZSB_MAT_LIST = "ZSB_MAT_LIST"

    FC_MAT_LIST="FC_MAT_LIST"

    BALANCE_MATERIEL_LIST="BALANCE_MATERIEL_LIST"

    ACTUAL_START_DATE="ACTUAL_START_DATE"

    ACTUAL_END_DATE="ACTUAL_END_DATE"

    REQUIRE_START_DATE="REQUIRE_START_DATE"

    REQUIRE_END_DATE="REQUIRE_END_DATE"

    DESIGN_BBU_NUM_LTE="DESIGN_BBU_NUM_LTE"

    DESIGN_RRU_NUM_LTE="DESIGN_RRU_NUM_LTE"

    DESIGN_CONTROL_BOARD_NUM="DESIGN_CONTROL_BOARD_NUM"

    DESIGN_BASEBAND_BOARD_NUM="DESIGN_BASEBAND_BOARD_NUM"

    DESIGN_ANTENNAS_NUM="DESIGN_ANTENNAS_NUM"
    END_BBU_NUM_LTE="END_BBU_NUM_LTE"
    END_RRU_NUM_LTE="END_RRU_NUM_LTE"
    END_CONTROL_BOARD_NUM="END_CONTROL_BOARD_NUM"
    END_BASEBAND_BOARD_NUM="END_BASEBAND_BOARD_NUM"
    END_ANTENNAS_NUM="END_ANTENNAS_NUM"

#材料类
class Materiel:
    def __init__(self, name, quantity, type, factor):
        self.name = name
        self.quantity = quantity
        self.type = type
        self.factor = factor
    def printDetail(self):
        print(self.name + " "+self.quantity.__str__() +"个 ,型号：" +self.type+" ,品牌:"+self.factor)
class TablePoint:
    def __init__(self,startTime,endTime,requireStartTime,requireEndTime):
        self.startTime=startTime
        self.endTime=endTime
        self.requireStartTime=requireStartTime
        self.requireEndTime=requireEndTime
    def printTime(self):
        print("实际开始时间:"+self.startTime+"， 实际结束时间："+self.endTime)
#初始化文件信息
def init():
    with open(txtUrl, 'r', encoding="utf-8") as f:
        content = f.read()
    return content
# 将内容转换为字典类型
content=init()
items_dict = ast.literal_eval(content)

def initPoint():
    tablePoint=TablePoint("","","","")
    tablePoint.startTime=items_dict[NameList.ACTUAL_START_DATE.value]
    tablePoint.endTime=items_dict[NameList.ACTUAL_END_DATE.value]
    tablePoint.requireStartTime=items_dict[NameList.REQUIRE_START_DATE.value]
    tablePoint.requireEndTime=items_dict[NameList.REQUIRE_END_DATE.value]
    return tablePoint


#检查规模数据
def checkBBURRU():
    if (items_dict[NameList.DESIGN_BBU_NUM_LTE.value]!=items_dict[NameList.END_BBU_NUM_LTE.value])\
        or (items_dict[NameList.DESIGN_RRU_NUM_LTE.value]!=items_dict[NameList.END_RRU_NUM_LTE.value])\
        or (items_dict[NameList.DESIGN_CONTROL_BOARD_NUM.value]!=items_dict[NameList.END_CONTROL_BOARD_NUM.value])\
        or (items_dict[NameList.DESIGN_BASEBAND_BOARD_NUM.value]!=items_dict[NameList.END_BASEBAND_BOARD_NUM.value])\
        or(items_dict[NameList.DESIGN_ANTENNAS_NUM.value]!=items_dict[NameList.END_ANTENNAS_NUM.value]):
        raise ValueError("完工规模填写错误")
    else:
        print("完工规模填写正确")

#获取主材辅材list
def getItemList(items_dict,listName):
    itemList = []
    for x in items_dict[listName]:
        Mat = Materiel("", 0, "", "")
        Mat.name = x["MATERIEL_NAME"]
        Mat.quantity = int(x["MATERIEL_QUANTITY"])
        Mat.type = x["MATERIEL_TYPE"]
        Mat.factor=x["MANUFACTOR"]
        itemList.append(Mat)
    return itemList

#获取完工平衡表list函数
def getBalenceItemList(items_dict,listName):
    itemList = []
    for x in items_dict[listName]:
        Mat = Materiel("", 0, "", "")
        Mat.name = x["MATERIEL_NAME_DESIGN"]
        count=x["USED_QUANTITY"]
        if count=='':
            Mat.quantity=0.0
        else:
            Mat.quantity = float(count)
        Mat.type = x["MATERIEL_NAME_SCM"]
        Mat.factor = x["MATERIEL_MODEL_SCM"]
        itemList.append(Mat)
    return itemList

# 完工平衡表  全局变量
balenceList = getBalenceItemList(items_dict, NameList.BALANCE_MATERIEL_LIST.value)
#获取平衡表总数 返回总数int
def getBalenceListAll():
    count=0.0
    for item in balenceList:
        count=count+item.quantity
    return count


def timeCheck(tablePoint):
    print("开工完工时间检查:",end="")
    actual_start_time=datetime.strptime(tablePoint.startTime, '%Y-%m-%d').strftime('%Y-%m-%d')
    require_start_time=datetime.strptime(tablePoint.requireStartTime, '%Y-%m-%d').strftime('%Y-%m-%d')
    actual_end_time=datetime.strptime(tablePoint.endTime,'%Y-%m-%d').strftime('%Y-%m-%d')
    require_end_time=datetime.strptime(tablePoint.requireEndTime,'%Y-%m-%d').strftime('%Y-%m-%d')
    today=datetime.now().strftime('%Y-%m-%d')
    if actual_end_time!=today:
        raise ValueError("完工日期错误")
    if actual_start_time<require_start_time:
        raise ValueError("实际开工日期早于要求开工日期，需要修改工单")
    if(actual_end_time>require_end_time):
        raise ValueError("实际完工日期晚于要求开工日期，需要申请延期")
    print("开工完工时间正确")

#得到主材表list
def getZSBList()->list:
    return getItemList(items_dict,NameList.ZSB_MAT_LIST.value)






