from xlrd import Book

from utils.excelCheckUtil import Sheet, Point, init, checkSheet5, initialize_objects, Sheet3, Sheet4, Sheet5, Sheet6, \
    checkSheet3, dealDesignSheet

# from utils.tableInfoCheckUtil import getBalenceListAll


def getResStr(xlsUrl,designUrl):
    wb: Book = init(xlsUrl)

    desigonWb: Book = init(designUrl)
    # 初始化
    point, sheet3obj, sheet4obj, sheet5obj, sheet6obj, desigonSheet = Point("", "", ""), Sheet3("", 0.0, 0), Sheet4("",
                                                                                                                    0.0,
                                                                                                                    0), Sheet5(
        "", 0.0, 0), Sheet6("", 0.0, 0), Sheet("", 0.0, 0)
    initialize_objects(wb, point, sheet3obj, sheet4obj, sheet5obj)

    desigonSheet = dealDesignSheet(desigonWb, desigonSheet)

    # 站点基本信息展示
    resStr="----------------------------站点基本信息------------------------------\n"
    resStr=resStr+point.getPointDetail()+"\n"
    # 表三甲
    resStr=resStr+"----------------------------表三甲比对--------------------------------\n"
    resStr=resStr+sheet3obj.getDetail()
    resStr=resStr+checkSheet3(sheet3obj, desigonSheet)
    # 表四甲
    resStr = resStr +"----------------------------表四甲比对--------------------------------\n"
    strAll = sheet4obj.strAll
    count = 0
    lineList = []
    for x in strAll:
        if x.find("普通护套") != -1:
            lineList.append(x)
            count = count + strAll[x]
    resStr = resStr +"馈线有:" + str(lineList) + "， 总数为：" + str(count)+"\n"
    # overALl = float(getBalenceListAll())
    sheet4All = float(sheet4obj.total)
    resStr = resStr +"表四甲和完工平衡表分别为：\n"+str(sheet4All)+"\n"+str(overALl)+"\n"


    # 表四乙
    resStr=resStr+"----------------------------表四乙比对--------------------------------\n"
    targetItems = ["空气开关", "室内吸顶天线支架"]
    resStr=resStr+ checkSheet5(targetItems, sheet5obj)+"\n"

    return  resStr



