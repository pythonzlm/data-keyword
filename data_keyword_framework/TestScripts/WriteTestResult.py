#encoding=utf-8
from Util.ParseExcel import *
from Config.Var import *
import traceback


excelObj=ParseExcel()
excelObj.loadWorkBook(dataFilePath)
#用例或用例步骤执行结束后，向excel中写执行结果信息
def writeTestResult(sheetObj,rowNo,colNo,testResult,errorinfo=None,picPath=None):
    #测试通过结果信息为绿色，失败为红色
    colorDict={"pass":"green","fail":"red","":None}
# 因为“测试用例”工作表和“用例步骤sheet表”中都有测试执行时间和
    # 测试结果列，定义此字典对象是为了区分具体应该写哪个工作表
    colsDict=\
        {"testCase":[testCase_runTime,testCase_testResult],
         "caseStep":[testStep_runTime,testStep_testResult],
         "dataSheet":[dataSource_runTime,dataSource_result]
         }
    try:

# 在测试步骤sheet中，写入测试结果，colsNo代表着
# testCase,testStep,dataSheet 三者之一
         excelObj.writeCell(sheetObj,content=testResult,rowNo=rowNo,colsNo=colsDict[colNo][1],
                            style=colorDict[testResult])
         if testResult=="":
    # 清空时间单元格内容
             excelObj.writeCell(sheetObj,content="",rowNo=rowNo,colsNo=colsDict[colNo][0])
         else:
    # 在测试步骤sheet中，写入测试时间
             excelObj.writeCellCurrentTime(sheetObj,rowNo=rowNo,colsNo=colsDict[colNo][0])
         if errorinfo and picPath:
    # 在测试步骤sheet中，写入异常信息
             excelObj.writeCell(sheetObj,content=errorinfo,rowNo=rowNo,colsNo=testStep_errorInfo)
    # 在测试步骤sheet中，写入异常截图路径
             excelObj.writeCell(sheetObj,content=picPath,rowNo=rowNo,colsNo=testStep_errorPic)
         else:
             if colNo=="caseStep":
                 excelObj.writeCell(sheetObj,content="",rowNo=rowNo,colsNo=testStep_errorInfo)
                 excelObj.writeCell(sheetObj,content="",rowNo=rowNo,colsNo=testStep_errorPic)
    except Exception,e:
        print u"写excel时发生异常"
        print traceback.print_exc()




