#encoding=utf-8
import time,os
from datetime import datetime
from Config.Var import *

#获取当前的日期
"""
def getCurrentDate():
    timeTup = time.localtime()
    currentDate = str(timeTup.tm_year) + "-" + \
        str(timeTup.tm_mon) + "-" + str(timeTup.tm_mday)
    return currentDate
    
def getCurrentTime():
    timeStr = datetime.now()
    nowTime = timeStr.strftime('%H-%M-%S-%f')
    return nowTime
"""
def getCurrentDate():
    currentDate = time.strftime("%Y-%m-%d",time.localtime())
    return currentDate

#获取当前的时间
def getCurrentTime():
    currentTime = time.strftime("%H:%M:%S",time.localtime())
    return currentTime

#创建截图存放的目录
def createCurrentDateDir():
    dirName = os.path.join(screenPicturesDir,getCurrentDate())
    if not os.path.exists(dirName):
        os.mkdir(dirName)
    return dirName

