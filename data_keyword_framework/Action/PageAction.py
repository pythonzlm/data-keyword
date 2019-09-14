#encoding=utf-8
from selenium import webdriver
from Config.Var import *
from Util.ObjectMap import *
from Util.ClipboardUtil import *
from Util.KeyBoardUtil import *
from Util.DirAndTime import *
from Util.WaitUtil import *
from selenium.webdriver.chrome.options import Options
import time

#定义全局driver变量
driver=None
#全局的等待类实例对象
waitUtil = None

def open_browser(browserName,*args):
    global driver,waitUtil
    try:
        if browserName.lower()=="ie":
            driver = webdriver.Ie(executable_path=ieDriverFilePath)
        elif browserName.lower()=="chrome":
            #创建chrome浏览器的一个Options实例对象
            chrome_options = Options()
            #添加屏蔽-ignore-certificate-errors提示信息的设置参数项
            chrome_options.add_experimental_option("excludeSwitches",["ignore-certificate-errors"])
            driver = webdriver.Chrome(executable_path=chromeDriverFilePath
                                      ,chrome_options=chrome_options)
        else:
            driver = webdriver.Firefox(executable_path=firefoxDriverFilePath)
    #driver 对象创建成功后，创建等待类实例对象
        waitUtil = WaitUtil(driver)
    except Exception,e:
        raise e

def visit_url(url,*arg):
    #访问某个网址
     global driver
     try:
         driver.get(url)
     except Exception,e:
         raise e

def close_browser(*arg):
    global driver
    try:
        driver.quit()
    except Exception,e:
        raise e

def sleep(sleepSeconds,*arg):
    #强制等待
    try:
        time.sleep(int(sleepSeconds))
    except Exception,e:
        raise e

def clear(locateType,locateExpression,*arg):
    #清除输入框默认内容
    global driver
    try:
        getElement(driver,locateType,locateExpression).clear()
    except Exception,e:
        raise e

def input_string(locateType,locateExpression,inputContent):
    global driver
    try:
        getElement(driver,locateType,locateExpression).send_keys(inputContent)
    except Exception,e:
        raise e

def click(locateType,locateExpression,*arg):
    global driver
    try:
        getElement(driver,locateType,locateExpression).click()
    except Exception,e:
        raise e

def assert_string_in_pagesource(assertString,*arg):
    #断言页面源码是否存在关键字
    global driver
    try:
        assert assertString in driver.page_source,\
        u"%s not found in page source!"%assertString
    except AssertionError,e:
        raise AssertionError(e)
    except Exception,e:
        raise e

def assert_title(titleString,*args):
    global driver
    try:
        assert titleString in driver.title,\
        u"% not found in title!" %titleString
    except AssertionError,e:
        raise AssertionError(e)
    except Exception,e:
        raise e

def getTitle(*arg):
    #获取页面标题
    global driver
    try:
        return driver.title
    except Exception,e:
        raise e

def getPageSource(*arg):
    #获取页面源码
    global driver
    try:
        return driver.page_source
    except Exception,e:
        raise e

def switch_to_frame(locateType,frameLocatorExpression,*arg):
    global driver
    try:
        driver.switch_to_frame(getElement(driver,locateType,frameLocatorExpression))
    except Exception,e:
        raise e

def switch_to_default_content(*arg):
    global driver
    #退出frame，回到默认对话框中
    try:
        driver.switch_to.default_content()
    except Exception,e:
        raise e

def paste_string(pasteString,*arg):
    #模拟ctrl+v操作
    try:
        Clipboard.setText(pasteString)
        time.sleep(2)
        KeyboardKeys.twoKeys("ctrl","v")
    except Exception,e:
        raise e

def press_tab_key(*arg):
    #模拟tab键
     try:
         KeyboardKeys.oneKey("tab")
     except Exception,e:
         raise e

def press_enter_key(*arg):
    #模拟enter键
     try:
         KeyboardKeys.oneKey("enter")
     except Exception,e:
         raise e

def maximize_browser():
    global driver
    try:
        driver.maximize_window()
    except Exception,e:
        raise e

def capture_screen(*args):
    global driver
    current_time=getCurrentTime()
    picturePath = str(createCurrentDateDir())+"\\"+str(current_time)+".png"
    try:
        #r"\\"为了防止字符转义
        driver.get_screenshot_as_file(picturePath.replace('\\',r'\\'))
    except Exception,e:
        raise e
    else:
        return picturePath

def waitPresenceOfElementLocated(locatorType,locatorExpression,*arg):
    """
    显示等待页面元素出现在DOM中，但并一定可以见，
            存在则返回该页面元素对象
    """
    global waitUtil
    try:
        waitUtil.presenceOfElementLocated(locatorType,locatorExpression)
    except Exception,e:
        raise e

def waitFrameToBeAvailableAndSwitchToIt(locateType,locateExpression):
    """检查frame是否存在，存在则切换进frame控件中"""
    global waitUtil
    try:
         waitUtil.frameToBeAvailableAndSwitchToIt(locateType,locateExpression)
    except Exception,e:
        raise e

def waitVisibilityOfElementLocated(locateType,locateExpression,*args):
    '''显示等待页面元素出现在DOM中，并且可见，存在返回该页面元素对象'''
    global waitUtil
    try:
        waitUtil.visibilityOfElementLocated(locateType,locateExpression)
    except Exception,e:
        raise e








