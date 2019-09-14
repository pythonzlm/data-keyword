#encoding=utf-8
from selenium.webdriver.support.ui import WebDriverWait

def getElement(driver,locateType,locateExpression):
    try:
        element = WebDriverWait(driver,30).until\
             (lambda x:x.find_element(by=locateType,value=locateExpression))
        return element
    except Exception,e:
        raise e
def getElements(driver,locateType,locateExpression):
    try:
        elements = WebDriverWait(driver,30).until\
            (lambda x:x.find_element(by=locateType,value=locateExpression))
        return elements
    except Exception,e:
        raise e
