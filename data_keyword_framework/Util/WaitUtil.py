#encoding=utf-8
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class WaitUtil(object):
    #映射定位方式的字典对象
     def __init__(self,driver):
         self.locationTypeDict={
             "xpath":By.XPATH,
             "id":By.ID,
             "name":By.NAME,
             "css_selector":By.CSS_SELECTOR,
             "clss_name":By.CLASS_NAME,
             "tag_name":By.TAG_NAME,
             "link_text":By.LINK_TEXT,

             "partial_link_text":By.PARTIAL_LINK_TEXT
         }
    #初始化driver对象
         self.driver = driver
         #创建显示等待实例对象
         self.wait = WebDriverWait(self.driver,30)

     def presenceOfElementLocated(self,locatorType,locateExpression,*args):
         """
         显示等待页面元素出现在DOM中，但不一定可以见，存在则返回该页面元素对象
         """
         try:
             if self.locationTypeDict.has_key(locatorType.lower()):
                 element = self.wait.until\
                           (EC.presence_of_element_located((self.locationTypeDict[locatorType.lower()],locateExpression)))
                 return element
             else:
                 raise TypeError(u"未找到定位方式，请确认定位方法是否写正确")
         except Exception,e:
             raise e

     def frameToBeAvailableAndSwitchToIt(self,locateType,locateExpression,*args):
         """
        检查frame是否存在，存在则切换近frame控件
         """
         try:
             self.wait.until\
                 (EC.frame_to_be_available_and_switch_to_it((self.locationTypeDict[locateType.lower()],locateExpression)))
         except Exception,e:
             raise e

     def visibilityOfElementLocated(self,locateType,locateExpression):
         """
      显示等待页面元素出现在DOM中，并且可见，存在返回该页面元素对象
         """
         try:
             element = self.wait.until\
                 (EC.presence_of_element_located((self.locationTypeDict[locateType.lower()],locateExpression)))
             return element
         except Exception,e:
             raise e


if __name__=="__main__":
    from selenium import webdriver
    from Config.Var import *
    driver = webdriver.Chrome(executable_path=chromeDriverFilePath)
    driver.get("http://mail.163.com")
    waitUtil = WaitUtil(driver)
    waitUtil.frameToBeAvailableAndSwitchToIt('xpath',"//div[@id='loginDiv']/iframe")
    driver.quit()