from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import selenium.webdriver.remote.webelement
from selenium.webdriver.common.by import By
# # from msedge.selenium_tools import EdgeOptions
# # from msedge.selenium_tools import Edge
import time,configparser,os,winreg,platform
import pandas as pd

hot_update=configparser.RawConfigParser() 
hot_update.read(r"bin\\hot_update.ini",encoding="utf-8")
arch=platform.architecture()[0]

_browser_regs={
    "IE":r"SOFTWARE\\Clients\\StartMenuInternet\\IEXPLORE.EXE\\DefaultIcon",
    "chrome":r"SOFTWARE\\Clients\\StartMenuInternet\\Google Chrome\\DefaultIcon",
    "edge":r"SOFTWARE\\Clients\\StartMenuInternet\\Microsoft Edge\\DefaultIcon",
    "firefox":r"SOFTWARE\\Clients\\StartMenuInternet\\FIREFOX.EXE\\DefaultIcon",
    "360":r"SOFTWARE\\Clients\\StartMenuInternet\\360Chrome\\DefaultIcon",
}

def get_browser_path(browser):
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,_browser_regs.get(browser))
    except FileNotFoundError:
        return None
    value,_type = winreg.QueryValueEx(key,"")
    #return the path of browser
    print(value)
    return value.split(",")[0]

class AsseccSubmit(object):
    def __init__(self, city, user, pwd, select, path, is_submit=True):
        if get_browser_path("edge") and arch=="64bit":
            self.driver = webdriver.Edge(r"bin\\msedgedriver.exe")
        elif get_browser_path("edge"):
            self.driver = webdriver.Edge(r"bin\\msedgedriver_x86.exe")
        elif get_browser_path("chrome"):
            #https://chromedriver.storage.googleapis.com/index.html
            self.driver = webdriver.Chrome(r"bin\\chromedriver.exe")#only supports Chrome version 114
        elif get_browser_path("firefox"):
            #https://github.com/mozilla/geckodriver/releases
            self.driver = webdriver.Firefox(r"bin\\geckodriver.exe")
        elif get_browser_path("360"):
            option=webdriver.ChromeOptions().binary_location(get_browser_path("360"))
            self.driver = webdriver.Chrome(r"bin\\chromedriver86.0.4240.22.exe",chrome_options=option)
        # elif get_browser_path("IE") and arch=="64bit" and system not win11?:
        #     #https://selenium-release.storage.googleapis.com/index.html
        #     self.driver = webdriver.Firefox(r"bin\\IEDriverServer.exe")
        # elif get_browser_path("IE"):
        #     self.driver = webdriver.Firefox(r"bin\\IEDriverServer_x86.exe")
        else:
            #未检测到最新的浏览器之类的
            return 0
        self.driver.maximize_window()
        self.driver.implicitly_wait(10)
        self.city = city
        self.user = user
        self.pwd = pwd
        self.path = path
        self.city_index = None
        self.is_submit = is_submit
        self.is_login = False
        self.opts = self.get_option(select)

    def get_option(self, select):
        options = {
            "正常填报": {
                "url": hot_update.get("upload_regular_path","url"),
                "weight1": hot_update.get("upload_regular_path","weight1"),
                "weight2": hot_update.get("upload_regular_path","weight2"),
                "explain": hot_update.get("upload_regular_path","explain"),
                "save_btn":hot_update.get("upload_regular_path","save_btn"),
            },
            "衔接填报": {
                "url": hot_update.get("upload_join_path","url"),
                "weight1": hot_update.get("upload_join_path","weight1"),
                "weight2": hot_update.get("upload_join_path","weight2"),
                "explain": hot_update.get("upload_join_path","explain"),
                "save_btn":hot_update.get("upload_join_path","save_btn"),
            }
        }
        return options[select]

    def get_login_url(self):
        return hot_update.get("upload_url","login_url")

    def pass_safe_check(self, sleep_time=1):
        self.driver.find_element(By.ID, "details-button").click()
        self.driver.find_element(By.ID, "proceed-link").click()
        time.sleep(sleep_time)

    def login(self, sleep_time=2):
        print(self.get_login_url())
        self.driver.get(self.get_login_url())  # 访问百度
        self.pass_safe_check()
        self.driver.find_element(By.XPATH,hot_update.get("upload_url","usr")).send_keys(self.user)  # 填入账号
        self.driver.find_element(By.XPATH,hot_update.get("upload_url","pwd")).send_keys(self.pwd)  # 填入密码
        self.driver.find_element(By.XPATH,hot_update.get("upload_url","login_btn")).click()  # 点击登陆
        self.is_login = True
        time.sleep(sleep_time)

    # def choose_season(self,):
    #     pass

    # 获取城市序号
    def get_index(self):
        citys = self.driver.find_elements(By.XPATH,hot_update.get("upload_url","citys"))  # 填入账号
        i = 0
        for n, city in enumerate(citys):
            if city.text == self.city:
                i = n + 1
                break
        return i

    Deafult_Weight = ("0.5", "0.5", "算术平均值")

    def set_weight(self):
        all_file = os.listdir(self.path)
        weight1 = "0.5"
        weight2 = "0.5"
        explain = u"算术平均值"
        if u"权重统计表.xlsx" in all_file:
            weight_path = os.path.join(self.path,u"权重统计表.xlsx")
            weight_sheet = pd.read_excel(weight_path,sheet_name="weight",usecols="A:D",header=1)
            #get 宗地编号(str)
            stander_id = "0"
            stander_group = weight_sheet.loc[weight_sheet["[1]"]==stander_id]
            weight1_pd = stander_group["[2]"]
            weight2_pd = stander_group["[3]"]
            explain_pd = stander_group["[4]"]
            if not (weight1_pd is None or weight2_pd is None or explain_pd is None):
                weight1 = weight1_pd[0]
                weight2 = weight2_pd[0]
                explain = explain_pd[0]
        self.Deafult_Weight = (weight1, weight2, explain)

    def tb_result(self, tb_content):
        self.driver.get(self.opts["url"])
        self.city_index = self.get_index()
        self.driver.find_element(By.XPATH,hot_update.get("upload_url","tb_result") % (self.city_index)).click()
        if "weight" in tb_content:
            tb_buttons = self.driver.find_elements(By.XPATH,hot_update.get("upload_url","tb_buttons"))
            for btn in tb_buttons:
                btn.click()
                time.sleep(0.5)
                weight1, weight2, explain = self.Deafult_Weight
                self.tb_weight(weight1, weight2, explain, self.opts)
        time.sleep(1)

    # 填写权重和说明
    def tb_weight(self, weight1, weight2, explain, opt, ):
        self.driver.find_element(By.XPATH, opt["weight1"]).send_keys(weight1)
        self.driver.find_element(By.XPATH, opt["weight2"]).send_keys(weight2)
        explain_input = self.driver.find_element(By.XPATH, opt["explain"])
        explain_input.clear()
        explain_input.send_keys(explain)
        self.driver.find_element(By.XPATH, opt["save_btn"]).click()
        time.sleep(2)

    def submit(self):
        if self.is_submit:
            self.driver.get(self.opts["url"])
            time.sleep(1)
            # 勾选城市//*[@id="root"]/div/div/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[1]/div[2]/table/tbody/tr[2]/td[1]
            self.driver.find_element(By.XPATH,hot_update.get("upload_url","submit_click") % (self.city_index)).click()
            # 点击提交
            self.driver.find_element(By.XPATH,hot_update.get("upload_url","submit")).click()

    def quit(self):
        self.submit()
        self.driver.quit()


def main(city, user, pwd, select, path, is_submit):
    obj = AsseccSubmit(city, user, pwd, select, path, is_submit)
    obj.login()
    obj.tb_result(["weight"])
    obj.quit()


# if __name__ == '__main__':
#     sblx = "正常填报"
#     wb = xlrd.open_workbook("新建 XLS 工作表.xls")
#     ws = wb.sheet_by_name("Sheet1")
#     city = "武汉市"
#     pwd = "djjc_1234"
#     fail_list = []
#     for i in range(1, ws.nrows):
#         row = ws.row_values(i)
#         if row[2] == "是":
#             print(row)
#             try:
#                 main(city, row[1], pwd, sblx, True)
#                 print("%s成功" % row[0])
#             except:
#                 print("%s失败" % row[0])
#                 fail_list.append(row[0])
#             time.sleep(1)
#     print(fail_list)
