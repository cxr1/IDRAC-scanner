# coding = utf-8
from selenium import webdriver
from time import sleep
from PIL import ImageGrab
import winreg
import os


def openwindow(urllist):
    global hwnd
    for posts in range(len(urllist)):
        browser.get(urllist[posts])
        if posts != len(urllist) - 1:
            browser.execute_script("window.open('');")
            hwnd = browser.window_handles
            browser.switch_to.window(hwnd[-1])
    for i in range(len(urllist)):
        browser.switch_to.window(hwnd[i])
        login()
        sleep(8)
        sn = browser.find_element_by_xpath("//span[@id='System.Info.ServiceTag']")
        sn = sn.text
        os.makedirs(sn)
        # im = ImageGrab.grab((0, 98, 1920, 1030))  # 可以添加一个坐标元组进去
        # im.save(sn + ' ' + c[i] + u'\\仪表板.png')
        browser.get_screenshot_as_file(sn + u'\\仪表板.png')  #截图
        browser.find_element_by_xpath("//button[@class='btn btn-sm btn-primary ng-scope']").click()
        sleep(5)
        im = ImageGrab.grab((0, 0, 1350, 1000))  # 可以添加一个坐标元组进去
        im.save(sn + u'\\虚拟控制台.png')
        # sleep(5)
        # curHandle = browser.current_window_handle  # 获取当前窗口句柄
        # allHandle = browser.window_handles  # 获取所有句柄
        # for h in allHandle:
        #     if h != curHandle:
        #         browser.switch_to.window(h)  # 切换句柄，到新弹出的窗口
        #         browser.close()
        #         break
        sleep(3)
        browser.find_element_by_xpath("//strong[@id='storage']").click()
        sleep(3)
        # im = ImageGrab.grab((0, 98, 1920, 1030))  # 可以添加一个坐标元组进去
        # im.save(sn + ' ' + c[i] + u'\\存储.png')
        browser.get_screenshot_as_file(sn + u'\\存储.png')  #截图


def login():
    browser.find_element_by_name("username").send_keys('root')  # 输入用户名
    browser.find_element_by_name("password").send_keys('calvin')  # 输入密码
    browser.find_element_by_tag_name("button").click()  # 点击登录


with open('iplist.txt', 'r') as fin:  # 打开ip列表文本文档
    oldurls = fin.read().splitlines()  # 逐行读出ip地址并写入列表
s = 'https://'
d = '/restgui/start.html'
for x in range(len(oldurls)):  # 将ip地址前缀和后缀拼接到ip列表里的ip上并写入新列表
    oldurls[x] = s + oldurls[x] + d
urls = oldurls
options = webdriver.ChromeOptions()
reg=winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome")  #根据注册表找到chrome的安装路径
path=winreg.QueryValueEx(reg,'InstallLocation')
path = path[0] + "\chrome.exe"
options.binary_location = path
options.add_argument('--ignore-certificate-errors')  #关闭ssl错误提示
options.add_experimental_option('excludeSwitches', ['enable-automation'])  #关闭正在受自动化测试软件控制的提示
browser = webdriver.Chrome(chrome_options=options)  # 打开谷歌浏览器
browser.maximize_window()  # 最大化窗口
openwindow(urls)  #传入ip地址列表到主函数并运行
