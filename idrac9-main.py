# coding = utf-8
from time import sleep
from PIL import ImageGrab
import re
import os
import sys
import zipfile
import winreg
import requests
import tkinter.messagebox #弹窗库
# from selenium import webdriver
# from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from seleniumwire import webdriver

# from seleniumwire.webdriver.common.desired_capabilities import DesiredCapabilities
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option('detach', True)
chrome_options.add_experimental_option('w3c', False)
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
# chrome_options.add_argument('--headless')
# chrome_options.add_argument('--disable-gpu')
# chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument("--auto-open-devtools-for-tabs")

# d = DesiredCapabilities.CHROME
# d['loggingPrefs'] = { 'performance':'ALL' }

url = 'http://npm.taobao.org/mirrors/chromedriver/'  # chromedriver download link


def get_path():
    return os.path.dirname(os.path.realpath(sys.argv[0]))
    # return os.path.dirname(os.path.realpath(__file__))


def get_Chrome_version():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Google\Chrome\BLBeacon')
    version, types = winreg.QueryValueEx(key, 'version')
    return version


def get_server_chrome_versions():
    '''return all versions list'''
    versionList = []
    url = "http://npm.taobao.org/mirrors/chromedriver/"
    rep = requests.get(url).text
    result = re.compile(r'\d.*?/</a>.*?Z').findall(rep)
    for i in result:
        version = re.compile(r'.*?/').findall(i)[0]  # 提取版本号
        versionList.append(version[:-1])  # 将所有版本存入列表
    return versionList


def download_driver(download_url):
    '''下载文件'''
    file = requests.get(download_url)
    with open("chromedriver.zip", 'wb') as zip_file:  # 保存文件到脚本所在目录
        zip_file.write(file.content)
        print('下载成功')


def download_lase_driver(download_url, chromeVersion, chrome_main_version):
    '''更新driver'''
    versionList = get_server_chrome_versions()
    if chromeVersion in versionList:
        download_url = f"{url}{chromeVersion}/chromedriver_win32.zip"
    else:
        for version in versionList:
            if version.startswith(str(chrome_main_version)):
                download_url = f"{url}{version}/chromedriver_win32.zip"
                break
        if download_url == "":
            print("暂无法找到与chrome兼容的chromedriver版本，请在http://npm.taobao.org/mirrors/chromedriver/ 核实。")

    download_driver(download_url=download_url)
    path = get_path()
    print('当前路径为：', path)
    unzip_driver(path)
    os.remove("chromedriver.zip")
    dri_version = get_version()
    if dri_version == 0:
        return 0
    else:
        print('更新后的Chromedriver版本为：', dri_version)


def get_version():
    '''查询系统内的Chromedriver版本'''
    outstd2 = os.popen('chromedriver --version').read()
    try:
        out = outstd2.split(' ')[1]
    except:
        return 0
    return out


def unzip_driver(path):
    '''解压Chromedriver压缩包到指定目录'''
    f = zipfile.ZipFile("chromedriver.zip", 'r')
    for file in f.namelist():
        f.extract(file, path)


def check_update_chromedriver():
    try:
        chromeVersion = get_Chrome_version()
    except:
        print('未安装Chrome，请在GooGle Chrome官网：https://www.google.cn/chrome/ 下载。')
        return 0

    chrome_main_version = int(chromeVersion.split(".")[0])  # chrome主版本号

    try:
        driverVersion = get_version()
        driver_main_version = int(driverVersion.split(".")[0])  # chromedriver主版本号
    except:
        print('未安装Chromedriver，正在为您自动下载>>>')
        download_url = ""
        if download_lase_driver(download_url, chromeVersion, chrome_main_version) == 0:
            return 0
        driverVersion = get_version()
        driver_main_version = int(driverVersion.split(".")[0])  # chromedriver主版本号

    download_url = ""
    if driver_main_version != chrome_main_version:
        print("chromedriver版本与chrome浏览器不兼容，更新中>>>")
        if download_lase_driver(download_url, chromeVersion, chrome_main_version) == 0:
            return 0
    else:
        print("chromedriver版本已与chrome浏览器相兼容，无需更新chromedriver版本！")

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
        try:
            os.makedirs(sn)
        except OSError:
            tkinter.messagebox.showerror('错误','文件夹已存在')
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
check_update_chromedriver()
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
