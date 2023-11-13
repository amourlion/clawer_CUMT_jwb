import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from chromedriver_py import binary_path
import ddddocr
import re
from selenium.webdriver.chrome.options import Options

# 获取当前工作目录
pwd = os.getcwd()

# 配置ChromeOptions以使用无头模式和设置下载路径
options = Options()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
prefs = {'download.default_directory': pwd}
options.add_experimental_option('prefs', prefs)

# 初始化浏览器（利用chromedriver_py自动检测driver版本）
driver = webdriver.Chrome(executable_path=binary_path, options=options)

# 配置Chrome浏览器以允许文件下载
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': pwd}}
command_result = driver.execute("send_command", params)

print("Headless Chrome Initiated")

# 登录页面
url = "http://jwxt.cumt.edu.cn/jwglxt/xtgl/login_slogin.html"
driver.set_window_size(1200, 800)
cookies = driver.get_cookies()
driver.get(url)

# 添加先前获取的cookies
for k in cookies:
    cookie_list = []
    driver.add_cookie(cookie_list)

# 重新加载页面
driver.get(url)

# 输入用户名和密码
username = "******"
password = "******"
username_input = driver.find_element(By.ID, "yhm")
password_input = driver.find_element(By.ID, "mm")
username_input.send_keys(username)
password_input.send_keys(password)

# 引入WebDriverWait，等待2秒
WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.NAME, 'yzmPic')))

# 获取验证码位置
element = driver.find_element(By.NAME, 'yzmPic')
element.screenshot("code.png")

# 使用OCR识别验证码
ocr = ddddocr.DdddOcr()
with open('code.png', 'rb') as f:
    img_bytes = f.read()
res = ocr.classification(img_bytes)

# 提取验证码中的数字和字母
pattern = re.compile(r'[a-zA-Z0-9]+')
matches = pattern.findall(res)
result = ''.join(matches)

# 移除可能的多余字符
if len(result) > 4:
    result = result[1:]

# 输入验证码
cap_input = driver.find_element(By.ID, "yzm")
cap_input.send_keys(result)

# 引入WebDriverWait，等待5秒
WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, 'dl'))).click()

# 引入WebDriverWait，等待5秒
WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.LINK_TEXT, "信息查询"))).click()

# 引入WebDriverWait，等待5秒
WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//a[text()='学生成绩查询']"))).click()

# 切换到新窗口
num = driver.window_handles
driver.switch_to.window(num[1])

# 引入WebDriverWait，等待3秒
WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='2023-2024']"))).click()

# 引入WebDriverWait，等待3秒
WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//li[text()='全部']"))).click()

# 引入WebDriverWait，等待3秒
WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='1']"))).click()

# 引入WebDriverWait，等待3秒
d = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH, "//li[text()='全部']")))
if len(d) > 1:
    d[1].click()

# 引入WebDriverWait，等待3秒
WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "search_go"))).click()

# 引入WebDriverWait，等待2秒
WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "btn_dc"))).click()

# 引入WebDriverWait，等待2秒
WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "btn_bcdc"))).click()

# 等待下载完成
for i in os.listdir(pwd):
    if ".crdownload" in i:
        time.sleep(1)

# 引入WebDriverWait，等待10秒
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "some_element_id")))

# 在完成后关闭浏览器
driver.quit()
