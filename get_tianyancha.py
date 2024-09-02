import openpyxl
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import random
 
# 设置 Edge 选项以启动最大化窗口
options = Options()
options.add_argument("--start-maximized")

options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-extensions")
options.add_argument("--disable-popup-blocking")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--remote-debugging-port=9222")


custom_user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
options.add_argument(f'--user-agent={custom_user_agent}')
 
# 获取查询名单
list_name = []
path = r'./企业名称和信用代码.xlsx'
 
wb = openpyxl.load_workbook(path)
wb_sheet = wb['Sheet1']
maxrows = wb_sheet.max_row
for i in range(maxrows - 1):
    name = wb_sheet.cell(i + 2, 1).value
    list_name.append(name)
 
# 初始化 Edge Driver
edge_service = EdgeService(EdgeChromiumDriverManager().install())
driver = webdriver.Edge(service=edge_service, options=options)
url = 'https://www.tianyancha.com/'
driver.get(url)
driver.refresh()

 
# 延时，手动扫码登录
sleep(30)
 
cnt = 0
for j in list_name:
    cnt += 1
    driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div/div[3]/div[2]/div[1]/div[1]/input').clear()  # 定位到搜索框
    driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div/div[3]/div[2]/div[1]/div[1]/input').send_keys(j)  # 在搜索框中输入查询企业名单
    try:
        driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div/div[3]/div[2]/div[1]/button').click()
    except:
        driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div/div[3]/div[2]/div[1]/button').click()
    try:
        name_id = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[2]/div/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div[4]/span').text
    except:
        name_id = "根据商家名称匹配不到数据"
    print(cnt, j, name_id)
    # 写入商家社会信用代码
    wb_sheet.cell(list_name.index(j) + 2, 2, value=name_id)
 
    # 随机暂缓，防止检测到异常
    sleep(random.uniform(0, 3))
    # 跳转回首页，无需再次登录
    try:
        driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/a').click()
    except:
        sleep(10)
        print("errors")
wb.save(path)
wb.close()
driver.close()
