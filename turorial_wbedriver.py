from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pymssql
import re
from selenium.common.exceptions import NoSuchElementException

path = "/Users/woody/Documents/python/practise_2/chromedriver_mac_arm64/chromedriver"
service = Service(path)
driver = webdriver.Chrome(service=service)

# 創建空的列表
data_list = []

# 連接到 SQL Server 資料庫
conn = pymssql.connect(server='localhost', user='sa', password='Passw0rd', database='Movie')

# 創建游標物件
cursor = conn.cursor()

# 執行SQL查詢統編
cursor.execute('SELECT 統一編號 from company') 

# 擷取查詢結果
results = cursor.fetchall()

# 將待查詢統編結果附加到列表
for row in results:
    # 移除括號並只保留數字部分
    #num = int(row[0].strip('()'))
    num = row[0].replace('(', '').replace(')', '')

    # 確保統一編號的領導0不被去除
    統一編號 = str(num).zfill(8)
    data_list.append(統一編號)
print("統編：",data_list)

# 商工登記公示資料查詢網址
web_url = "https://findbiz.nat.gov.tw/fts/query/QueryBar/queryInit.do?disj=51849B1B295172D6ADFDD31EF5A6A597&fhl=zh_TW"

#開啟瀏覽器
driver.get(web_url)

# 創建空的結果列表
result_list = []

#點選資料種類(全選)
wait = WebDriverWait(driver, 1)
datatypeclick1 = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="queryListForm"]/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[3]')))
datatypeclick2 = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="queryListForm"]/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[5]')))
datatypeclick3 = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="queryListForm"]/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[7]')))
datatypeclick4 = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="queryListForm"]/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[9]')))  
datatypeclick1.click()
datatypeclick2.click()
datatypeclick3.click()
datatypeclick4.click()

count = 0
#開始查詢
for i in data_list:
    search = driver.find_element(By.NAME, "qryCond")
    search.clear()
    search.send_keys(str(i))
    search.send_keys(Keys.RETURN)
    print("統編:",i)
    # print(data.text)
    time.sleep(3)
    
    elements = driver.find_elements(By.XPATH, '//*[@id="queryListForm"]/div[3]/div/div/div/div/div[1]/div[1]')
    for element in elements:
        text = element.text
        match = re.search(r'共 (\d+) 筆', text)
        if match:
            count = int(match.group(1))
            print("共有幾筆：", count)           
    for count in range(1, count + 1):
        try:
            wait = WebDriverWait(driver, 1)
            #公司名稱
            #data_name_xpath = '//*[@id="vParagraph"]/div[1]/div[1]/a'
            data_name_xpath = '//*[@id="vParagraph"]/div['+ str(count) +']/div[1]/a'
            print("公司名稱",data_name_xpath)
            data_name = wait.until(EC.visibility_of_element_located((By.XPATH, data_name_xpath)))

            #公司細項
            #data_detail_xpath = '//*[@id="vParagraph"]/div[1]/div[2]'
            data_detail_xpath = '//*[@id="vParagraph"]/div['+str(count)+']/div[2]'
            print("公司細項",data_detail_xpath)
            data_detail = wait.until(EC.visibility_of_element_located((By.XPATH, data_detail_xpath)))

            #result_list.append(data.text)
            result_list.append(data_name.text + " " + data_detail.text)

        except NoSuchElementException:
            result_list.append("資料不存在")

        except Exception as e:
            result_list.append("爬取資料出現異常：" + str(e))   

# 提取最大核准變更日期的記錄
result_dict = {}

for data in result_list:
    match = re.search(r'核准變更日期：(\d{7})', data)
    if match:
        核准變更日期 = match.group(1)
        match = re.search(r'統一編號：(\d{8})', data)
        if match:
            統一編號 = match.group(1)
            if 統一編號 not in result_dict:
                result_dict[統一編號] = {'data': data, '最大核准變更日期': 核准變更日期}
            else:
                if 核准變更日期 > result_dict[統一編號]['最大核准變更日期']:
                    result_dict[統一編號] = {'data': data, '最大核准變更日期': 核准變更日期}

# 提取結果列表
result_list_filtered = [result_dict[統一編號]['data'] for 統一編號 in result_dict]

# 將換行符號和空格移除和登記編號
result_list_filtered = [item.replace('\n', '') for item in result_list_filtered]
result_list_filtered = [item.replace(' ', '') for item in result_list_filtered]
result_list_filtered = [re.sub(r'登記編號：\d{8},', '', item) for item in result_list_filtered]
print(result_list_filtered)

# 定義資料庫查詢語法
check_query = "SELECT COUNT(*) FROM company WHERE 統一編號 = %s"

for data in result_list_filtered:
    if data == "資料不存在":
        continue

    # 使用正則表達式取值
    match = re.match(r'(.*?)統一編號：(.*),登記機關：(.*),登記現況：(.*),地址：(.*),資料種類：(.*),核准設立日期：(.*),核准變更日期：(.{7})詳細資料', data)
    #match = re.match(r'(.*?)統一編號：(.*?)(,|，)?(登記機關：(.*?),)?(登記現況：(.*?),)?(地址：(.*?),)?(資料種類：(.*?),)?(核准設立日期：(.*?),)?(核准變更日期：(.{7}))?(詳細資料)?', data)
    if match:
        # 提取字串中的值
        公司名稱 = match.group(1)
        統一編號 = match.group(2)
        登記機關 = match.group(3)
        登記現況 = match.group(4)
        地址 = match.group(5)
        資料種類 = match.group(6)
        核准設立日期 = match.group(7)
        核准變更日期 = match.group(8)

         # 確認統編是否存在 避免重複新增
        cursor.execute(check_query, (統一編號,))
        record_count = cursor.fetchone()[0]
        
        if record_count > 0:
            update_query = """
                UPDATE company
                SET 公司名稱 = %s,
                    登記機關 = %s,
                    登記現況 = %s,
                    地址 = %s,
                    資料種類 = %s,
                    核准設立日期 = %s,
                    核准變更日期 = %s
                WHERE 統一編號 = %s
            """
            cursor.execute(update_query, (公司名稱 or None, 登記機關 or None, 登記現況 or None, 地址 or None, 資料種類 or None, 核准設立日期 or None, 核准變更日期 or None, 統一編號))
            print(公司名稱)
conn.commit()

# 關閉游標和連接
cursor.close()
conn.close()
driver.quit()