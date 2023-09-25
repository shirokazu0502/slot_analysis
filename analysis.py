import time
import datetime
from bs4 import BeautifulSoup
from selenium import webdriver #Selenium Webdriverをインポートする
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

# Serviceオブジェクトを作成
service = Service(ChromeDriverManager().install())

# Serviceオブジェクトを使用してWebDriverを作成
options = Options()
driver = webdriver.Chrome(service=service, options=options)
#ホールの情報
hall_pref="大阪府"
hall_name="123n松原店"

#昨日の年、月、日を取得
target_date = datetime.datetime.now() - datetime.timedelta(days=1)
#for i in range(365):
#nowから年前まで繰り返す
target_year = target_date.year
#2桁で月をtarget_dateの月を取得
target_month = target_date.month
#target_monthが1桁の場合、頭に0をつける
if target_month < 10:
    target_month = "0" + str(target_month)
target_day = target_date.day
#target_dayが1桁の場合、頭に0をつける
if target_day < 10:
    target_day = "0" + str(target_day)

hp_url=f"https://ana-slo.com/{target_year}-{target_month}-{target_day}-{hall_name}-data/"
driver.get(hp_url)
time.sleep(5) #5秒待つ。冒頭のimport timeで利用可能なtimeメソッド
# スクロールするためのJavaScriptコードを実行
driver.execute_script("window.scrollTo(0, 500);")
#クリックする
driver.find_element_by_id("all_data_btn").click()
time.sleep(5) #5秒待つ。
# ページのソースコード（JavaScript実行後のコードも含む）を取得
page_source = driver.page_source
# BeautifulSoupを使って情報を抽出
print(page_source)
soup = BeautifulSoup(page_source, 'html.parser')
#全体情報テーブルを抽出
all_data_table = soup.find_all(class_="total_get_medals_table")
print(all_data_table)
#tbody要素を取得
all_data_table_tbody = all_data_table.find("tbody")
print(all_data_table_tbody)
#2番目のtr要素を取得    
all_data_table_second_tr = all_data_table_tbody.find_all("tr")[1]
#3番目のtd要素（平均差枚）を取得
ave_diff_medal = all_data_table_second_tr.find_all("td")[2]
print(ave_diff_medal.text)
#dataTabl


driver.quit() #Chromeブラウザを閉じる