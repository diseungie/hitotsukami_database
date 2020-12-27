# プログラム0　「ライブラリの設定（下準備）」
from bs4 import BeautifulSoup
import urllib.request as req
from selenium import webdriver
import time
import pandas as pd
from datetime import datetime
import openpyxl as px
from openpyxl.styles import PatternFill
import requests
import numpy as np
from selenium.webdriver.support.ui import Select

risyu_list = []


# プログラム1-1　「GoogleChromeを起動」
browser = webdriver.Chrome(executable_path = 'C:\\Users\\Sota2020\\Desktop\\MyPandas\\chromedriver.exe')
browser.implicitly_wait(1)


# プログラム1-2　「学外用シラバス検索サイトへアクセス」
url_syllabus = "https://syllabus.cels.hit-u.ac.jp/"
browser.get(url_syllabus)
time.sleep(1)


for yobi in range(1, 6):
    for jigen in range(1, 6):
# プログラム1-4　「ドロップダウンメニュー選択（曜日）」
        yobi_element = browser.find_element_by_name('yobi')
        yobi_select_element = Select(yobi_element)
        yobi_select_element.select_by_value(str(yobi))

# プログラム1-5　「ドロップダウンメニュー選択（時限）」
        jigen_element = browser.find_element_by_name('jigen')
        jigen_select_element = Select(jigen_element)
        jigen_select_element.select_by_value(str(jigen))

# プログラム1-6　「ドロップダウンメニュー選択（検索結果表示件数）」
        _displayCount_element = browser.find_element_by_name('_displayCount')
        _displayCount_select_element = Select(_displayCount_element)
        _displayCount_select_element.select_by_value('200')

# プログラム1-7　「入力した検索条件で検索開始」
        browser_from = browser.find_element_by_name('_eventId_search')
        time.sleep(1)
        browser_from.click()

# プログラム2-1　「授業科目一覧ページのURLを取得」
        cur_url = browser.current_url


# プログラム2-2　「授業科目一覧ページ（ブラウザで表示中）の全ソースを取得」
        cur_source = browser.page_source


# プログラム2-3　「HTML解析と必要情報取得（と整理）」
        soup = BeautifulSoup(cur_source, 'lxml')
        TD = soup.find_all('td')

        for j in TD:
            risyu_list.append(j.text)
        

# プログラム2-4　「検索画面へ戻る」
        browser.back()
        
# プログラム2-5　「取得情報の整理」
risyu_list = list(map(lambda x : x.replace("\n",""),risyu_list))
data = np.array(risyu_list).reshape(-1,7)


# プログラム2-6　「不要科目削除」
# 大学院科目
data = data[data[:, 0] != '\n経営管理研究科']
data = data[data[:, 0] != '\n経済学研究科']
data = data[data[:, 0] != '\n法学研究科']
data = data[data[:, 0] != '\n社会学研究科']
data = data[data[:, 0] != '\n言語社会研究科']
data = data[data[:, 0] != '\n国際企業戦略研究科']
data = data[data[:, 0] != '\n国際・公共政策教育部']
data = data[data[:, 0] != '\n商学研究科']
data = data[data[:, 0] != '\n法科大学院']

# 全学共通教育科目ゼミ
data = data[data[:, 3] != '\n共通ゼミナール(3・4年)\n']
data = data[data[:, 3] != '\n共通ゼミナール（副）(3・4年)\n']
data = data[data[:, 3] != '\n共通ゼミナール(3年)\n']
data = data[data[:, 3] != '\n共通ゼミナール（副）(3年)\n']
data = data[data[:, 3] != '\n共通ゼミナール(4年)\n']
data = data[data[:, 3] != '\n共通ゼミナール（副）(4年)\n']

# 商学部ゼミ
data = data[data[:, 3] != '\n主ゼミナール(3年)\n']
data = data[data[:, 3] != '\n副ゼミナール(3年)\n']
data = data[data[:, 3] != '\n主ゼミナール(4年)\n']
data = data[data[:, 3] != '\n副ゼミナール(4年)\n']

# 経済学部ゼミ
data = data[data[:, 3] != '\nゼミナール（3年）\n']
data = data[data[:, 3] != '\nゼミナール（4年）\n']
data = data[data[:, 3] != '\n副ゼミナール(3年)\n']
data = data[data[:, 3] != '\n副ゼミナール(4年)\n']

# 法学部ゼミ
data = data[data[:, 3] != '\n主ゼミナール(3・4年)\n']
data = data[data[:, 3] != '\n副ゼミナール(3・4年)\n']
data = data[data[:, 3] != '\n主ゼミナール(3年)\n']
data = data[data[:, 3] != '\n副ゼミナール(3年)\n']
data = data[data[:, 3] != '\n主ゼミナール(4年)\n']
data = data[data[:, 3] != '\n副ゼミナール(4年)\n']

# 社会学部ゼミ
data = data[data[:, 3] != '\n主ゼミナール(3・4年)\n']
data = data[data[:, 3] != '\n副ゼミナール(3・4年)\n']
data = data[data[:, 3] != '\n主ゼミナール(3年)\n']
data = data[data[:, 3] != '\n副ゼミナール(3年)\n']
data = data[data[:, 3] != '\n主ゼミナール(4年)\n']
data = data[data[:, 3] != '\n副ゼミナール(4年)\n']

# PACE
data = data[data[:, 3] != '\nPACEⅠ(再履修1)\n']
data = data[data[:, 3] != '\nPACEⅡ(再履修1)\n']

for a in range(1, 63):
    data = data[data[:, 3] != '\nPACEⅠ(' + str(a) + ')\n']

data = data[data[:, 3] != '\nPACEⅠ(再履修2)\n']
data = data[data[:, 3] != '\nPACEⅡ(再履修2)\n']

for b in range(1, 63):
    data = data[data[:, 3] != '\nPACEⅡ(' + str(b) + ')\n']



# プログラム3-1　「エクセルを取得」
wb = px.Workbook()
ws = wb.active
 
# プログラム3-2　「エクセルのヘッダーの背景色を設定」
fill = PatternFill(patternType='solid', fgColor='e0e0e0', bgColor='e0e0e0')
 
# プログラム3-3　「エクセル1行目のヘッダーを出力」
headers = ['所属', '学期', '曜日・時限', '科目名','担当','教授言語','抽選対象']
for i, header in enumerate(headers):
    ws.cell(row=1, column=1+i, value=headers[i])
    ws.cell(row=1, column=1+i).fill = fill

# プログラム3-4　「エクセル2行目以降のデータを出力」
for y, row in enumerate(data):
    ws.cell(row= y+2, column= 1, value= y+1)
    for x, cell in enumerate(row):
        ws.cell(row= y+2, column= x+1, value=data[y][x])
        
# プログラム3-5　「セル幅の調整」
for col in ws.columns:
    max_length = 0
    column = col[0].column

    for cell in col:
        if len(str(cell.value)) > max_length:
            max_length = len(str(cell.value))

    adjusted_width = (max_length + 2) * 1.5
    ws.column_dimensions[col[0].column_letter].width = adjusted_width

    
# プログラム3-6　「日付を取得（エクセルファイルに日付をつけてわかりやすくした！！）」
now = datetime.now()
hiduke = now.strftime('%Y-%m-%d')


# プログラム3-7　「エクセルファイルの保存」
filename = hiduke + '_' + 'シラバスクローリング' + '.xlsx'
wb.save(filename)

print("情報を取得し、エクセルファイルに保存しました。")
