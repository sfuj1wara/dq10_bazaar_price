# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import time
import re
import config
import openpyxl
import datetime

# 出品情報を格納するリスト
exhibition_data = []

options = webdriver.ChromeOptions()

#ヘッドレスモード
options.add_argument('--headless')

#webdriverのパスを指定
driver = webdriver.Chrome(executable_path="chromedriver.exe", chrome_options=options)
#検索先のURL
driver.get('https://hiroba.dqx.jp/sc/search/')

time.sleep(3)

#検索窓入力
s = driver.find_elements_by_xpath('//*[@id="sqexid"]')
s[0].send_keys(config.USERID)
s = driver.find_elements_by_xpath('//*[@id="password"]')
s[0].send_keys(config.PASSWORD)

# ログインボタンクリック
driver.find_element_by_xpath('//*[@id="login-button"]').click()
driver.find_element_by_xpath('//*[@id="welcome_box"]/div[2]/a').click()
driver.find_element_by_xpath('//*[@id="contentArea"]/div/div[2]/form/table/tbody/tr[2]/td[3]/a').click()

# 検索
time.sleep(2)

while True:
    try:
        print("検索するアイテム名を入力してください：")
        # 検索するアイテム名を格納
        search_word = input()

        # 何も入力されなければ終了
        if search_word == "":
            print("プログラムを終了します")
            break
    
        # 検索フォームに入力
        s = driver.find_elements_by_xpath('//*[@id="searchword"]')
        s[0].send_keys(search_word)

        driver.find_element_by_xpath('//*[@id="searchBoxArea"]/form/p[2]/input').click()
        driver.find_element_by_xpath('//*[@id="contentArea"]/div/div[4]/table/tbody/tr/th/table/tbody/tr/td[3]/a').click()
        time.sleep(3)
        driver.find_element_by_xpath('//*[@id="btn_lock"]/a').click()
        time.sleep(5)

        # 出品データの取得
        elements = driver.find_elements_by_tag_name('tr')
    
        # リストに格納
        for elem in elements:
            exhibition_data.append(elem.text.split())
    
        # 先頭に5つ入ってくる無駄な要素を削除
        del exhibition_data[:6]

        workbook = openpyxl.Workbook()
        sheet = workbook.active
    
        # Excelのシートへ出力
        for i in range(len(exhibition_data)):
            for j in range(len(exhibition_data[i])):

                sheet.cell(row=i + 1, column=j + 1).value = exhibition_data[i][j]
    
        # リストの中身を一旦リセット
        del exhibition_data[:]

        # Excelファイルを保存
        workbook.save(search_word + "_" + datetime.datetime.now().strftime("%Y%m%d%H%M") + '.xlsx')

        # ワークブックを閉じる
        workbook.close()

    except NoSuchElementException:
        print("検索結果がありません")

# ドライバーを閉じる
driver.quit()