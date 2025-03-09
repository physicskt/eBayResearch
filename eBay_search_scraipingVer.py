import csv
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup  # BeautifulSoupをインポート
import re
import time
import pandas as pd
import os
import sys

# 自分のモジュール読み込み
sys.path.append("module_py")
from module_py import read_excel
from module_py import write_excel
from module_py import delete_sheets
from module_py import eBay_check_identity
from module_py import get_category_id
from module_py import get_all_sheets_list
from module_py import log
from module_py import column_letter_to_number

def search_words_in_eBay(
        search_words_list, 
        exclude_words_list,
        search_number, 
        skip_flags, 
        ):
    
    log.log_less_message(search_words_list)
    log.log_less_message(exclude_words_list)

    # Chromeのオプションを設定
    options = Options()
    options.add_argument("--headless")  # GUIなし（バックグラウンド実行）
    options.add_argument("--no-sandbox")  # サンドボックスを無効化
    options.add_argument("--disable-dev-shm-usage")  # 共有メモリの問題を回避
    options.headless = False  # ヘッドレスモードを無効化（必要に応じてTrueに設定）

    # サービスを設定
    service = Service(ChromeDriverManager().install())

    # Chromeドライバを起動
    driver = webdriver.Chrome(service=service, options=options)

    output_dict = {}
    used_sheets = ["ワードリスト"]
    for i, search_words in enumerate(search_words_list):
        used_sheets.append(str( search_number[i] ))

        if not "いいえ" in skip_flags:
            log.log_less_message("全て Skip する、に設定されています。")
            log.log_less_message("処理を終了します。")
            continue

        log.log_less_message("検索ワード: ")
        log.log_less_message(search_words)
        if pd.isna(search_words):
            log.log_less_message("#####################################")
            log.log_less_message("空白です。処理をSkipします。")
            log.log_less_message("")
            continue

        search_words = re.split(r'\s+', search_words)
        search_string = " ".join(search_words)

        exclude_words_arr = False
        if pd.isna(exclude_words_list[i]) is False:
            exclude_words_arr = re.split(r'\s+', exclude_words_list[i])
            log.log_less_message("分割・除去後検索ワード:")
            log.log_less_message(exclude_words_arr)

        # eBayを開く
        search_url = f"https://www.ebay.com/sch/i.html?_nkw={search_string}"
        log.log_less_message("")
        log.log_less_message("##################################")
        log.log_less_message(f"シート名: {search_number[i]}")
        log.log_less_message(search_url)
        log.log_less_message(f"スキップするか: {skip_flags[i]}")
        log.log_less_message("##################################")
        if (skip_flags[i] == "はい"):
            continue

        log.log_less_message(f"検索開始")
        driver.get(search_url)
        time.sleep(0.2)

        # 商品タイトルと価格を取得
        items = driver.find_elements(By.CLASS_NAME, 's-item')

        new_header = ["Title", "Price", "Sold", "Reduced title", "URL"]
        output_arr = [new_header]
        for item in items:
            data_record_flag = True
            title, price_text, sold_text, title_reduced, title_url = "", "", "", "", ""
            try:
                # 商品名取得
                title = item.find_element(By.CLASS_NAME, 's-item__title').text
                # title = title.replace('"', '')

                # 除去ワードリストを削除した、商品名を生成
                title_reduced = title
                if exclude_words_arr is not False:
                    for word in exclude_words_arr:
                        title_reduced = re.sub(rf"{re.escape(word)}", "", title_reduced, flags=re.IGNORECASE)
                    
                # 価格取得
                price_element = item.find_element(By.CLASS_NAME, 's-item__price').get_attribute('outerHTML')
                soup = BeautifulSoup(price_element, 'html.parser')
                price_text = soup.get_text().replace('"', '')  # コメントを除いた価格部分を抽出
                # [^.....] で....部以外を、削除する。 
                price_text = re.sub(r"[^\d円~～.]", "", price_text)

                # 販売数量
                sold_element = item.find_element(By.CLASS_NAME, 's-item__dynamic.s-item__quantitySold').get_attribute('innerHTML')
                soup = BeautifulSoup(sold_element, 'html.parser')
                sold_text = soup.get_text().replace('"', '')
                sold_text = re.sub(r"[^\d]", "", sold_text)
                sold_text = int(sold_text)

                # URL取得
                title_url = item.find_element(By.CLASS_NAME, 's-item__link').get_attribute('href')

                log.log_less_message(f"Title: {title}, Price: {price_text}, Sold: {sold_text}")
                
            except Exception as e:
                if title == "":
                    continue
                # if price_text == "":
                #     continue
                # if sold_text == "":
                #     continue

            # 販売数0ならリストに追加しない
            if sold_text == "":
                data_record_flag = False

            if data_record_flag is True:
                output_arr.append([title, price_text, sold_text, title_reduced, title_url]) 

        output_dict[i] = output_arr

    driver.quit()

    return output_dict


def scraiping():
    log.log_message("scraping 処理開始")

    os.chdir( os.path.dirname(os.path.abspath(__file__)) )
    search_words_excel = "scraiping_list.xlsx"
    output_dir = "output"
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)  
    output_file = "output_scraiping.csv"

    # excelからデータを取得
    skip_flags, filter_counts, search_number, search_words_list, exclude_words_list \
        = read_excel.get_excel_data(
            input_excel=search_words_excel,
            sheet_name="ワードリスト",
            usecols=[0, 1, 2, 3, 4],
            col_names=["skip_flag", "filter_counts", "number", "search", "exclude"]
        )

    # eBayでデータを取得
    output_dict = search_words_in_eBay(
        search_words_list, 
        exclude_words_list,
        search_number, 
        skip_flags, 
        )

    # データ書き込み
    log.log_less_message("")
    log.log_message("eBayデータをエクセルへ書き込み")
    for i, output_arr in output_dict.items():
        log.log_less_message("eBayデータをエクセルへ書き込み: シート名" + str(i+1))
        # Excelへデータ書き込み
        write_excel.write_to_data(file_name=search_words_excel,
                                  sheet_name=str( search_number[i] ),
                                  data=output_arr)

        # CSVに書き出し
        output_path = os.path.abspath( os.path.join(output_dir, str( search_number[i] ) + "_" + output_file) )
        with open(output_path, mode="w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file, quoting=csv.QUOTE_MINIMAL)
            writer.writerows(output_arr)  # データを書き込み

    # 使用されたシートを以外を削除
    log.log_message("使用されたシートを以外を削除")
    used_sheets = ["ワードリスト"]
    write_excel.close_excel_file(search_words_excel)
    sheet_num = read_excel.get_excel_data(search_words_excel, "ワードリスト", [2], ["sheet_number"])
    used_sheets.extend(sheet_num[0])
    used_sheets = [str(item) for item in used_sheets]
    log.log_less_message("次のシート以外を削除")
    log.log_less_message(used_sheets)
    # used_sheets 以外を消す
    write_excel.close_excel_file(search_words_excel)
    delete_sheets.delete_other_sheets(search_words_excel, used_sheets)

    # 集計
    log.log_message("集計中")
    # パターン1,3の集計
    used_sheets.remove("ワードリスト")
    for sheet_name in used_sheets:
        try:
            int(sheet_name)
        except Exception as e:
            # log.log_less_message(e)
            continue

        filter_count = filter_counts[int(sheet_name)-1]
        eBay_check_identity.check_and_calc_sold(search_words_excel, sheet_name, filter_count)
        log.log_less_message(f"シート[{sheet_name}] の製品名は {filter_count} 個の単語でフィルタされました。")

    log.log_message("scraping 処理終了")
    # エクセルを開く
    write_excel.open_visible_excel(search_words_excel)


def write_category_id_to_excel():
    search_words_excel = "scraiping_list.xlsx"
    xls = pd.ExcelFile(search_words_excel)
    all_sheets = xls.sheet_names  # シート名のリスト

    # "ワードリスト" 以外のシートを取得
    sheets_list = [sheet for sheet in all_sheets if sheet != "ワードリスト"]
    log.log_less_message(sheets_list)

    # 各シートを処理 for
    for sheet in sheets_list:
        log.log_message("最安値検索開始")
        # 1,2,3,4..... 以外のシートはスキップ
        try:
            int(sheet)
        except:
            log.log_less_message(f"sheet {sheet} は処理をスキップします。")
            continue

        reduced_titles, urls, flags = read_excel.get_excel_data(
            search_words_excel, 
            sheet, 
            usecols=[3,4,7], 
            col_names=["reduced_title", "url", "flag"]
            )

        category_ids = {}
        output_ids = ["Category_ID"]
        for i, title in enumerate(reduced_titles):
            category_ids[i] = None
            if bool(flags[i]) == True:
                # log.log_less_message(urls[i])
                category_ids[i] = get_category_id.get_category_id_from([urls[i]])
                log.log_less_message(f"Sheet名:{sheet}, 行{str(i+1)}")
                output_ids.append( list(category_ids[i].values())[0] )
            else:
                output_ids.append("Skip")
    
        # カテゴリIDを取得し書き込み
        log.log_less_message(output_ids)
        write_excel.append_data_to_excel(
            search_words_excel,
            sheet,
            output_ids,
            1,
            "J"
            )
    
    log.log_message("最安値検索終了")


def search_cheapst_price():
    search_words_excel = "scraiping_list.xlsx"
    sheet_names = read_excel.get_excel_data(
        input_excel=search_words_excel,
        sheet_name= "ワードリスト",
        usecols = [2],
        col_names = ["number"]
    )
    sheet_names = [str(name) for name in sheet_names[0] if not pd.isna(name) ]
    print(sheet_names)

    usecols = [
        column_letter_to_number.column_letter_to_number("D")-1,
        column_letter_to_number.column_letter_to_number("H")-1,
        ]
    
    # シート毎に処理
    for sheet in sheet_names:
        reduced_titles, flags1 = read_excel.get_excel_data(
            input_excel=search_words_excel,
            sheet_name= sheet,
            usecols = usecols,
            col_names = ["Reduced_title", "flags1"]
        )
        # print(reduced_titles)
        print(flags1)

        # ちょっと上手く動かなかった、、、、
        exclude_list = ["." for titles in reduced_titles]
        search_number = [1 for titles in reduced_titles]
        search_results = search_words_in_eBay(reduced_titles, exclude_list, search_number, flags1)
        print(search_results)

if __name__ == "__main__":
    # scraiping()
    # write_category_id_to_excel()
    search_cheapst_price()
    write_excel.open_visible_excel("scraiping_list.xlsx")
