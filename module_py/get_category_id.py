import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import re

def get_category_id_from(url_list):
    """
    指定されたURLリストの各ページからパンくずリストの最後の項目を取得
    """
    # Seleniumのオプションを設定
    options = Options()
    options.add_argument("--headless")  # ヘッドレスモード（GUIなし）
    options.add_argument("--no-sandbox")  # サンドボックスを無効化
    options.add_argument("--disable-dev-shm-usage")  # 共有メモリの問題を回避
    options.headless = False  # ヘッドレスモードを無効化（デバッグ時はFalse）

    # Chrome WebDriverを起動
    driver = webdriver.Chrome(options=options)

    try:
        results = {}  # URLごとの結果を格納

        for url in url_list:
            print(f"Processing: {url}")
            driver.get(url)

            # ページが完全に読み込まれるまで待機
            time.sleep(0.5)  # WebDriverWaitを使用するのも可

            breadcrumbs = driver.find_elements(By.CSS_SELECTOR, ".seo-breadcrumb-text")

            results[url] = None
            if breadcrumbs:
                # 最後の 'li' 要素のテキストを取得
                href = breadcrumbs[-1].get_attribute("href")
                # print(f"Category: {last_text}")
                match = re.search(r"/(\d+)/", href)
                if not match:
                    continue
                category_id = match.group(1)
                print(category_id)
                results[url] = category_id
            else:
                # print("パンくずリストが見つかりませんでした")
                results[url] = None

        return results  # 結果を返す

    finally:
        # ブラウザを閉じる
        driver.quit()


if __name__ == "__main__":
    url_list = [
        r"https://www.ebay.com/itm/254654754368?_trkparms=amclksrc%3DITM%26aid%3D777008%26algo%3DPERSONAL.TOPIC%26ao%3D1%26asc%3D20230823115209%26meid%3D73fbdba455e2473b95f3106d7db26d91%26pid%3D101800%26rk%3D1%26rkt%3D1%26itm%3D254654754368%26pmt%3D1%26noa%3D1%26pg%3D4375194%26algv%3DRecentlyViewedItemsV2SignedOut%26brand%3DHP&_trksid=p4375194.c101800.m5481&_trkparms=parentrq%3A74d020f91950a56d303d1f49ffffdeee%7Cpageci%3A1594897e-fbf5-11ef-a024-464a9449f2de%7Ciid%3A1%7Cvlpname%3Avlp_homepage",
        r"https://www.ebay.com/itm/286387809659?_skw=Nintendo+64+Mario&epid=2164341466&itmmeta=01JNTC2MEMAZFDJWCYTZB0FEJ1&hash=item42ae0b2d7b:g:NgwAAeSwC2hnye8f&itmprp=enc%3AAQAKAAAA8FkggFvd1GGDu0w3yXCmi1dMBvEKZqOUeq7xBmaE2qIYIO6Vq2jL0zTzx8HQURfRDtBe9trkRFPxHnz9nPoU5yC2GuUiPcwe2TrrZRxk5el%2BEK41Z7q5aahZl9bfAjPgY2twWAyA7XxsrW2hXUitPktoNweOd6DfSkP2oimn9946Qcy3UwX5veu9lw1%2Flc62UmpOB4JzPXZI4SOVB6r7HGwVjVoWqaFpglksec9dMN2SmhCLig%2FJ%2F1YxOj%2BJj3dJv%2FaFs0V05eRx8Y6EzvtmRwrKERofot%2FsgXW48PTXwZcxae67KIRIrtXZtQkjv%2FXaMg%3D%3D%7Ctkp%3ABk9SR77HisyuZQ",
    ]
    results = get_category_id_from(url_list)

    # 最終結果を出力
    print()
    for url, category in results.items():
        print(f"Category: {category}")
