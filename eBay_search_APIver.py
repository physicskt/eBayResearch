import requests
import csv

keyword = "Nike" #ここに検索したいキーワードを入れる
app_name = "Takahash-sand-PRD-368e0ca88-ca192a9c"  # 自分のAppIDに置き換えてください
entries_per_page = 10  # 1ページあたりの最大結果数
category_id = "20649"  # カテゴリーID（例として20649 はキッチン用品カテゴリーのID）
min_price = "1000"  # 最小価格（ドル単位）
max_price = "1"  # 最大価格（ドル単位）
csv_name = "ebay_output.csv"

URL = f"https://svcs.ebay.com/services/search/FindingService/v1"
URL += "?OPERATION-NAME=findItemsByKeywords"
URL += f"&SERVICE-VERSION=1.0.0&SECURITY-APPNAME={app_name}"
URL += "&RESPONSE-DATA-FORMAT=JSON&REST-PAYLOAD"
URL += f"&keywords={keyword}&paginationInput.entriesPerPage={entries_per_page}"
URL += f"&categoryId={category_id}"  # カテゴリーIDをURLに追加します
URL += f"&itemFilter(0).name=MinPrice&itemFilter(0).value={min_price}&itemFilter(0).paramName=Currency&itemFilter(0).paramValue=USD"
URL += f"&itemFilter(1).name=MaxPrice&itemFilter(1).value={max_price}&itemFilter(1).paramName=Currency&itemFilter(1).paramValue=USD"

def get_page(page_number):
    request_url = f"{URL}&paginationInput.pageNumber={page_number}"
    print(request_url)
    request = requests.get(request_url)
    products_data = request.json()
    print(products_data)

    with open(csv_name, "a", encoding="utf-8", newline="") as csvfile:
        writer = csv.writer(csvfile)

        for item in products_data.get("findItemsByKeywordsResponse", [{}])[0].get("searchResult", [{}])[0].get("item", []):
            item_id = item.get("itemId", [""])[0]
            title = item.get("title", [""])[0]
            currency = item.get("sellingStatus", [{}])[0].get("currentPrice", [{}])[0].get("@currencyId", "")
            price = item.get("sellingStatus", [{}])[0].get("currentPrice", [{}])[0].get("__value__", "")
            item_url = item.get("viewItemURL", [""])[0]

            product_id = item.get("productId", [{}])[0].get("__value__", "")
            gallery_url = item.get("galleryURL", [""])[0]
            location = item.get("location", [""])[0]
            condition = item.get("condition", [{}])[0].get("conditionDisplayName", [""])[0]
            description = item.get("description", [""])[0]

            row = [item_id, title, currency, price, item_url, product_id, gallery_url, location, condition, description]
            writer.writerow(row)


if __name__ == "__main__":
    # 取得するページ数を指定します
    page_count = 1  # 例：10ページを取得します（1ページあたり100件のアイテム）
    
    for page_number in range(1, page_count + 1):
        get_page(page_number)

