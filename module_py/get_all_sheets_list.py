import pandas as pd

def get_all_sheets_list(search_words_excel):
    xls = pd.ExcelFile(search_words_excel)
    all_sheets = xls.sheet_names  # シート名のリスト
    # print(all_sheets)
    return all_sheets

if __name__ == "__main__":
    print(get_all_sheets_list("scraiping_list.xlsx"))