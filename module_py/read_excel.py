import pandas as pd

def get_excel_data(input_excel="scraiping_list.xlsx", 
                    sheet_name="ワードリスト",
                    usecols:list=[0, 1, 2, 3, 4],
                    col_names:list[str]=["skip_flag", "filter_counts", "number", "search", "exclude"]
                   ):
    """
    エクセルデータを取得する関数。
    Args:
        input_excel(str): エクセルファイルのパス
        sheet_name(str): エクセルシート名
        usecols(list[int]): データを取得する列。0から始まる。
        col_names(list[str]): usecols と同じ数。データを取得したときに、それぞれのデータにつける名前。
    Returns:
        list: usecols と同じ数のリストを返す。
    Raises:
        NA: NA
    Examples:
        if __name__ == "__main__": を確認
    """
    
    for val in [input_excel, sheet_name, usecols, col_names]:
        # print(val)
        pass

    df = pd.read_excel(input_excel, sheet_name=sheet_name, 
                       usecols=usecols,
                       names=col_names,
                       header=None,
                       )
    df = df.iloc[1:]

    return [df[col].tolist() for col in col_names]


if __name__ == "__main__":
    print(get_excel_data())
    print(get_excel_data(sheet_name="1", usecols=[3,4], col_names=["reduced_title", "url"]))