import openpyxl

def column_letter_to_number(column_letter:str):
    """列アルファベットを列番号に変換（例: 'A' -> 1, 'B' -> 2）"""
    return openpyxl.utils.column_index_from_string(column_letter)