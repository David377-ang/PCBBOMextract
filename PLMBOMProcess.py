import subprocess
import importlib
import math
import pandas as pd
import re

def read_excel_auto_safe(file_path, max_display_rows=200, preview_rows=10, **kwargs):
    """
    自動安裝缺少套件 + 自動讀取 Excel + 智慧防卡顯示
    :param file_path: Excel 檔案路徑
    :param max_display_rows: 超過這個行數就啟動防卡模式
    :param preview_rows: 防卡模式下，前後各顯示幾行
    :param kwargs: 傳給 pd.read_excel 的其他參數
    """
    # 1️⃣ 自動安裝必要套件
    def ensure_package(pkg_name):
        try:
            importlib.import_module(pkg_name)
        except ImportError:
            print(f"📦 偵測到缺少套件 {pkg_name}，正在安裝...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg_name])
    
    # 判斷副檔名選擇引擎
    if file_path.lower().endswith('.xls'):
        engine = 'xlrd'
        ensure_package('xlrd')
    else:
        engine = 'openpyxl'
        ensure_package('openpyxl')
    
    # 確保 pandas 存在
    ensure_package('pandas')

    # 2️⃣ 讀取 Excel
    df = pd.read_excel(file_path, engine=engine, **kwargs)
    
    # 3️⃣ 顯示統計資訊
    total_rows, total_cols = df.shape
    print(f"📊 檔案讀取完成：{file_path}")
    print(f"➡ 資料筆數（rows）：{total_rows}")
    print(f"➡ 欄位數（columns）：{total_cols}")
    print(f"➡ 欄位名稱：{list(df.columns)}")
    print("-" * 50)
    
    # 4️⃣ 智慧防卡顯示
    if total_rows <= max_display_rows:
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        print(df)
    else:
        print(f"⚠ 資料超過 {max_display_rows} 行，啟動防卡模式")
        print(f"🔍 顯示前 {preview_rows} 行：")
        print(df.head(preview_rows))
        print("...")
        print(f"🔍 顯示最後 {preview_rows} 行：")
        print(df.tail(preview_rows))
    
    return df

# 使用範例
# df = read_excel_auto_safe("你的檔案.xls", max_display_rows=200, preview_rows=10)

def extract_selected_columns(df, columns=None, output="df"):
    """
    從 DataFrame 中抽取指定欄位內容
    
    參數：
        df (pd.DataFrame): 已讀取的 DataFrame
        columns (list): 欲抽取的欄位名稱清單，預設為 Part Number, Part Classification, BOM.Location
        output (str): 輸出格式
                      "df"  -> 回傳 DataFrame
                      "list" -> 回傳 dict，每個欄位對應一個 list
                      "dict" -> 同 list（只是名稱不同，方便語意）
    
    回傳：
        pd.DataFrame 或 dict
    """
    if columns is None:
        columns = ["Part Number", "Part Classification", "BOM.Location"]
    
    # 檢查欄位是否存在
    missing_cols = [col for col in columns if col not in df.columns]
    if missing_cols:
        raise ValueError(f"缺少欄位: {missing_cols}")
    
    selected_df = df[columns]
    
    if output.lower() in ["list", "dict"]:
        return {col: selected_df[col].tolist() for col in columns}
    else:
        return selected_df





def split_locations(location_list, unique=True, natural_sort=True):
    """
    將 BOM.Location 欄位的多位置字串拆成單一位置 list，
    過濾掉 NaN，並可選擇去重與自然排序。
    
    :param location_list: 例如 result_dict["BOM.Location"]
    :param unique: 是否去除重複位置（預設 True）
    :param natural_sort: 是否使用自然排序（D2 在 D10 前）
    :return: list，例如 ["D1", "D2", "D35", ...]
    """
    result = []
    for item in location_list:
        # 跳過 None、空字串、pandas NaN、字串 "NaN"
        if item is None or (isinstance(item, float) and math.isnan(item)) or str(item).strip().lower() == "nan":
            continue

        # 拆分並去除空白
        parts = [p.strip() for p in str(item).split(",") if p.strip()]
        result.extend(parts)

    if unique:
        # 保留順序去重
        seen = set()
        result = [x for x in result if not (x in seen or seen.add(x))]

    if natural_sort:
        # 自然排序：數字部分按數值排序
        def natural_key(s):
            return [int(text) if text.isdigit() else text.lower()
                    for text in re.split(r'(\d+)', s)]
        result.sort(key=natural_key)

    return result

def extract_location_texts_PLM(file_name):

    df = read_excel_auto_safe(file_name, max_display_rows=5, preview_rows=5)

    # 取得 DataFrame 格式
    # result_df = extract_selected_columns(df)
    # print(result_df)

    # 取得 dict 格式
    result_dict = extract_selected_columns(df, output="dict")
    # print(result_dict["Part Number"])  # 只看 Part Number 欄
    # print(result_dict["BOM.Location"])  # 只看 BOM.Location 欄

    locations = split_locations(result_dict["BOM.Location"])
    # print(locations)


    return locations

