import argparse
import os
from os.path import join, exists
import re
import sys
import pandas as pd
import subprocess
import importlib
import math




def get_executable_path():
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # PyInstaller --onefile 或 --onedir 模式
        return os.path.dirname(sys.executable)
    else:
        # 一般情況
        return os.path.dirname(os.path.abspath(__file__))

def list_files_scandir(path):
    """使用 os.scandir() 列出指定路徑下的所有檔案 (不包含子目錄)"""
    try:
        for entry in os.scandir(path):
            if entry.is_file():  # 使用 is_file() 方法檢查是否為檔案
                print(entry.path)
    except FileNotFoundError:
        print(f"找不到路徑：{path}")
    except Exception as e:
        print(f"發生錯誤：{e}")

def get_file_list(path):
    """使用 os.scandir() 獲取檔案列表 (不包含子目錄)"""
    try:
        return [entry.path for entry in os.scandir(path) if entry.is_file()]
    except FileNotFoundError:
        print(f"找不到路徑：{path}")
        return []  # 找不到路徑時返回空列表
    except Exception as e:
        print(f"發生錯誤：{e}")
        return []

def search_string_in_file(filepath, search_string):
    """搜尋檔案中特定字串"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:  # 確保能處理中文
            for line_num, line in enumerate(f, 1):
                if search_string in line:
                    print(f"在第 {line_num} 行找到：{line.strip()}")
    except FileNotFoundError:
        print(f"找不到檔案：{filepath}")
    except Exception as e:
        print(f"發生錯誤：{e}")

def find_string_in_file_with_re(filepath, target_pattern):
    """
    使用 re.compile 在文件中查找符合模式的字符串，并返回包含该字符串的行。

    Args:
        filepath: 文件路径。
        target_pattern: 要查找的目标字符串的正则表达式模式。

    Returns:
        一个列表，包含所有包含目标字符串的行（包含行号）。
        如果未找到目标字符串，则返回空列表。
        如果发生错误（如文件未找到），则返回错误消息字符串。
    """
    matching_lines = []
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            compiled_pattern = re.compile(target_pattern)  # 编译正则表达式
            for line_number, line in enumerate(file, 1):
                if compiled_pattern.search(line):  # 使用 search() 查找
                    matching_lines.append(f"Line {line_number}: {line.strip()}")
    except FileNotFoundError:
        return f"Error: File not found at {filepath}"
    except Exception as e:
        return f"Error: An unexpected error occurred: {e}"

    return matching_lines



def print_file_content(filepath):
    """印出檔案的完整內容"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:  # 確保能處理中文
            lines = f.readlines()
            for line in lines:
                print(line.strip()) # 逐行印出，並去除每行末尾的換行符
    except FileNotFoundError:
        print(f"找不到檔案：{filepath}")
    except Exception as e:
        print(f"發生錯誤：{e}")

def create_or_replace_file(file_name):
    """
    檢查指定檔案是否存在，若存在則刪除後重新建立，若不存在則直接建立。

    Args:
        file_name: 要建立或替換的檔案名稱。
    """
    if os.path.exists(file_name):
        try:
            os.remove(file_name)
            print(f"檔案 {file_name} 已存在，已刪除並重新建立。")
        except OSError as e:
            print(f"刪除檔案 {file_name} 時發生錯誤: {e}")

        with open(file_name, "w") as f:
            # 可以在這裡寫入檔案內容
            pass  
    else:
        with open(file_name, "w") as f:
            # 可以在這裡寫入檔案內容
            pass  
        print(f"檔案 {file_name} 不存在，已建立。")

def write_list_to_file(data_list, filename="parser_result.txt"):
    """
    將列表的內容逐行寫入指定的檔案。如果檔案已存在，則會覆蓋原有內容。

    Args:
        data_list: 要寫入的列表。
        filename: 要寫入的檔案名稱，預設為 "parser_result.txt"。
    """
    try:
        with open(filename, "a", encoding="utf-8") as f:  # 使用 utf-8 編碼開啟檔案
            for item in data_list:
                f.write(str(item) + "\n")  # 將每個列表元素轉換為字串，並加上換行符號
        print(f"已將列表內容寫入檔案: {filename}")
    except Exception as e:
        print(f"寫入檔案時發生錯誤: {e}")

def write_string_to_file(text, filename="parser_result.txt"):
    """
    將字串寫入指定的檔案。如果檔案已存在，則會覆蓋原有內容。

    Args:
        text: 要寫入的字串。
        filename: 要寫入的檔案名稱，預設為 "parser_result.txt"。
    """
    try:
        with open(filename, "a", encoding="utf-8") as f:  # 使用 utf-8 編碼開啟檔案
            f.write(text + "\n")  # 將字串寫入檔案
        print(f"已將字串寫入檔案: {filename}")
    except Exception as e:
        print(f"寫入檔案時發生錯誤: {e}")


def find_next_line_after_block(file_content, test_item):
    """
    找到包含指定 test_item 的 {@BLOCK} 行的下一行。

    Args:
        file_content: 檔案內容 (字串).
        test_item: 要尋找的測試項名稱 (字串，例如 "led1%cr%led").

    Returns:
        如果找到測試項和下一行，則返回下一行的字串。
        如果找不到測試項或沒有下一行，則返回 None。
    """

    # 使用正則表達式找到包含 test_item 的 {@BLOCK} 行
    #  .*? 匹配任意字符（非貪婪模式）
    block_match = re.search(r'{@BLOCK\|' + re.escape(test_item) + r'\|.*?}', file_content)

    if block_match:
        block_line = block_match.group(0)  # 取得匹配到的 {@BLOCK} 行
        #print(block_line)
        # 找到 {@BLOCK} 行在整個檔案內容中的位置
        block_start = file_content.find(block_line)
        #print(block_start)

        # 從 {@BLOCK} 行之後的部分開始尋找
        remaining_content = file_content[block_start + len(block_line):]
        #print(remaining_content)

        # 找到第一個換行符號 (\n) 的位置
        newline_pos = remaining_content.find('\n')
        #print(newline_pos)
        if newline_pos != -1:
            # 提取下一行
            next_line = remaining_content[:newline_pos].strip()
            return next_line

    return None  # 找不到測試項或沒有下一行

def find_single_result_after_BLOCK(filepath, test_item):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            for line in f:

                match = re.search(r"{@BTEST\|(\w+)\|", line)
                if match:
                    extracted_value = match.group(1)  # 取得第一個分組 SN
                    # print(extracted_value)
                else:
                    pass
                    # print("找不到符合的字串")

                if re.search(r'{@BLOCK\|' + re.escape(test_item) + r'\|', line):
                    try:
                        next_line = next(f).strip() # 獲取下一行
                        return " ".join([extracted_value, next_line])
                    except StopIteration:
                        return "{@BLOCK} 行後沒有內容"
            return " ".join([extracted_value, "找不到包含指定內容的 {@BLOCK} 行"])
    except FileNotFoundError:
        return f"找不到檔案：{filepath}"
    except Exception as e:
        return f"發生錯誤：{e}"


def get_value_from_Keyfile(filepath="Key.txt"):
    """
    從檔案中讀取 "key=" 對應的值。

    Args:
        filepath: 檔案路徑。

    Returns:
        如果找到 "key="，則傳回其對應的值（字串），否則傳回 None。
    """
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line.startswith("key="):  # 直接檢查 "key="
                    return line[4:]  # "key=" 長度為 4
    except FileNotFoundError:
        print(f"找不到檔案: {filepath}")
        return None
    except Exception as e:
        print(f"讀取檔案時發生錯誤: {e}")
        return None

def extract_location_texts_SFCS(file_name):

    # 定義正則表達式模式
    pattern = r"\|\s*([\w\.]+)\s*\|([^\|]+)\|.*?\n((?:\s*\|[\w\s]+\n?)+)"

    # 初始化空列表存儲結果
    results = []

    # 打開並讀取檔案
    with open(file_name, "r", encoding="utf-8") as file:
        content = file.read()
        # 匹配所有模式
        matches = re.finditer(pattern, content)
        for match in matches:
            item = match.group(1).strip()  # 提取 Item
            description = match.group(2).strip()  # 提取 Description
            locations = match.group(3).replace("\n", " ").replace("|", "").strip()  # 合併多行
            locations = re.sub(r"\s{2,}", " ", locations)  # 消除多餘空格

            # 將 Locations 拆分為獨立項目
            for location in locations.split():
                results.append((location))
                # results.append((item, description, location))

        # 輸出結果
        # for item, description, location in results:
        #     print(f"Item: {item}, Description: {description}, Location Texts: {location}")

    return results



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
    print(locations)


    return locations



def main():

    executable_dir = get_executable_path()
    print(f"執行檔所在目錄: {executable_dir}")

    create_or_replace_file(os.path.join(executable_dir, "parser_result.txt"))

    Output_list =[]

    # Output_list = extract_location_texts_SFCS(r"BOM.20250228_B91.10H10.001M.txt")
    # print(Output_list)

    Output_list = extract_location_texts_PLM(r"agile_20250828_052340650.xls")


    write_list_to_file(Output_list)    


if __name__ == "__main__":
    main()







