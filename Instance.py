import argparse
import os
from os.path import join, exists
import re
import sys

from PLMBOMProcess import extract_location_texts_PLM
from tkinter import Tk, filedialog

BOM_SFCS_output = "BOM_SFCS_output.txt"
testcoverage_output = "testcoverage_output.txt"

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


def extract_location_texts_testcoverage(file_name):

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
                # results.append((location))
                results.append((location, item, description))

        # 輸出結果
        # for item, description, location in results:
        #     print(f"Item: {item}, Description: {description}, Location Texts: {location}")

    return results


def main():

    executable_dir = get_executable_path()
    print(f"執行檔所在目錄: {executable_dir}")

    create_or_replace_file(os.path.join(executable_dir, BOM_SFCS_output))
    create_or_replace_file(os.path.join(executable_dir, testcoverage_output))

    Output_list =[]

    Tk().withdraw()  # 隱藏主視窗
    filename = filedialog.askopenfilename(title="選擇檔案")
    # print(f"你選擇的檔案是：{filename}")

    Output_list = extract_location_texts_SFCS(filename)
    # Output_list = extract_location_texts_SFCS(r"BOM.20250228_B91.10H10.001M.txt")
    # print(Output_list)

    write_list_to_file(Output_list, BOM_SFCS_output)    


    Output_list = extract_location_texts_testcoverage(filename)
    write_list_to_file(Output_list, testcoverage_output)   


if __name__ == "__main__":
    main()







