import argparse
import os
from os.path import join, exists
import re
import sys
from bs4 import BeautifulSoup
from Instance import *
import csv

HTML_output = "HTML_output.txt"


def read_html_by_name(file_name):
    """
    使用 BeautifulSoup 讀取特定名稱的 HTML 文件，並返回解析後的 BeautifulSoup 物件。

    :param file_name: HTML 檔案的名稱（包含路徑）
    :return: BeautifulSoup 物件
    """
    try:
        try:
            with open(file_name, 'r', encoding='Big5') as file:
                content = file.read()
        except UnicodeDecodeError:
            with open(file_name, 'r', encoding='utf-8') as file:
                content = file.read()

        # 使用內建 html.parser，避免額外安裝 lxml
        soup = BeautifulSoup(content, 'html.parser')
        return soup


    except FileNotFoundError:
        print(f"檔案 '{file_name}' 不存在！")
    except Exception as e:
        print(f"讀取檔案時發生錯誤: {e}")


def write_list_to_csv(data, file_path):
    """
    將二維 list 寫入 CSV 檔案，UTF-8 編碼。
    """
    try:
        with open(file_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(data)
        print(f"資料已輸出到 {file_path}")
    except Exception as e:
        print(f"寫入檔案時發生錯誤: {e}")


def extract_all_component(soup, start_idx=5, end_idx=10, output_path=None, write_func=None):
    """
    從 soup 中擷取指定範圍的 <table width="100%"> 資料，跳過每個表格的第一行。

    參數：
        soup        : BeautifulSoup 物件
        start_idx   : 起始表格索引（預設 5）
        end_idx     : 結束表格索引（預設 10，包含該索引）
        output_path : 輸出檔案路徑（可選）
        write_func  : 寫檔函式，例如 write_list_to_file 或 write_list_to_csv（可選）

    回傳：
        data (list of list) : 擷取到的資料
    """
    if not soup:
        print("⚠ soup 物件為空，無法處理")
        return []

    # 找到所有 <table width="100%">
    tables = soup.find_all('table', attrs={'width': '100%'})

    # 防呆：確保有足夠的表格
    if len(tables) <= end_idx:
        print(f"⚠ 找到的表格數量不足（只有 {len(tables)} 個），無法取得 table[{start_idx}] 到 table[{end_idx}]")
        return []

    data = []

    for idx in range(start_idx, end_idx + 1):
        specific_table = tables[idx]
        rows = specific_table.find_all('tr')

        for row_idx, row in enumerate(rows):
            # 第一行直接跳過
            if row_idx == 0:
                continue

            cells = row.find_all('td')
            row_data = [
                cell.get_text(separator=' ', strip=True).replace('\xa0', ' ')
                for cell in cells
            ]
            data.append(row_data)

    # 檢查輸出
    for item in data:
        print(item)

    # 寫入檔案（如果有提供）
    if output_path and write_func:
        write_func(data, output_path)

    return data

def main_HTMLparser():

    executable_dir = get_executable_path()
    print(f"執行檔所在目錄: {executable_dir}")
    create_or_replace_file(os.path.join(executable_dir, HTML_output))
    # Output_list =[]

    html_file_name = "1.htm"  # 替換成你的 HTML 檔案名稱
    soup = read_html_by_name(os.path.join(executable_dir, r"VF", html_file_name))

    output_file = os.path.join(executable_dir, HTML_output)
    extract_all_component(
        soup,
        start_idx=5,
        end_idx=10,
        output_path=output_file,
        write_func=write_list_to_file
    )


    # write_list_to_file(Output_list)    


if __name__ == "__main__":
    main_HTMLparser()







