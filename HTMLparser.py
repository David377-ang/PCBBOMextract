import argparse
import os
from os.path import join, exists
import re
import sys
from bs4 import BeautifulSoup
from Instance import *
import csv

HTML_comp_raw_output = "HTML_comp_raw_output.txt"
comp_testability_output = "comp_testability_output.txt"


html_1_file_name = "1.htm"  # 替換成你的 HTML 檔案名稱

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



def Get_comp_raw_list(soup_raw, dir_src):

    output_file = os.path.join(dir_src, HTML_comp_raw_output)
                               
    Comp_list = []
    Comp_list = extract_all_component(
        soup_raw,
        start_idx=5,
        end_idx=10,
        output_path=output_file,
        write_func=write_list_to_file
    )

    return Comp_list

def evaluate_testability(Comp_raw_list_src, BOM_Comp_list_src):
    """
    根據 BOM 清單比對 HTML 元件資料，判斷每個元件是否可測試，
    並回傳格式為 [元件編號, 腳數, 不可植針腳數, 測試狀態]

    Comp_raw_list 的欄位定義：
        item[0] = 元件編號
        item[3] = 腳數 (Pin Count)
        item[4] = 不可植針腳數 (Unpluggable Pin Count)
        item[5] = 測試百分比 (Test Coverage)
    """
    results = []

    for bom_part in BOM_Comp_list_src:
        match = next((item for item in Comp_raw_list_src if item[0] == bom_part), None)

        if match:
            percent = match[5].strip()
            if percent == "100.0%":
                status = "testable"
            elif percent == "0%":
                status = "untestable"
            else:
                status = "limit testable"
            pin_count = match[3]
            unpluggable_pins = match[4]
        else:
            status = "not found"
            pin_count = "-"
            unpluggable_pins = "-"

        results.append([bom_part, pin_count, unpluggable_pins, status])

    return results


def main_HTMLparser():

    executable_dir = get_executable_path()
    print(f"執行檔所在目錄: {executable_dir}")
    create_or_replace_file(os.path.join(executable_dir, HTML_comp_raw_output))
    create_or_replace_file(os.path.join(executable_dir, comp_testability_output))
    # Output_list =[]

    # html1 get 元件清單, HTML_comp_raw_output.txt
    soup = read_html_by_name(os.path.join(executable_dir, r"VF", html_1_file_name))
    Comp_raw_list = []
    Comp_raw_list = Get_comp_raw_list(soup, executable_dir)

    # get BOM 元件清單
    BOM_Comp_list = []
    BOM_Comp_list = extract_location_texts_SFCS(r"BOM.20241007_B91.04G10.000V.txt")


    # 判斷 BOM 元件可測度 , comp_testability_output.txt
    Comp_testability_list = []
    Comp_testability_list = evaluate_testability(Comp_raw_list, BOM_Comp_list)

    write_list_to_file(Comp_testability_list, comp_testability_output) 


    # write_list_to_file(Output_list)    


if __name__ == "__main__":
    main_HTMLparser()







