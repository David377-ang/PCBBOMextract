import argparse
import os
from os.path import join, exists
import re
import sys
from bs4 import BeautifulSoup
from Instance import *


def read_html_by_name(file_name):
    """
    使用 BeautifulSoup 讀取特定名稱的 HTML 文件，並返回解析後的 BeautifulSoup 物件。

    :param file_name: HTML 檔案的名稱（包含路徑）
    :return: BeautifulSoup 物件
    """
    try:
        # 開啟指定的 HTML 檔案
        with open(file_name, 'r', encoding='Big5') as file:
            content = file.read()  # 讀取文件內容

        # 使用 BeautifulSoup 解析 HTML
        soup = BeautifulSoup(content, 'lxml')  # 或 'html.parser' 根據需求選擇解析器

        return soup  # 返回 BeautifulSoup 物件供進一步操作

    except FileNotFoundError:
        print(f"檔案 '{file_name}' 不存在！")
    except Exception as e:
        print(f"讀取檔案時發生錯誤: {e}")

def main_HTMLparser():

    executable_dir = get_executable_path()
    print(f"執行檔所在目錄: {executable_dir}")
    create_or_replace_file(os.path.join(executable_dir, "parser_result.txt"))
    # Output_list =[]

    html_file_name = "1.htm"  # 替換成你的 HTML 檔案名稱
    soup = read_html_by_name(os.path.join(executable_dir, r"VF", html_file_name))

    if soup:
        print(soup.title.string)  # 例如，印出 <title> 的內容


    # write_list_to_file(Output_list)    


if __name__ == "__main__":
    main_HTMLparser()







