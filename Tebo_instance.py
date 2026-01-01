import os
from os.path import join, exists

import math
import pandas as pd


from TeboCADProcess import *

# 人工確認
New_CAD_folder = "25W12-SB_1216WYHQ1400_cad-Basic"  # 替換資料夾名稱
Old_CAD_folder = "25W12-SB_1212YHQ1340_CAD-Basic"


def Tebo_instance():

    print("please key in new cad folder:")
    new_folder = input().strip()

    print("please key in old cad folder:")
    old_folder = input().strip()


    # 確認路徑存在
    if not os.path.exists(new_folder):
        print(f"Error: new cad folder '{new_folder}' not found.")
        return
    if not os.path.exists(old_folder):
        print(f"Error: old cad folder '{old_folder}' not found.")
        return

    execute_Nails_summary(Nails_asc_output, new_folder, old_folder)
    execute_Parts_summary(Parts_asc_output, new_folder, old_folder)
    
    # execute_Nails_summary(Nails_asc_output, New_CAD_folder, Old_CAD_folder)
    # execute_Parts_summary(Parts_asc_output, New_CAD_folder, Old_CAD_folder)


    return None


if __name__ == "__main__":
    Tebo_instance()