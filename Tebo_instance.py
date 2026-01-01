import os
from os.path import join, exists

import math
import pandas as pd


from TeboCADProcess import *

# 人工確認
New_CAD_folder = "25W12-SB_1216WYHQ1400_cad-Basic"  # 替換資料夾名稱
Old_CAD_folder = "25W12-SB_1212YHQ1340_CAD-Basic"


def Tebo_instance():


    execute_Nails_summary(Nails_asc_output, New_CAD_folder, Old_CAD_folder)

    return None


if __name__ == "__main__":
    Tebo_instance()