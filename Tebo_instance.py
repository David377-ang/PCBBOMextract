import os
from os.path import join, exists

import math
import pandas as pd

from Instance import get_executable_path
from Instance import create_or_replace_file

from TeboCADProcess import *



def Tebo_instance():

    executable_dir = get_executable_path()
    print(f"執行檔所在目錄: {executable_dir}")


    CAD_new = parse_Nailsasc(os.path.join(executable_dir, "25W12-SB_1216WYHQ1400_cad-Basic", Nails_asc_name))
    # print(CAD_new.head())
    # print(CAD_new.tail())

    CAD_old = parse_Nailsasc(os.path.join(executable_dir, "25W12-SB_1212YHQ1340_CAD-Basic", Nails_asc_name))
    # print(CAD_old.tail())
     
    create_or_replace_file(os.path.join(executable_dir, Nails_asc_output))

    CAD_Nailsasc_shift = find_Nailsasc_shift(CAD_new, CAD_old)
    save_Nails_shift_notebook(CAD_Nailsasc_shift, Nails_asc_output, Nails_shift_threshold, "25W12-SB_1216", "25W12-SB_1212")  # 預設存成 Diff_Nails_report.txt


    CAD_Nailsasc_del = find_Nailsasc_Del(CAD_new, CAD_old)
    save_Nails_del_notebook(CAD_Nailsasc_del, Nails_asc_output, "25W12-SB_1216", "25W12-SB_1212")


    CAD_Nailsasc_add = find_Nailsasc_Add(CAD_new, CAD_old)
    save_Nails_add_notebook(CAD_Nailsasc_add, Nails_asc_output, "25W12-SB_1216", "25W12-SB_1212")


     
    return None


if __name__ == "__main__":
    Tebo_instance()