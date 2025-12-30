import os
from os.path import join, exists
import re

import pandas as pd
from Instance import get_executable_path


def parse_Nailsasc(filepath):
    """
    解析 ASC 檔案，回傳包含 X, Y, T/B, Net Name 的 DataFrame。
    - 自動跳過表頭行
    - 只解析以 $ 開頭的資料行
    """
    records = []
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            # 跳過表頭與空行，只處理 $ 開頭的資料行
            if not line.startswith("$"):
                continue
            parts = line.split()
            # parts 範例：
            # ['$1','2.4303','4.3900','5','AA','(T)','#798','FM_PVDDIO_MEM_S3_EN','T','PIN','TB_TP103.1']
            x = float(parts[1])
            y = float(parts[2])
            tb = parts[5].strip("()")   # 去掉括號，只留 T 或 B
            net_name = parts[7]         # Net Name 在第 8 欄
            records.append({"X": x, "Y": y, "T/B": tb, "Net Name": net_name})

    return pd.DataFrame(records)





def Tebo_instance():

    executable_dir = get_executable_path()
    print(f"執行檔所在目錄: {executable_dir}")


    data = parse_Nailsasc(os.path.join(executable_dir, "25W12-SB_1212YHQ1340_CAD-Basic", "Nails.asc"))
    print(data.head())
    print(data.tail())
    return None


if __name__ == "__main__":
    Tebo_instance()