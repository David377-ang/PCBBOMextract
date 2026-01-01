import math
import pandas as pd

Nails_asc_name = "Nails.asc"
Nails_asc_output = "Diff_Nails_report.txt"

def parse_Nailsasc(filepath, return_df=False):
    """
    解析 ASC 檔案，回傳包含 X, Y, T/B, Net Name 的資料。
    
    參數:
        filepath: str
            ASC 檔案路徑
        return_df: bool, 預設 False
            - False → 回傳 list of dict
            - True  → 回傳 pandas.DataFrame
    
    功能:
    - 自動跳過表頭行
    - 只解析以 $ 開頭的資料行
    """
    records = []
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line.startswith("$"):
                continue
            parts = line.split()
            if len(parts) < 8:
                continue
            try:
                x = float(parts[1])
                y = float(parts[2])
                tb = parts[5].strip("()")   # 去掉括號，只留 T 或 B
                net_name = parts[7]         # Net Name 在第 8 欄
                records.append({
                    "X": x,
                    "Y": y,
                    "T/B": tb,
                    "Net Name": net_name
                })
            except (ValueError, IndexError):
                continue

    # 根據選項回傳
    if return_df:
        return pd.DataFrame(records)
    else:
        return records



def find_Nailsasc_shift(CAD_new, CAD_old):
    """
    比較 CAD_new 和 CAD_old，找出 Net Name 相同但位置不同的項目
    參數:
        CAD_new, CAD_old: list of dict
            每個 dict 包含 {"X":..., "Y":..., "T/B":..., "Net Name":...}
    回傳:
        shift_list: list of dict
            包含來自 CAD_new 和 CAD_old 的 shift 類別資料
    """
    shift_list = []

    # 建立以 Net Name 為 key 的查詢表
    dict_new = {item["Net Name"]: item for item in CAD_new}
    dict_old = {item["Net Name"]: item for item in CAD_old}

    # 找出 Net Name 相同的項目
    common_names = set(dict_new.keys()) & set(dict_old.keys())

    for name in common_names:
        item_new = dict_new[name]
        item_old = dict_old[name]
        # 比較位置 (X, Y, T/B)
        if (item_new["X"], item_new["Y"], item_new["T/B"]) != (item_old["X"], item_old["Y"], item_old["T/B"]):
            # 如果位置不同 → shift 類別
            shift_list.append(item_new.copy())
            shift_list.append(item_old.copy())

    return shift_list




def save_Nails_shift_notebook(shift_list, filepath=Nails_asc_output, threshold_mil=3):
    """
    將 find_Nailsasc_shift 的結果存成筆記本文字檔
    格式：
    [Part 1] Shift Nails
    TOP Side = N
    Bottom Side = M

    Following xy location of test point be shifted between two version
    CAD_new     X_new    Y_new   (T/B)   Net Name
    CAD_old     X_old    Y_old   (T/B)   Net Name
    Distance = d inch (d_mil mil)   (*** if > threshold_mil)
    """
    lines = []
    lines.append("[Part 1] Shift Nails")

    # 只統計 CAD_new 的面數
    top_count = sum(1 for i in range(0, len(shift_list), 2) if shift_list[i]["T/B"].upper() == "T")
    bottom_count = sum(1 for i in range(0, len(shift_list), 2) if shift_list[i]["T/B"].upper() == "B")

    lines.append(f"TOP Side  = {top_count}")
    lines.append(f"Bottom Side  = {bottom_count}")
    lines.append("")  # 空行分隔

    lines.append("Following xy location of test point be shifted between two version")

    # 將 mil 轉換成 inch
    threshold_inch = threshold_mil / 1000.0

    for i in range(0, len(shift_list), 2):
        new_item = shift_list[i]
        old_item = shift_list[i+1]

        # 計算距離 (inch)
        dx = new_item['X'] - old_item['X']
        dy = new_item['Y'] - old_item['Y']
        distance_inch = math.sqrt(dx*dx + dy*dy)
        distance_mil = distance_inch * 1000  # 1 inch = 1000 mil

        # 判斷是否超過閾值
        mark = " ***" if distance_inch > threshold_inch else ""

        # 寫入文字
        lines.append(
            f"CAD_new     {new_item['X']:.4f}    {new_item['Y']:.4f}   ({new_item['T/B']})   {new_item['Net Name']}"
        )
        lines.append(
            f"CAD_old     {old_item['X']:.4f}    {old_item['Y']:.4f}   ({old_item['T/B']})   {old_item['Net Name']}"
        )
        lines.append(f"Distance = {distance_inch:.4f} inch ({distance_mil:.1f} mil){mark}")
        lines.append("")  # 空行分隔

    with open(filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"Shift 結果已存到 {filepath} (閾值 = {threshold_mil} mil)")