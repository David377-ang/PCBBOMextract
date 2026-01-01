import math
import pandas as pd

Nails_shift_threshold = 3  # mil

Nails_asc_name = "Nails.asc"
Nails_asc_output = "Diff_Nails_report.txt"

def separator(char="-", length=100):
    return char * length


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


def save_Nails_shift_notebook(
    shift_list,
    filepath=Nails_asc_output,
    threshold_mil=3,
    label_new="CAD_new",
    label_old="CAD_old"
):
    """
    將 find_Nailsasc_shift 的結果存成筆記本文字檔
    格式：
    [Part 1] Shift Nails
    TOP Side = N
    Bottom Side = M

    Following xy location of test point be shifted between two version
    <label_new>   X_new    Y_new   (T/B)   Net Name
    <label_old>   X_old    Y_old   (T/B)   Net Name
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
    threshold_inch = float(threshold_mil) / 1000.0

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
            f"{label_new}     {new_item['X']:.4f}    {new_item['Y']:.4f}   ({new_item['T/B']})   {new_item['Net Name']}"
        )
        lines.append(
            f"{label_old}     {old_item['X']:.4f}    {old_item['Y']:.4f}   ({old_item['T/B']})   {old_item['Net Name']}"
        )
        lines.append(f"Distance = {distance_inch:.4f} inch ({distance_mil:.1f} mil){mark}")
        lines.append("")  # 空行分隔

    with open(filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"Shift 結果已存到 {filepath} (閾值 = {threshold_mil} mil)")



def find_Nailsasc_Add(CAD_new, CAD_old):
    """
    找出只存在 CAD_new 而 CAD_old 沒有的 Net Name (Add 類別)，
    回傳 list of dict，不做存檔。
    """
    # 建立 Net Name 集合
    new_names = {item["Net Name"] for item in CAD_new}
    old_names = {item["Net Name"] for item in CAD_old}

    # 找出只存在於 CAD_new 的 Net Name
    add_names = new_names - old_names

    # 篩選出 Add 類別的資料
    add_list = [item for item in CAD_new if item["Net Name"] in add_names]

    return add_list


def save_Nails_add_notebook(add_list, filepath=Nails_asc_output, label_new="CAD_new", label_old="CAD_old"):
    """
    將 find_Nailsasc_Add 的結果存成筆記本文字檔
    格式：
    [Part 3] Add Nails
    TOP Side = N
    Bottom Side = M
    Following are in <label_new> Version, but not found in <label_old> Version
    <label_new>   X    Y   (T/B)   Net Name
    """
    lines = []
    lines.append(separator())
    lines.append("[Part 3] Add Nails")

    # 統計 add_list 裡的正面/背面數量
    top_count = sum(1 for item in add_list if item["T/B"].upper() == "T")
    bottom_count = sum(1 for item in add_list if item["T/B"].upper() == "B")

    lines.append(f"TOP Side  = {top_count}")
    lines.append(f"Bottom Side  = {bottom_count}")
    lines.append("")
    lines.append(f"Following are in {label_new} Version, but not found in {label_old} Version")


    for item in add_list:
        lines.append(
            f"{label_new}     {item['X']:.4f}    {item['Y']:.4f}   ({item['T/B']})   {item['Net Name']}"
        )

    lines.append("")  # 空行分隔

    # 續寫到檔案 (append 模式，不覆蓋前面內容)
    with open(filepath, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"Add 結果已續寫到 {filepath}")


def find_Nailsasc_Del(CAD_new, CAD_old):
    """
    找出只存在 CAD_old 而 CAD_new 沒有的 Net Name (Del 類別)，
    回傳 list of dict，不做存檔。
    """
    new_names = {item["Net Name"] for item in CAD_new}
    old_names = {item["Net Name"] for item in CAD_old}

    # 找出只存在於 CAD_old 的 Net Name
    del_names = old_names - new_names

    # 篩選出 Del 類別的資料
    del_list = [item for item in CAD_old if item["Net Name"] in del_names]

    return del_list


def save_Nails_del_notebook(del_list, filepath="Diff_Nails_report.txt", label_new="CAD_new", label_old="CAD_old"):
    """
    將 find_Nailsasc_Del 的結果存成筆記本文字檔
    格式：
    [Part 2] Del Nails
    TOP Side = N
    Bottom Side = M

    Following are in <label_old> Version, but not found in <label_new> Version
    <label_old>   X    Y   (T/B)   Net Name
    """
    lines = []
    lines.append(separator())
    lines.append("[Part 2] Del Nails")

    # 統計 del_list 裡的正面/背面數量
    top_count = sum(1 for item in del_list if item["T/B"].upper() == "T")
    bottom_count = sum(1 for item in del_list if item["T/B"].upper() == "B")

    lines.append(f"TOP Side  = {top_count}")
    lines.append(f"Bottom Side  = {bottom_count}")
    lines.append("")  # 空行分隔
    lines.append(f"Following are in {label_old} Version, but not found in {label_new} Version")

    for item in del_list:
        lines.append(
            f"{label_old}     {item['X']:.4f}    {item['Y']:.4f}   ({item['T/B']})   {item['Net Name']}"
        )

    lines.append("")  # 空行分隔

    # 續寫到檔案 (append 模式)
    with open(filepath, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"Del 結果已續寫到 {filepath}")