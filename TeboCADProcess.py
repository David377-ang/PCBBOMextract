import os
from os.path import join, exists
import math
import pandas as pd
from datetime import datetime
from Instance import get_executable_path
from Instance import create_or_replace_file

Nails_shift_threshold = 3  # mil
Parts_shift_threshold = 3  # mil

Nails_asc_name = "Nails.asc"
Nails_asc_output = "Diff_Nails_report.txt"


Parts_asc_name = "Parts.asc"
Parts_asc_output = "Diff_Parts_report.txt"


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

    with open(filepath, "a", encoding="utf-8") as f:
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


def save_Nails_del_notebook(del_list, filepath=Nails_asc_output, label_new="CAD_new", label_old="CAD_old"):
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

def save_Nails_summary_notebook(filepath=Nails_asc_output, label_new="CAD_new", label_old="CAD_old"):
    """
    在報告檔案 Diff_Nails_report.txt 加入 Summary 區塊
    格式：
    ================================================================================================
    Summary :(The Comparison base Version is <label_old>)
    New Version :<label_new>
    Old Version :<label_old>
    ------------------------------------------------------------------------------------------------
    """
    now = datetime.now()
    time_str = now.strftime("%Y/%m/%d %H:%M")


    lines = []
    lines.append(f"       WYMTN Difference Report For Nails       Time {time_str}       Unit:Inch")
    lines.append(separator("="))
    lines.append(f"Summary :(The Comparison base Version is {label_old})")
    lines.append(f"New Version :{label_new}")
    lines.append(f"Old Version :{label_old}")
    lines.append(separator())
    lines.append("")  # 空行分隔

    # 續寫到檔案 (append 模式，不覆蓋前面內容)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"Summary 已續寫到 {filepath}")

def execute_Nails_summary(filepath=Nails_asc_output, label_new="CAD_new", label_old="CAD_old"):

    executable_dir = get_executable_path()
    print(f"執行檔所在目錄: {executable_dir}")

     
    create_or_replace_file(os.path.join(executable_dir, filepath))

    CAD_new = parse_Nailsasc(os.path.join(executable_dir, label_new, Nails_asc_name))
    # print(CAD_new.head())
    # print(CAD_new.tail())

    CAD_old = parse_Nailsasc(os.path.join(executable_dir, label_old, Nails_asc_name))
    # print(CAD_old.tail())

    save_Nails_summary_notebook(filepath, label_new, label_old)

    CAD_Nailsasc_shift = find_Nailsasc_shift(CAD_new, CAD_old)
    save_Nails_shift_notebook(CAD_Nailsasc_shift, filepath, Nails_shift_threshold, label_new, label_old)  # 預設存成 Diff_Nails_report.txt


    CAD_Nailsasc_del = find_Nailsasc_Del(CAD_new, CAD_old)
    save_Nails_del_notebook(CAD_Nailsasc_del, filepath, label_new, label_old)


    CAD_Nailsasc_add = find_Nailsasc_Add(CAD_new, CAD_old)
    save_Nails_add_notebook(CAD_Nailsasc_add, filepath, label_new, label_old)


    return None


def parse_Partsasc(filename):
    parts_list = []
    with open(filename, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("Part"):  # 跳過標題行
                continue

            tokens = line.split()
            try:
                x = float(tokens[1])
                y = float(tokens[2])
                rot = float(tokens[3])
            except ValueError:
                # 如果不是數字，跳過這行
                continue

            part = tokens[0]
            grid = tokens[4]
            tb = tokens[5].strip("()")

            parts_list.append({
                "Part": part,
                "X": x,
                "Y": y,
                "Rot": rot,
                "Grid": grid,
                "T/B": tb
            })
    return parts_list


def save_Parts_summary_notebook(filepath=Parts_asc_output, label_new="CAD_new", label_old="CAD_old"):
    """
    在報告檔案 Diff_Parts_report.txt 加入 Summary 區塊
    格式：
           WYMTN Difference Report For Parts       Time YYYY/MM/DD HH:MM       Unit:Inch

    ================================================================================================
    Summary :(The Comparison base Version is <label_old>)
    New Version :<label_new>
    Old Version :<label_old>
    ------------------------------------------------------------------------------------------------
    """

    now = datetime.now()
    time_str = now.strftime("%Y/%m/%d %H:%M")


    lines = []
    lines.append(f"       WYMTN Difference Report For Partss       Time {time_str}       Unit:Inch")
    lines.append("=" * 100)
    lines.append(f"Summary :(The Comparison base Version is {label_old})")
    lines.append(f"New Version :{label_new}")
    lines.append(f"Old Version :{label_old}")
    lines.append("-" * 100)
    lines.append("")  # 空行分隔

    # 續寫到檔案 (append 模式，不覆蓋前面內容)
    with open(filepath, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"Parts Summary 已續寫到 {filepath}")



def find_Partsasc_shift(CAD_new, CAD_old, threshold_mil=3.0):
    """
    找出同一個 Part 在新舊版本座標或旋轉角度不同的情況 (Shift 類別)
    - 若 XY 距離 >= threshold_mil，列入 Shift
    - 若 Rot 不同，也列入 Shift
    回傳 list of dict，包含新舊座標、旋轉角度與距離
    """
    old_map = {item["Part"]: item for item in CAD_old}
    new_map = {item["Part"]: item for item in CAD_new}

    shift_list = []
    for part, new_item in new_map.items():
        if part in old_map:
            old_item = old_map[part]
            dx = new_item["X"] - old_item["X"]
            dy = new_item["Y"] - old_item["Y"]
            dist_inch = math.sqrt(dx**2 + dy**2)
            dist_mil = dist_inch * 1000.0

            rot_diff = abs(new_item["Rot"] - old_item["Rot"])

            # 判斷是否列入 Shift
            if dist_mil >= threshold_mil or rot_diff > 0.0001:
                shift_list.append({
                    "Part": part,
                    "New": new_item,
                    "Old": old_item,
                    "Distance_inch": dist_inch,
                    "Distance_mil": dist_mil,
                    "Rot_diff": rot_diff
                })
    return shift_list

def save_Parts_shift_notebook(shift_list, filepath="Diff_Parts_report.txt", label_new="CAD_new", label_old="CAD_old"):
    """
    將 find_Partsasc_shift 的結果存成筆記本文字檔
    格式：
    [Part 1] Shift Parts
    TOP Side = N
    Bottom Side = M

    Following xy location of parts be shifted between two version
    <label_new>   Part   X   Y   Rot   Grid   (T/B)
    <label_old>   Part   X   Y   Rot   Grid   (T/B)
    Distance = xxx inch (yyy mil), Rot diff = zzz deg ***
    """
    lines = []
    lines.append("[Part 1] Shift Parts")

    # 統計正面/背面數量
    top_count = sum(1 for item in shift_list if item["New"]["T/B"].upper() == "T")
    bottom_count = sum(1 for item in shift_list if item["New"]["T/B"].upper() == "B")

    lines.append(f"TOP Side  = {top_count}")
    lines.append(f"Bottom Side  = {bottom_count}")
    lines.append("")  # 空行
    lines.append("Following xy location of parts be shifted between two version")

    for item in shift_list:
        new = item["New"]
        old = item["Old"]
        lines.append(f"{label_new}   {new['Part']}   {new['X']:.4f}   {new['Y']:.4f}   {new['Rot']:.1f}   {new['Grid']}   ({new['T/B']})")
        lines.append(f"{label_old}   {old['Part']}   {old['X']:.4f}   {old['Y']:.4f}   {old['Rot']:.1f}   {old['Grid']}   ({old['T/B']})")
        lines.append(f"Distance = {item['Distance_inch']:.4f} inch ({item['Distance_mil']:.1f} mil), Rot diff = {item['Rot_diff']:.1f} deg ***")
        lines.append("")  # 每組之間空行

    lines.append("")  # 區塊結尾空行

    with open(filepath, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"Shift 結果已續寫到 {filepath}")


def find_Partsasc_Del(CAD_new, CAD_old):
    """
    找出只存在 CAD_old 而 CAD_new 沒有的 Part (Del 類別)
    回傳 list of dict
    """
    new_parts = {item["Part"] for item in CAD_new}
    old_parts = {item["Part"] for item in CAD_old}

    del_parts = old_parts - new_parts
    del_list = [item for item in CAD_old if item["Part"] in del_parts]

    return del_list

def save_Parts_del_notebook(del_list, filepath=Parts_asc_output, label_new="CAD_new", label_old="CAD_old"):
    """
    將 find_Partsasc_Del 的結果存成筆記本文字檔
    格式：
    [Part 2] Del Parts
    TOP Side = N
    Bottom Side = M

    Following are in <label_old> Version, but not found in <label_new> Version
    <label_old>   Part   X   Y   Rot   Grid   (T/B)
    """
    lines = []
    lines.append(separator())
    lines.append("[Part 2] Del Parts")

    top_count = sum(1 for item in del_list if item["T/B"].upper() == "T")
    bottom_count = sum(1 for item in del_list if item["T/B"].upper() == "B")

    lines.append(f"TOP Side  = {top_count}")
    lines.append(f"Bottom Side  = {bottom_count}")
    lines.append("")  # 空行
    lines.append(f"Following are in {label_old} Version, but not found in {label_new} Version")

    for item in del_list:
        lines.append(
            f"{label_old}   {item['Part']}   {item['X']:.4f}   {item['Y']:.4f}   {item['Rot']:.1f}   {item['Grid']}   ({item['T/B']})"
        )

    lines.append("")  # 區塊結尾空行

    with open(filepath, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"Del 結果已續寫到 {filepath}")


def find_Partsasc_Add(CAD_new, CAD_old):
    """
    找出只存在 CAD_new 而 CAD_old 沒有的 Part (Add 類別)
    回傳 list of dict
    """
    new_parts = {item["Part"] for item in CAD_new}
    old_parts = {item["Part"] for item in CAD_old}

    add_parts = new_parts - old_parts
    add_list = [item for item in CAD_new if item["Part"] in add_parts]

    return add_list



def save_Parts_add_notebook(add_list, filepath=Parts_asc_output, label_new="CAD_new", label_old="CAD_old"):
    """
    將 find_Partsasc_Add 的結果存成筆記本文字檔
    格式：
    [Part 3] Add Parts
    TOP Side = N
    Bottom Side = M

    Following are in <label_new> Version, but not found in <label_old> Version
    <label_new>   Part   X   Y   Rot   Grid   (T/B)
    """
    lines = []
    lines.append(separator())
    lines.append("[Part 3] Add Parts")

    top_count = sum(1 for item in add_list if item["T/B"].upper() == "T")
    bottom_count = sum(1 for item in add_list if item["T/B"].upper() == "B")

    lines.append(f"TOP Side  = {top_count}")
    lines.append(f"Bottom Side  = {bottom_count}")
    lines.append("")  # 空行
    lines.append(f"Following are in {label_new} Version, but not found in {label_old} Version")

    for item in add_list:
        lines.append(
            f"{label_new}   {item['Part']}   {item['X']:.4f}   {item['Y']:.4f}   {item['Rot']:.1f}   {item['Grid']}   ({item['T/B']})"
        )

    lines.append("")  # 區塊結尾空行

    with open(filepath, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"Add 結果已續寫到 {filepath}")


def execute_Parts_summary(filepath=Parts_asc_output, label_new="CAD_new", label_old="CAD_old"):

    executable_dir = get_executable_path()
    print(f"執行檔所在目錄: {executable_dir}")

     
    create_or_replace_file(os.path.join(executable_dir, filepath))

    CAD_new = parse_Partsasc(os.path.join(executable_dir, label_new, Parts_asc_name))
    print("總筆數 =", len(CAD_new))

    CAD_old = parse_Partsasc(os.path.join(executable_dir, label_old, Parts_asc_name))
    print("總筆數 =", len(CAD_old))

    save_Parts_summary_notebook(filepath, label_new, label_old)

    CAD_Partsasc_shift = find_Partsasc_shift(CAD_new, CAD_old, Parts_shift_threshold)
    save_Parts_shift_notebook(CAD_Partsasc_shift, filepath, label_new, label_old)

    CAD_Partsasc_Del = find_Partsasc_Del(CAD_new, CAD_old)
    save_Parts_del_notebook(CAD_Partsasc_Del, filepath, label_new, label_old)

    CAD_Partsasc_Add = find_Partsasc_Add(CAD_new, CAD_old)
    save_Parts_add_notebook(CAD_Partsasc_Add, filepath, label_new, label_old)


    return None