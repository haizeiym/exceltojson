import pandas as pd
import json
import os

def process_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, str):
        try:
            return eval(value)
        except:
            return value
    return value

def excel_tojson_side(df, export_columns, dir_name, file_name, side):
    data = []
    for row in range(4, len(df)):
        obj = {}
        for col, key in export_columns:
            value = process_value(df.iloc[row, col])
            if value is not None:
                obj[key] = value
        if obj:
            data.append(obj)
    export_dir = os.path.join(dir_name, side)
    os.makedirs(export_dir, exist_ok=True)
    json_file = os.path.join(export_dir, f"{file_name}.json")
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"导出: {json_file}")


def excel_to_json(excel_file: str):
    df = pd.read_excel(excel_file, header=None)
    # 找到"导出文件头"和"导出目录"所在列
    export_file_col = None
    export_dir_col = None
    for idx, v in enumerate(df.iloc[0]):
        if str(v).strip() == '导出文件头':
            export_file_col = idx
        if str(v).strip() == '导出目录':
            export_dir_col = idx
    if export_file_col is None or export_dir_col is None:
        print('未找到导出文件头或导出目录列')
        return
    
    # 文件名和目录在其右侧一列
    file_name = str(df.iloc[0, export_file_col+1]).strip()
    dir_name = str(df.iloc[0, export_dir_col+1]).strip()
   
    client_columns = []
    server_columns = []
    # 收集所有flag为s/c/sc的列及其key
    export_columns = []
    for col in range(df.shape[1]):
        flag = str(df.iloc[2, col]).strip().lower()
        key = str(df.iloc[3, col]).strip()
        if flag in ("s", "c", "sc") and key != '' and not key.startswith('#'):
            export_columns.append((col, key))
            if flag in ("c", "sc"):
                client_columns.append((col, key))
            if flag in ("s", "sc"):
                server_columns.append((col, key))
    # # 组装对象数组
    # data = []
    # for row in range(4, len(df)):
    #     obj = {}
    #     for col, key in export_columns:
    #         value = process_value(df.iloc[row, col])
    #         if value is not None:
    #             obj[key] = value
    #     if obj:
    #         data.append(obj)
    # # 导出
    # for location in (["server"] if str(df.iloc[2, export_file_col+1]).strip().lower()=="s" else ["client"] if str(df.iloc[2, export_file_col+1]).strip().lower()=="c" else ["server","client"]):
    #     export_dir = os.path.join(dir_name, location)
    #     os.makedirs(export_dir, exist_ok=True)
    #     json_file = os.path.join(export_dir, f"{file_name}.json")
    #     with open(json_file, 'w', encoding='utf-8') as f:
    #         json.dump(data, f, ensure_ascii=False, indent=2)
    #     print(f"导出: {json_file}")

    # 输出
    if client_columns:
        excel_tojson_side(df, client_columns, dir_name, file_name, "client")
    if server_columns:
        excel_tojson_side(df, server_columns, dir_name, file_name, "server")

if __name__ == "__main__":
    excel_file = "./乐享捕鱼.xlsx"
    excel_to_json(excel_file)
