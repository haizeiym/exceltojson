import pandas as pd
import json
import os
import sys

def process_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, str):
        try:
            v = eval(value)
            if isinstance(v, set):
                return list(v)
            return v
        except:
            return value
    if isinstance(value, set):
        return list(value)
    return value

def excel_tojson_side(df, export_columns, dir_name, file_name, side, output_root, indent_val=2):
    # 如果只有id字段则不导出
    if len(export_columns) == 1 and export_columns[0][1] == "id":
        print(f"仅有id字段，跳过导出: {file_name} ({side})")
        return
    data = []
    for row in range(4, len(df)):
        obj = {}
        for col, key in export_columns:
            value = process_value(df.iloc[row, col])
            if value is not None:
                obj[key] = value
        if obj:
            data.append(obj)
    if not data or (len(data[0]) == 1 and "id" in data[0]):
        print(f"仅有id字段数据，跳过导出: {file_name} ({side})")
        return

    # 检查是否有game_id字段
    has_game_id = any(key == "game_id" for _, key in export_columns)
    
    # 按game_id分组数据
    if has_game_id:
        grouped_data = {}
        for item in data:
            game_id = item.get("game_id")
            if game_id is not None:
                if game_id not in grouped_data:
                    grouped_data[game_id] = []
                grouped_data[game_id].append(item)
        
        # 导出每个game_id的数据
        for game_id, game_data in grouped_data.items():
            export_dir = os.path.join(output_root, "output", str(game_id), dir_name, side)
            os.makedirs(export_dir, exist_ok=True)
            json_file = os.path.join(export_dir, f"{file_name}.json")
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(game_data, f, ensure_ascii=False, indent=indent_val)
            print(f"导出: {json_file}")
    else:
        # 如果没有game_id，使用原来的导出路径
        export_dir = os.path.join(output_root, "output", dir_name, side)
        os.makedirs(export_dir, exist_ok=True)
        json_file = os.path.join(export_dir, f"{file_name}.json")
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=indent_val)
        print(f"导出: {json_file}")


def excel_to_json(excel_file: str, output_root: str):
    all_sheets = pd.read_excel(excel_file, header=None, sheet_name=None, engine="openpyxl")
    for sheet_name, df in all_sheets.items():
        # 跳过sheet名或首行有#的sheet
        if '#' in str(sheet_name) or any('#' in str(cell) for cell in df.iloc[0]):
            print(f"跳过含#的sheet: {sheet_name}")
            continue
        export_file_col = None
        export_dir_col = None
        indent_val = 2
        for idx, v in enumerate(df.iloc[0]):
            if str(v).strip() == '导出文件头':
                export_file_col = idx
            if str(v).strip() == '导出目录':
                export_dir_col = idx
            if str(v).strip() == '缩进':
                try:
                    val = df.iloc[0, idx+1]
                    indent_val = int(float(str(val).strip()))
                except Exception:
                    indent_val = 2
        if export_file_col is None or export_dir_col is None:
            print(f'未找到导出文件头或导出目录列 in sheet {sheet_name}')
            continue
        
        file_name = str(df.iloc[0, export_file_col+1]).strip()
        dir_name = str(df.iloc[0, export_dir_col+1]).strip()
       
        client_columns = []
        server_columns = []
        for col in range(df.shape[1]):
            if str(df.iloc[1, col]).strip().startswith('#'):
                continue
            flag = str(df.iloc[2, col]).strip().lower()
            key = str(df.iloc[3, col]).strip()
            if key != '' and not key.startswith('#'):
                if flag in ("c", "sc"):
                    client_columns.append((col, key))
                if flag in ("s", "sc"):
                    server_columns.append((col, key))
                    
        if client_columns:
            excel_tojson_side(df, client_columns, dir_name, file_name, "client", output_root, indent_val)
        if server_columns:
            excel_tojson_side(df, server_columns, dir_name, file_name, "server", output_root, indent_val)

if __name__ == "__main__":
    folder = sys.argv[1] if len(sys.argv) > 1 else '.'
    output_root = os.path.abspath(folder)
    for fname in os.listdir(folder):
        if fname.endswith('.xlsx') and not fname.startswith('~') and not fname.startswith('.~'):
            fpath = os.path.join(folder, fname)
            print(f"处理文件: {fpath}")
            excel_to_json(fpath, output_root)
