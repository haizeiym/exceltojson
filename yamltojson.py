#!/usr/bin/env python3
import os
import argparse
import yaml
import json


def convert_yaml_to_json(input_dir, output_dir):
    if not os.path.exists(input_dir):
        print(f"输入目录不存在: {input_dir}")
        return
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.endswith('.yaml') or file.endswith('.yml'):
                yaml_path = os.path.join(root, file)
                rel_path = os.path.relpath(yaml_path, input_dir)
                json_filename = os.path.splitext(rel_path)[0] + '.json'
                json_path = os.path.join(output_dir, json_filename)
                json_dir = os.path.dirname(json_path)
                if not os.path.exists(json_dir):
                    os.makedirs(json_dir)
                try:
                    with open(yaml_path, 'r', encoding='utf-8') as yf:
                        data = yaml.safe_load(yf)
                    with open(json_path, 'w', encoding='utf-8') as jf:
                        json.dump(data, jf, ensure_ascii=False, indent=2)
                    print(f"转换成功: {yaml_path} -> {json_path}")
                except Exception as e:
                    print(f"转换失败: {yaml_path}，错误: {e}")

def main():
    parser = argparse.ArgumentParser(description='批量将yaml/yml文件转换为json文件')
    parser.add_argument('-i', '--input', required=True, help='输入目录')
    parser.add_argument('-o', '--output', required=False, help='输出目录，默认与输入目录相同')
    args = parser.parse_args()
    output_dir = args.output if args.output else args.input
    convert_yaml_to_json(args.input, output_dir)

if __name__ == '__main__':
    main()
