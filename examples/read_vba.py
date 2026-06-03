#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
示例：讀取工程總表宏系統的 VBA 代碼

演示如何使用 VBAReader 讀取 .xlsm 文件中的宏代碼
"""

import sys
import os

# 添加 scripts 目錄到路徑
sys.path.insert(0, os.path.join(
    os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'
))

from vba_reader import VBAReader


def example_read_all_modules():
    """示例 1：讀取所有模塊"""
    print("=" * 60)
    print("示例 1：讀取所有模塊")
    print("=" * 60)

    filepath = r"./Desktop\工程總表宏\工程總表格宏.xlsm"

    try:
        with VBAReader(filepath, use_win32com=False) as reader:
            modules = reader.get_all_modules()

            print(f"\n找到 {len(modules)} 個模塊：")
            for name in modules.keys():
                print(f"  - {name}")

            # 保存到文件
            output_dir = r"./Desktop\工程總表宏\vba_backup"
            os.makedirs(output_dir, exist_ok=True)

            for name, code in modules.items():
                filename = f"{name}.bas"
                filepath = os.path.join(output_dir, filename)

                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(code)

                print(f"已保存: {filename}")

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")
        print("請修改 filepath 為實際路徑")


def example_read_specific_module():
    """示例 2：讀取特定模塊"""
    print("\n" + "=" * 60)
    print("示例 2：讀取特定模塊（mod_Main）")
    print("=" * 60)

    filepath = r"./Desktop\工程總表宏\工程總表格宏.xlsm"

    try:
        with VBAReader(filepath, use_win32com=False) as reader:
            code = reader.get_module("mod_Main")

            if code:
                print(f"\n模塊代碼（前 500 字符）：")
                print(code[:500])

                print(f"\n完整代碼長度: {len(code)} 字符")
            else:
                print("模塊不存在")

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")


def example_list_procedures():
    """示例 3：列出所有過程"""
    print("\n" + "=" * 60)
    print("示例 3：列出所有 Sub/Function")
    print("=" * 60)

    filepath = r"./Desktop\工程總表宏\工程總表格宏.xlsm"

    try:
        with VBAReader(filepath, use_win32com=False) as reader:
            procedures = reader.list_procedures()

            print(f"\n找到 {len(procedures)} 個過程：\n")

            for proc in procedures[:10]:  # 只顯示前 10 個
                params = ", ".join(proc['params']) if proc['params'] else "無參數"
                print(f"{proc['type']:8} {proc['module']}.{proc['name']}({params})")

            if len(procedures) > 10:
                print(f"\n... 還有 {len(procedures) - 10} 個過程")

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")


def example_analyze_code():
    """示例 4：分析代碼質量"""
    print("\n" + "=" * 60)
    print("示例 4：分析代碼質量")
    print("=" * 60)

    filepath = r"./Desktop\工程總表宏\工程總表格宏.xlsm"

    try:
        with VBAReader(filepath, use_win32com=False) as reader:
            modules = reader.list_modules()

            print(f"\n分析 {len(modules)} 個模塊：\n")

            for module_name in modules:
                analysis = reader.analyze_code(module_name)

                print(f"模塊: {module_name}")
                print(f"  代碼行數: {analysis.get('line_count', 0)}")
                print(f"  複雜度: {analysis.get('complexity', 0)}")
                print(f"  圈複雜度: {analysis.get('cyclomatic', 0)}")

                if analysis.get('suggestions'):
                    print(f"  建議:")
                    for suggestion in analysis['suggestions']:
                        print(f"    - {suggestion}")

                print()

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")


def example_extract_parameters():
    """示例 5：提取過程參數"""
    print("\n" + "=" * 60)
    print("示例 5：提取過程參數")
    print("=" * 60)

    filepath = r"./Desktop\工程總表宏\工程總表格宏.xlsm"

    try:
        with VBAReader(filepath, use_win32com=False) as reader:
            # 提取 FillTemplate 過程的參數
            params = reader.extract_parameters("mod_Main", "FillTemplate")

            print(f"\nFillTemplate 過程參數：")
            for i, param in enumerate(params, 1):
                print(f"  {i}. {param}")

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")


if __name__ == "__main__":
    # 運行所有示例
    example_read_all_modules()
    example_read_specific_module()
    example_list_procedures()
    example_analyze_code()
    example_extract_parameters()

    print("\n" + "=" * 60)
    print("所有示例完成！")
    print("=" * 60)
