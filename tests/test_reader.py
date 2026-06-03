#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試 VBA Reader
"""

import sys
import os
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.vba_reader import VBAReader


def test_oletools_mode():
    """測試 oletools 只讀模式"""
    print("測試 oletools 模式...")

    # 使用測試文件（如果存在）
    test_file = r"./Desktop\工程總表宏\工程總表格宏.xlsm"

    if not os.path.exists(test_file):
        print(f"跳過測試：文件不存在 {test_file}")
        return

    try:
        with VBAReader(test_file, use_win32com=False) as reader:
            # 測試列出模塊
            modules = reader.list_modules()
            print(f"✓ 找到 {len(modules)} 個模塊")

            # 測試讀取代碼
            for module_name in modules:
                code = reader.get_module(module_name)
                assert code, f"無法讀取 {module_name}"
                print(f"✓ 讀取成功: {module_name} ({len(code)} 字符)")

            # 測試列出過程
            procedures = reader.list_procedures()
            print(f"✓ 找到 {len(procedures)} 個過程")

            # 測試分析代碼
            for module_name in modules:
                analysis = reader.analyze_code(module_name)
                assert 'line_count' in analysis
                print(f"✓ 分析成功: {module_name}")

        print("\n✓ oletools 模式測試通過\n")

    except Exception as e:
        print(f"✗ 測試失敗: {e}")
        raise


def test_win32com_mode():
    """測試 win32com 模式"""
    print("測試 win32com 模式...")

    # 使用測試文件
    test_file = r"./Desktop\工程總表宏\工程總表格宏.xlsm"

    if not os.path.exists(test_file):
        print(f"跳過測試：文件不存在 {test_file}")
        return

    try:
        with VBAReader(test_file, use_win32com=True) as reader:
            # 測試列出模塊
            modules = reader.list_modules()
            print(f"✓ 找到 {len(modules)} 個模塊")

            # 測試讀取代碼
            for module_name in modules:
                code = reader.get_module(module_name)
                assert code, f"無法讀取 {module_name}"
                print(f"✓ 讀取成功: {module_name}")

        print("\n✓ win32com 模式測試通過\n")

    except ImportError as e:
        print(f"跳過測試：{e}")
    except Exception as e:
        print(f"✗ 測試失敗: {e}")
        raise


if __name__ == "__main__":
    print("=" * 60)
    print("VBA Reader 測試")
    print("=" * 60 + "\n")

    test_oletools_mode()
    test_win32com_mode()

    print("=" * 60)
    print("所有測試完成！")
    print("=" * 60)
