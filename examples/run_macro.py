#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
示例：執行 VBA 宏並生成文件

演示如何使用 VBAReader 執行宏代碼，生成文件和處理結果
"""

import sys
import os

# 添加 scripts 目錄到路徑
sys.path.insert(0, os.path.join(
    os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'
))

from vba_reader import VBAReader


def example_run_macro():
    """示例 1：直接執行宏"""
    print("=" * 60)
    print("示例 1：直接執行宏")
    print("=" * 60)

    filepath = r"./Desktop\数据批量套用表格.xltm"

    try:
        # 初始化（必須使用 win32com 模式）
        reader = VBAReader(filepath, use_win32com=True)

        print(f"\n文件: {filepath}")
        print(f"模式: win32com")

        # 列出可用的宏
        procedures = reader.list_procedures()
        print(f"\n可用的宏:")
        for proc in procedures:
            if proc['module'] not in ['ThisWorkbook', 'Sheet1']:
                print(f"  - {proc['name']}")

        # 執行宏
        print(f"\n正在執行宏: FillExcelTemplate")
        print("注意：會彈出用戶界面，請選擇行範圍和模板文件")

        result = reader.run_macro("FillExcelTemplate")

        print(f"\n執行結果:")
        print(f"  狀態: {result['status']}")
        print(f"  耗時: {result['duration']} 秒")

        if result['status'] == 'error':
            print(f"  錯誤: {result['error']}")
        else:
            print(f"  消息: {result['message']}")

            # 獲取生成的文件
            generated_files = reader.get_generated_files()
            print(f"\n生成的文件:")
            for f in generated_files:
                print(f"  - {f}")

        # 關閉
        reader.close()

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")
    except Exception as e:
        print(f"錯誤: {e}")


def example_smart_run():
    """示例 2：智能執行宏"""
    print("\n" + "=" * 60)
    print("示例 2：智能執行宏（自動備份、超時、重試）")
    print("=" * 60)

    filepath = r"./Desktop\数据批量套用表格.xltm"

    try:
        reader = VBAReader(filepath, use_win32com=True)

        # 智能執行
        result = reader.smart_run(
            macro_name="FillExcelTemplate",
            auto_save=True,
            auto_backup=True,
            timeout_seconds=300,  # 5 分鐘超時
            retry_on_error=True,
            max_retries=3,
            log_to_file=True
        )

        print(f"\n執行結果:")
        print(f"  狀態: {result['status']}")
        print(f"  耗時: {result['duration']} 秒")
        print(f"  嘗試次數: {result['attempts']}")

        if result.get('backup_path'):
            print(f"  備份文件: {result['backup_path']}")

        if result['status'] == 'error':
            print(f"  錯誤: {result.get('error')}")

        reader.close()

    except Exception as e:
        print(f"錯誤: {e}")


def example_batch_run():
    """示例 3：批量執行多個宏"""
    print("\n" + "=" * 60)
    print("示例 3：批量執行多個宏")
    print("=" * 60)

    filepath = r"./Desktop\数据批量套用表格.xltm"

    try:
        reader = VBAReader(filepath, use_win32com=True)

        # 定義要執行的宏列表
        macros_to_run = [
            ("FillExcelTemplate", []),  # 無參數
            # 可以添加更多宏
        ]

        print(f"\n將執行 {len(macros_to_run)} 個宏:")

        results = reader.run_macros_batch(macros_to_run)

        print(f"\n執行結果:")
        for i, result in enumerate(results, 1):
            print(f"\n  宏 {i}: {result['macro']}")
            print(f"    狀態: {result['result']['status']}")
            print(f"    耗時: {result['result']['duration']} 秒")

        reader.close()

    except Exception as e:
        print(f"錯誤: {e}")


def example_run_with_params():
    """示例 4：執行帶參數的宏"""
    print("\n" + "=" * 60)
    print("示例 4：執行帶參數的宏")
    print("=" * 60)

    filepath = r"./Desktop\数据批量套用表格.xltm"

    try:
        reader = VBAReader(filepath, use_win32com=True)

        # 獲取工作表對象
        ws = reader.workbook.Sheets(1)

        # 執行帶參數的宏
        result = reader.run_macro_with_params(
            "ExportWorksheetToPDF",
            [ws, "測試報告_20260603"]
        )

        print(f"\n執行結果:")
        print(f"  狀態: {result['status']}")
        print(f"  耗時: {result['duration']} 秒")

        if result['status'] == 'error':
            print(f"  錯誤: {result['error']}")

        reader.close()

    except Exception as e:
        print(f"錯誤: {e}")


if __name__ == "__main__":
    print("""
============================================================
    VBA Macro Runner - 執行宏示例
============================================================

本示例演示如何使用 VBAReader 執行 VBA 宏代碼

注意事項：
1. 執行宏需要使用 win32com 模式（use_win32com=True）
2. 執行宏時 Excel 會被打開（Visible=True）
3. 某些宏可能需要用戶交互（如選（如選擇文件、輸入參數）
4. 執行完成後記得調用 reader.close() 關閉 Excel

============================================================
""")

    # 運行示例
    example_run_macro()

    # 其他示例（可選）
    # example_smart_run()
    # example_batch_run()
    # example_run_with_params()

    print("\n" + "=" * 60)
    print("示例完成！")
    print("=" * 60)
