#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VBA Macro Reader - 讀取 .xlsm/.xlam 文件中的 VBA 宏代碼

支持兩種模式：
1. win32com 模式（可讀可寫，需要 Excel）
2. oletools 模式（只讀，不需要 Excel）

作者：QClaw Agent
日期：2026-06-03
"""

import os
import re
import json
from typing import Dict, List, Optional, Tuple
from pathlib import Path


class VBAReader:
    """VBA 宏代碼讀取器"""

    def __init__(self, filepath: str, use_win32com: bool = True):
        """
        初始化 VBA 讀取器

        Args:
            filepath: .xlsm 文件路徑
            use_win32com: True 使用 win32com（可讀可寫）
                         False 使用 oletools（只讀）
        """
        self.filepath = Path(filepath)
        self.use_win32com = use_win32com
        self.excel = None
        self.workbook = None
        self.vbaparser = None
        self._file_snapshot = set()  # 用於 get_generated_files 的快照對比

        if not self.filepath.exists():
            raise FileNotFoundError(f"文件不存在: {filepath}")

        if use_win32com:
            self._init_win32com()
        else:
            self._init_oletools()

    def _take_snapshot(self) -> set:
        """拍攝文件所在目錄的快照，返回文件路徑集合"""
        import os
        dir_path = os.path.dirname(str(self.filepath))
        if not dir_path:
            dir_path = "."
        return set(
            os.path.join(dir_path, f)
            for f in os.listdir(dir_path)
            if os.path.isfile(os.path.join(dir_path, f))
        )

    def _init_win32com(self):
        """初始化 win32com"""
        try:
            import win32com.client

            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False

            self.workbook = self.excel.Workbooks.Open(str(self.filepath.absolute()))

        except ImportError:
            raise ImportError("未安裝 pywin32，請運行: pip install pywin32")
        except Exception as e:
            raise RuntimeError(f"無法打開文件: {e}")

    def _init_oletools(self):
        """初始化 oletools"""
        try:
            from oletools.olevba import VBA_Parser

            self.vbaparser = VBA_Parser(str(self.filepath.absolute()))

        except ImportError:
            raise ImportError("未安裝 oletools，請運行: pip install oletools")
        except Exception as e:
            raise RuntimeError(f"無法解析文件: {e}")

    def get_all_modules(self) -> Dict[str, str]:
        """
        返回所有 VBA 模塊代碼

        Returns:
            Dict[模塊名, 代碼內容]
        """
        if self.use_win32com:
            return self._get_modules_win32com()
        else:
            return self._get_modules_oletools()

    def _get_modules_win32com(self) -> Dict[str, str]:
        """使用 win32com 讀取所有模塊"""
        modules = {}

        try:
            vb_project = self.workbook.VBProject

            for comp in vb_project.VBComponents:
                comp_type = comp.Type

                # 1=標準模塊, 2=類模塊, 3=窗體
                if comp_type in [1, 2, 3]:
                    code_module = comp.CodeModule
                    line_count = code_module.CountOfLines

                    if line_count > 0:
                        code = code_module.Lines(1, line_count)
                        modules[comp.Name] = code

        except Exception as e:
            raise RuntimeError(f"讀取 VBA 模塊失敗: {e}")

        return modules

    def _get_modules_oletools(self) -> Dict[str, str]:
        """使用 oletools 讀取所有模塊"""
        modules = {}

        try:
            for (filename, stream_path, vba_filename, vba_code) in self.vbaparser.extract_macros():
                if vba_filename and vba_code:
                    # 清理模塊名
                    module_name = vba_filename.replace(".bas", "").replace(".cls", "")
                    modules[module_name] = vba_code

        except Exception as e:
            raise RuntimeError(f"解析 VBA 失敗: {e}")

        return modules

    def get_module(self, module_name: str) -> Optional[str]:
        """
        返回指定模塊代碼

        Args:
            module_name: 模塊名稱

        Returns:
            模塊代碼，如果不存在返回 None
        """
        modules = self.get_all_modules()
        return modules.get(module_name)

    def list_modules(self) -> List[str]:
        """
        列出所有模塊名稱

        Returns:
            模塊名稱列表
        """
        return list(self.get_all_modules().keys())

    def list_procedures(self) -> List[Dict]:
        """
        列出所有 Sub/Function 及參數

        Returns:
            過程列表: [{"type": "Sub|Function", "name": "...", "params": [...]}]
        """
        procedures = []
        modules = self.get_all_modules()

        for module_name, code in modules.items():
            # 提取 Sub
            sub_pattern = r'Sub\s+(\w+)\s*\(([^)]*)\)'
            for match in re.finditer(sub_pattern, code, re.IGNORECASE):
                params_str = match.group(2).strip()
                params = [p.strip() for p in params_str.split(',') if p.strip()]

                procedures.append({
                    "module": module_name,
                    "type": "Sub",
                    "name": match.group(1),
                    "params": params
                })

            # 提取 Function
            func_pattern = r'Function\s+(\w+)\s*\(([^)]*)\)'
            for match in re.finditer(func_pattern, code, re.IGNORECASE):
                params_str = match.group(2).strip()
                params = [p.strip() for p in params_str.split(',') if p.strip()]

                procedures.append({
                    "module": module_name,
                    "type": "Function",
                    "name": match.group(1),
                    "params": params
                })

        return procedures

    def extract_parameters(self, module_name: str, procedure_name: str) -> List[str]:
        """
        提取指定過程的參數列表

        Args:
            module_name: 模塊名稱
            procedure_name: 過程名稱

        Returns:
            參數列表
        """
        procedures = self.list_procedures()

        for proc in procedures:
            if proc["module"] == module_name and proc["name"] == procedure_name:
                return proc["params"]

        return []

    def update_module(self, module_name: str, code: str) -> bool:
        """
        更新模塊代碼（僅 win32com）

        Args:
            module_name: 模塊名稱
            code: 新代碼

        Returns:
            是否成功
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持寫入操作")

        try:
            vb_project = self.workbook.VBProject

            for comp in vb_project.VBComponents:
                if comp.Name == module_name:
                    code_module = comp.CodeModule
                    line_count = code_module.CountOfLines

                    # 刪除舊代碼
                    if line_count > 0:
                        code_module.DeleteLines(1, line_count)

                    # 寫入新代碼
                    code_module.AddFromString(code)
                    return True

            # 如果模塊不存在，創建新模塊
            return self.add_module(module_name, code)

        except Exception as e:
            raise RuntimeError(f"更新模塊失敗: {e}")

    def add_module(self, module_name: str, code: str) -> bool:
        """
        新增模增模塊（僅 win32com）

        Args:
            module_name: 模塊名稱
            code: 代碼

        Returns:
            是否成功
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持寫入操作")

        try:
            import win32com.client

            vb_project = self.workbook.VBProject

            # 創建新模塊
            new_comp = vb_project.VBComponents.Add(1)  # 1=標準模塊
            new_comp.Name = module_name
            new_comp.CodeModule.AddFromString(code)

            return True

        except Exception as e:
            raise RuntimeError(f"新增模塊失敗: {e}")

    def delete_module(self, module_name: str) -> bool:
        """
        刪除模塊（僅 win32com）

        Args:
            module_name: 模塊名稱

        Returns:
            是否成功
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持寫入操作")

        try:
            vb_project = self.workbook.VBProject

            for comp in vb_project.VBComponents:
                if comp.Name == module_name:
                    vb_project.VBComponents.Remove(comp)
                    return True

            return False

        except Exception as e:
            raise RuntimeError(f"刪除模塊失敗: {e}")

    def analyze_code(self, module_name: str) -> Dict:
        """
        分析代碼質量

        Args:
            module_name: 模塊名稱

        Returns:
            分析結果
        """
        code = self.get_module(module_name)
        if not code:
            return {}

        lines = code.split('\n')
        code_lines = [l.strip() for l in lines if l.strip() and not l.strip().startswith("'")]

        # 圈複雜度：每個決策點 +1（排除註釋行中的關鍵字）
        decision_keywords = [
            'If ', 'ElseIf ', 'For ', 'For Each ',
            'While ', 'Do While', 'Do Until',
            'Select Case',
        ]
        cyclomatic = 1  # 基礎複雜度
        for line in code_lines:
            for kw in decision_keywords:
                if kw in line:
                    cyclomatic += 1
            # And/Or 也算分支（排除 Dim 宣告行）
            if not line.startswith('Dim ') and not line.startswith('Public '):
                if ' And ' in line:
                    cyclomatic += 1
                if ' Or ' in line:
                    cyclomatic += 1

        # 嵌套深度
        max_depth = 0
        current_depth = 0
        depth_openers = ['If ', 'For ', 'While ', 'Do ', 'Select Case', 'With ']
        depth_closers = ['End If', 'Next', 'Wend', 'Loop', 'End Select', 'End With']
        for line in code_lines:
            for opener in depth_openers:
                if opener in line:
                    current_depth += 1
                    break
            for closer in depth_closers:
                if closer in line:
                    current_depth = max(0, current_depth - 1)
                    break
            max_depth = max(max_depth, current_depth)

        # 建議
        suggestions = []

        if cyclomatic > 15:
            suggestions.append(f"圈複雜度 {cyclomatic} 過高，建議拆分成多個小函數")
        elif cyclomatic > 10:
            suggestions.append(f"圈複雜度 {cyclomatic} 偏高，考慮簡化條件邏輯")

        if max_depth > 4:
            suggestions.append(f"嵌套深度 {max_depth} 過深，建議提取子過程或使用提前返回")

        if 'On Error Resume Next' in code:
            suggestions.append("使用 On Error Resume Next 可能隱藏錯誤，建議使用明確的錯誤處理")

        if 'Select *' in code or 'select *' in code.lower():
            suggestions.append("避免使用 SELECT *，建議明確指定字段")

        # 檢查是否有註釋
        comment_lines = [l for l in lines if l.strip().startswith("'")]
        if len(comment_lines) < len(code_lines) * 0.1:
            suggestions.append("註釋不足，建議增加代碼註釋")

        return {
            "line_count": len(lines),
            "code_lines": len(code_lines),
            "cyclomatic_complexity": cyclomatic,
            "max_nesting_depth": max_depth,
            "has_error_handling": 'On Error' in code,
            "suggestions": suggestions
        }

    def optimize_module(self, module_name: str) -> str:
        """
        返回優化後的代碼（建議性優化）

        Args:
            module_name: 模塊名稱

        Returns:
            優化後的代碼
        """
        code = self.get_module(module_name)
        if not code:
            return ""

        lines = code.split('\n')
        optimized_lines = []

        has_option_explicit = any(
            'Option Explicit' in line for line in lines
        )

        # 若缺少 Option Explicit，在文件頂部插入（VBA 要求在所有過程之前）
        if not has_option_explicit:
            optimized_lines.append("Option Explicit")
            optimized_lines.append("")  # 空行分隔

        for line in lines:
            stripped = line.strip()
            # 對高密度屬性訪問行添加建議（排除註釋、聲明行）
            if (not stripped.startswith("'")
                    and not stripped.startswith("Dim ")
                    and not stripped.startswith("Public ")
                    and not stripped.startswith("Private ")
                    and line.count('.') > 3
                    and 'With' not in line):
                optimized_lines.append(
                    f"' OPTIMIZE: 考慮使用 With 語句 — {stripped}"
                )
            optimized_lines.append(line.rstrip())

        return '\n'.join(optimized_lines)

    def save(self) -> bool:
        """
        保存修改（僅 win32com）

        Returns:
            是否成功
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持保存操作")

        try:
            self.workbook.Save()
            return True
        except Exception as e:
            raise RuntimeError(f"保存失敗: {e}")

    # ==================== ★執行宏功能 ====================

    def run_macro(self, macro_name: str) -> Dict:
        """
        執行宏代碼（無參數）

        Args:
            macro_name: 宏名稱（如 "FillExcelTemplate"）

        Returns:
            執行結果: {"status": "success/error", "message": "...", "duration": 秒}
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持執行宏")

        import time
        start_time = time.time()

        # 拍攝快照以追蹤生成的文件
        self._file_snapshot = self._take_snapshot()

        try:
            # 確保 Excel 可見（可選）
            self.excel.Visible = True

            # 執行宏
            self.excel.Run(f"'{self.workbook.Name}'!{macro_name}")

            duration = time.time() - start_time

            return {
                "status": "success",
                "message": f"宏 {macro_name} 執行成功",
                "duration": round(duration, 2)
            }

        except Exception as e:
            duration = time.time() - start_time
            return {
                "status": "error",
                "error": str(e),
                "duration": round(duration, 2)
            }

    def run_macro_with_params(self, macro_name: str, params: List) -> Dict:
        """
        執行宏代碼並傳遞參數

        Args:
            macro_name: 宏名稱
            params: 參數列表

        Returns:
            執行結果
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持執行宏")

        import time
        start_time = time.time()

        # 拍攝快照以追蹤生成的文件
        self._file_snapshot = self._take_snapshot()

        try:
            self.excel.Visible = True

            # 構建參數字符串
            param_str = ", ".join([repr(p) if isinstance(p, str) else str(p) for p in params])

            # 執行帶參數的宏
            result = self.excel.Run(f"'{self.workbook.Name}'!{macro_name}", *params)

            duration = time.time() - start_time

            return {
                "status": "success",
                "message": f"宏 {macro_name} 執行成功",
                "result": result,
                "duration": round(duration, 2)
            }

        except Exception as e:
            duration = time.time() - start_time
            return {
                "status": "error",
                "error": str(e),
                "duration": round(duration, 2)
            }

    def run_macro_monitored(self, macro_name: str, timeout_seconds: int = 300, log_file: str = None) -> Dict:
        """
        執行宏並監控執行過程

        Args:
            macro_name: 宏名稱
            timeout_seconds: 超時秒數
            log_file: 日誌文件路徑

        Returns:
            執行結果
        """
        import time
        import threading

        # 拍攝快照以追蹤生成的文件
        self._file_snapshot = self._take_snapshot()

        result = {"status": "running", "start_time": time.time()}

        def run_macro_thread():
            import pythoncom
            pythoncom.CoInitialize()  # COM 線程安全初始化
            try:
                self.excel.Visible = True
                self.excel.Run(f"'{self.workbook.Name}'!{macro_name}")
                result["status"] = "success"
                result["end_time"] = time.time()
            except Exception as e:
                result["status"] = "error"
                result["error"] = str(e)
                result["end_time"] = time.time()
            finally:
                pythoncom.CoUninitialize()

        # 啟動線程
        thread = threading.Thread(target=run_macro_thread)
        thread.start()

        # 等待完成或超時
        thread.join(timeout=timeout_seconds)

        if thread.is_alive():
            result["status"] = "timeout"
            result["message"] = f"執行超時（>{timeout_seconds}秒）"

        result["duration"] = time.time() - result["start_time"]

        # 寫入日誌
        if log_file:
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] {macro_name}: {result}\n")

        return result

    def run_macros_batch(self, macros: List[Tuple[str, List]]) -> List[Dict]:
        """
        批量執行多個宏

        Args:
            macros: 宏列表 [("宏名", [參數列表]), ...]

        Returns:
            執行結果列表
        """
        results = []

        for macro_name, params in macros:
            if params:
                result = self.run_macro_with_params(macro_name, params)
            else:
                result = self.run_macro(macro_name)

            results.append({
                "macro": macro_name,
                "result": result
            })

        return results

    def smart_run(self, macro_name: str, auto_save: bool = True, auto_backup: bool = True,
                  timeout_seconds: int = 300, retry_on_error: bool = False,
                  max_retries: int = 3, log_to_file: bool = False) -> Dict:
        """
        智能執行宏（自動處理參數、監控、錯誤處理）

        Args:
            macro_name: 宏名稱
            auto_save: 執行後自動保存
            auto_backup: 執行前自動備份
            timeout_seconds: 超時秒數
            retry_on_error: 錯誤時自動重試
            max_retries: 最大重試次數
            log_to_file: 記錄到日誌文件

        Returns:
            執行結果
        """
        import time
        import shutil
        from datetime import datetime

        # 執行前備份
        backup_path = None
        if auto_backup:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = str(self.filepath).replace(".xlsm", f"_backup_{timestamp}.xlsm")
            try:
                shutil.copy2(self.filepath, backup_path)
            except:
                pass

        # 執行宏
        attempts = 0
        result = None
        log_path = None
        if log_to_file:
            import os
            from datetime import datetime
            log_dir = os.path.dirname(str(self.filepath))
            if not log_dir:
                log_dir = "."
            log_path = os.path.join(
                log_dir,
                f"macro_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            )

        while attempts <= (max_retries if retry_on_error else 0):
            result = self.run_macro_monitored(
                macro_name, timeout_seconds, log_file=log_path
            )

            if result["status"] == "success":
                break

            attempts += 1
            if attempts <= max_retries:
                time.sleep(2)  # 等待 2 秒後重試

        # 執行後保存
        if auto_save and result["status"] == "success":
            try:
                self.save()
            except:
                pass

        # 添加備份信息
        result["backup_path"] = backup_path
        result["attempts"] = attempts

        return result

    def get_generated_files(self) -> List[str]:
        """
        獲取宏執行後生成的文件列表（基於快照對比）

        Returns:
            生成的文件路徑列表
        """
        current = self._take_snapshot()
        return sorted(current - self._file_snapshot)

    def close(self):
        """關閉文件"""
        try:
            if self.use_win32com:
                if self.workbook:
                    self.workbook.Close(SaveChanges=False)
                if self.excel:
                    self.excel.Quit()
            else:
                if self.vbaparser:
                    self.vbaparser.close()
        except:
            pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def __del__(self):
        self.close()


# 命令行接口
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="VBA 宏代碼讀取器")
    parser.add_argument("file", help=".xlsm 文件路徑")
    parser.add_argument("--module", "-m", help="指定模塊名稱")
    parser.add_argument("--list", "-l", action="store_true", help="列出所有模塊")
    parser.add_argument("--procs", "-p", action="store_true", help="列出所有過程")
    parser.add_argument("--analyze", "-a", action="store_true", help="分析代碼質量")
    parser.add_argument("--read-only", action="store_true", help="使用只讀模式（不需要 Excel）")
    parser.add_argument("--run", "-r", help="執行宏名稱")
    parser.add_argument("--params", nargs="*", help="宏參數列表")

    args = parser.parse_args()

    use_win32com = not args.read_only

    try:
        with VBAReader(args.file, use_win32com=use_win32com) as reader:
            if args.run:
                # 執行宏
                print(f"=== 執行宏: {args.run} ===")
                
                if args.params:
                    result = reader.run_macro_with_params(args.run, args.params)
                else:
                    result = reader.run_macro(args.run)
                
                print(f"\n執行結果:")
                print(f"  狀態: {result['status']}")
                print(f"  耗時: {result['duration']} 秒")
                
                if result['status'] == 'error':
                    print(f"  錯誤: {result.get('error', 'Unknown error')}")
                else:
                    print(f"  消息: {result.get('message', 'Success')}")
                    
                    # 獲取生成的文件
                    generated = reader.get_generated_files()
                    if generated:
                        print(f"\n生成的文件:")
                        for f in generated:
                            print(f"  - {f}")
            
            elif args.list:
                print("=== 模塊列表 ===")
                for module in reader.list_modules():
                    print(f"  - {module}")

            elif args.procs:
                print("=== 過程列表 ===")
                for proc in reader.list_procedures():
                    print(f"  {proc['type']}: {proc['module']}.{proc['name']}({', '.join(proc['params'])})")

            elif args.analyze:
                module_name = args.module or reader.list_modules()[0]
                print(f"=== 分析: {module_name} ===")
                analysis = reader.analyze_code(module_name)
                print(json.dumps(analysis, indent=2, ensure_ascii=False))

            elif args.module:
                code = reader.get_module(args.module)
                if code:
                    print(f"=== {args.module} ===")
                    print(code)
                else:
                    print(f"模塊 {args.module} 不存在")

            else:
                print("=== 所有模塊 ===")
                for name, code in reader.get_all_modules().items():
                    print(f"\n{'='*50}")
                    print(f"模塊: {name}")
                    print(f"{'='*50}")
                    print(code)

    except Exception as e:
        print(f"錯誤: {e}")
        exit(1)
