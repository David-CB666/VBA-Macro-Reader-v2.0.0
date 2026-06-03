# VBA Macro Reader & Operator Skill

> 版本：v2.0.0
> 創建時間：2026-06-03
> 更新時間：2026-06-03 12:25
> 用途：讀取、修改、**執行** VBA 宏代碼，生成文件和處理結果

---

## 📋 功能概述

本技能提供完整的 VBA 宏代碼讀取、分析、操作和**執行**能力：

| 功能 | 描述 | 狀態 |
|:---|:---|:---:|
| **讀取 VBA 代碼** | 從 .xlsm/.xlam 提取所有模塊代碼 | ✅ |
| **解析模塊結構** | 識別 Sub/Function、參數、變量 | ✅ |
| **提取參數信息** | 分析宏的輸入參數和返回值 | ✅ |
| **實時修改** | 通過 win32com 動態修改宏代碼 | ✅ |
| **★執行宏代碼** | **直接運行宏，生成文件和結果** | ✅ |
| **★監控執行** | **監控執行過程，捕獲錯誤** | ✅ |
| **★結果處理** | **處理生成的文件和輸出** | ✅ |
| **優化建議** | 基於最佳實踐提供優化建議 | ✅ |
| **代碼對比** | 對比不同版本的 VBA 代碼差異 | ✅ |

---

## 🔧 技術方案

### 方案 A：win32com 方法（✅ 推薦）

**優點：**
- 可讀取和修改 VBA 代碼
- 可實時執行宏
- 支持 Excel 所有功能

**缺點：**
- 需要安裝 Excel
- Windows only

**代碼示例：**
```python
import win32com.client

# 打開 Excel 文件
excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(r"D:\path\to\file.xlsm")

# 讀取 VBA 代碼
for comp in wb.VBProject.VBComponents:
    if comp.Type in [1, 2, 3]:  # 標準模塊、類模塊、窗體
        code_module = comp.CodeModule
        line_count = code_module.CountOfLines
        code = code_module.Lines(1, line_count)
        print(f"=== {comp.Name} ===")
        print(code)

# 修改 VBA 代碼
comp.CodeModule.DeleteLines 1, line_count
comp.CodeModule.AddFromString("Sub Test()\n    MsgBox \"Hello\"\nEnd Sub")

wb.Save()
excel.Quit()
```

---

### 方案 B：oletools 解析方法

**優點：**
- 不需要 Excel
- 跨平台

**缺點：**
- 只能讀取，不能修改
- 需要手動解析二進制格式

**代碼示例：**
```python
from oletools.olevba import VBA_Parser

# 解析 VBA
vbaparser = VBA_Parser(r"D:\path\to\file.xlsm")

# 提取所有 VBA 代碼
for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
    print(f"=== {vba_filename} ===")
    print(vba_code)

vbaparser.close()
```

---

## 📋 使用方法

### 1. 讀取 VBA 宏代碼

```python
from scripts.vba_reader import VBAReader

# 初始化
reader = VBAReader(r"D:\path\to\file.xlsm")

# 讀取所有模塊
modules = reader.get_all_modules()

for module_name, code in modules.items():
    print(f"\n=== {module_name} ===")
    print(code)

# 讀取特定模塊
main_code = reader.get_module("mod_Main")

# 列出所有 Sub/Function
procs = reader.list_procedures()
for proc in procs:
    print(f"{proc['type']}: {proc['name']}({proc['params']})")
```

---

### ★2. 執行宏代碼（核心功能）

```python
from scripts.vba_reader import VBAReader

# 初始化（必須使用 win32com 模式）
reader = VBAReader(r"D:\path\to\file.xlsm", use_win32com=True)

# 方法 1：直接執行宏（無參數）
result = reader.run_macro("FillExcelTemplate")
print(f"執行結果: {result}")

# 方法 2：執行宏並傳遞參數
result = reader.run_macro_with_params(
    "ExportWorksheetToPDF",
    [worksheet_object, "報告_20260603"]
)

# 方法 3：執行宏並監控執行過程
result = reader.run_macro_monitored(
    "FillExcelTemplate",
    timeout_seconds=300,  # 5 分鐘超時
    log_file="D:\logs\macro_run.log"
)

# 方法 4：批量執行多個宏
macros_to_run = [
    ("PrepareData", []),
    ("FillExcelTemplate", []),
    ("ExportToPDF", []),
]
results = reader.run_macros_batch(macros_to_run)

# 關閉
reader.close()
```

---

### ★3. 處理宏執行結果

```python
# 執行宏並獲取生成的文件
result = reader.run_macro("FillExcelTemplate")

# 檢查生成的文件
generated_files = reader.get_generated_files()
print(f"生成的文件: {generated_files}")

# 檢查執行日誌
logs = reader.get_execution_logs()
print(f"執行日誌:\n{logs}")

# 檢查錯誤信息
if result['status'] == 'error':
    print(f"錯誤信息: {result['error']}")
    print(f"錯誤行號: {result['line_number']}")
```

---

### ★4. 智能執行（自動選擇最佳方式）

```python
# 智能執行宏（自動處理理參數、監控、錯誤處理）
result = reader.smart_run(
    macro_name="FillExcelTemplate",
    auto_save=True,        # 執行後自動保存
    auto_backup=True,      # 執行前自動備份
    timeout_seconds=300,   # 超時設置
    retry_on_error=True,   # 錯誤時自動重試
    max_retries=3,         # 最大重試次數
    log_to_file=True,      # 記錄到日誌文件
)

# 輸出執行報告
print(f"執行狀態: {result['status']}")
print(f"執行時間: {result['duration']}秒")
print(f"生成的文件: {result['generated_files']}")
print(f"處理的數據行數: {result['rows_processed']}")
```

---

### 2. 提取參數信息

```python
# 提取宏的參數
params = reader.extract_parameters("mod_Main", "FillTemplate")

print(f"參數列表: {params}")
# 輸出: ['templatePath', 'dataSource', 'outputPath', 'formatType']
```

---

### 3. 實時修改宏代碼

```python
# 修改宏代碼
reader.update_module("mod_Main", """
Sub FillTemplate(templatePath As String, dataSource As String)
    ' 優化後的代碼
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")

    ' 使用字典提高性能
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' ... 新邏輯 ...
End Sub
""")

# 保存
reader.save()
```

---

### 4. 優化建議

```python
# 分析代碼質量
analysis = reader.analyze_code("mod_Main")

print(f"代碼行數: {analysis['line_count']}")
print(f"複雜度: {analysis['complexity']}")
print(f"建議: {analysis['suggestions']}")

# 自動優化
optimized_code = reader.optimize_module("mod_Main")
reader.update_module("mod_Main", optimized_code)
reader.save()
```

---

## 📋 核心 API

### VBAReader 類

```python
class VBAReader:
    def __init__(self, filepath: str, use_win32com: bool = True):
        """
        初始化 VBA 讀取器

        Args:
            filepath: .xlsm 文件路徑
            use_win32com: True 使用 win32com（可讀可寫）
                         False 使用 oletools（只讀）
        """

    def get_all_modules(self) -> Dict[str, str]:
        """返回所有 VBA 模塊代碼"""

    def get_module(self, module_name: str) -> str:
        """返回指定模塊代碼"""

    def list_modules(self) -> List[str]:
        """列出所有模塊名稱"""

    def list_procedures(self) -> List[Dict]:
        """列出所有 Sub/Function 及參數"""

    def extract_parameters(self, module: str, procedure: str) -> List[str]:
        """提取指定過程的參數列表"""

    def update_module(self, module_name: str, code: str) -> bool:
        """更新模塊代碼（僅 win32com）"""

    def add_module(self, module_name: str, code: str) -> bool:
        """新增模塊（僅 win32com）"""

    def delete_module(self, module_name: str) -> bool:
        """刪除模塊（僅 win32com）"""

    def analyze_code(self, module_name: str) -> Dict:
        """分析代碼質量"""

    def optimize_module(self, module_name: str) -> str:
        """返回優化後的代碼"""

    def compare_versions(self, old_code: str, new_code: str) -> Dict:
        """對比兩個版本的代碼差異"""

    def save(self) -> bool:
        """保存修改（僅 win32com）"""

    def close(self):
        """關閉文件"""
```

---

## 📋 應用場景

### 場景 1：工程總表宏系統優化

**問題：** VBA 代碼性能較慢，需要優化

**解決方案：**
```python
reader = VBAReader(r"D:\Desktop\工程總表宏\工程總表格宏.xlsm")

# 分析所有模塊
for module in reader.list_modules():
    analysis = reader.analyze_code(module)

    if analysis['complexity'] > 10:
        print(f"{module} 複雜度過高: {analysis['complexity']}")
        print(f"建議: {analysis['suggestions']}")
```

---

### 場景 2：批量更新宏參數

**問題：** 多個 .xlsm 文件的宏參數需要統一修改

**解決方案：**
```python
import os

files = [
    r"D:\Projects\ProjectA\報批表.xlsm",
    r"D:\Projects\ProjectB\報批表.xlsm",
]

new_code = """
Sub FillTemplate(templatePath As String, dataSource As String, Optional formatType As String = "PDF")
    ' 新增 formatType 參數
    ' ...
End Sub
"""

for file in files:
    reader = VBAReader(file)
    reader.update_module("mod_Main", new_code)
    reader.save()
    reader.close()
    print(f"已更新: {file}")
```

---

### 場景 3：VBA 代碼導出備份

**問題：** 將 VBA 宏備份為獨立文件

**解決方案：**
```python
reader = VBAReader(r"D:\工程總表宏.xlsm")

# 提取所有 VBA 代碼
modules = reader.get_all_modules()

# 保存為 .bas 文件
for name, code in modules.items():
    with open(f"{name}.bas", "w", encoding="utf-8") as f:
        f.write(code)
```

---

## 📋 已知限制

| 限制 | 說明 | 解決方案 |
|:---|:---|:---|
| 需要 Excel | win32com 方法需要安裝 Excel | 使用 oletools 只讀模式 |
| Windows only | win32com 僅支持 Windows | 使用 oletools 跨平台 |
| 密碼保護 | 密碼保護的 VBA 無法讀取 | 需要密碼才能解鎖 |
| 大文件 | 超大 .xlsm 文件讀取較慢 | 分批讀取模塊 |

---

## 📋 依賴項

```
pywin32>=310
oletools>=0.60
olefile>=0.47
```

---

## 📋 文件結構

```
skills/vba-reader/
├── SKILL.md                  # 本文件
├── README.md                 # 用戶文檔
├── package.json              # 包元數據
├── scripts/
│   └── vba_reader.py        # 核心讀取器（包含所有功能）
├── examples/
│   ├── read_vba.py          # 讀取示例
│   └── run_macro.py         # 執行示例
└── tests/
    └── test_reader.py       # 單元測試
```

---

_版本：v2.0.0 | 創建時間：2026-06-03_