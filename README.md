# VBA Macro Reader & Operator Skill

> 版本：v2.0.0
> 最後更新：2026-06-03
> 作者：[姓名]

---

## 📋 功能概述

完整的 VBA 宏代碼讀取、修改、**執行**解決方案：

| 功能 | 描述 |
|:---|:---|
| **讀取 VBA 代碼** | 從 .xlsm/.xlam 提取所有模塊代碼 |
| **解析模塊結構** | 識別 Sub/Function、參數、變量 |
| **實時修改** | 通過 win32com 動態修改宏代碼 |
| **★執行宏代碼** | **直接運行宏，生成文件和結果** |
| **★監控執行** | **超時控制、日誌記錄、錯誤捕獲** |
| **★結果處理** | **獲取生成的文件列表** |

---

## 🔧 安裝

### 依賴項

```bash
pip install pywin32 oletools
```

---

## 📋 使用方法

### 1. 讀取 VBA 代碼

```python
from scripts.vba_reader import VBAReader

reader = VBAReader("file.xlsm")
modules = reader.get_all_modules()

for name, code in modules.items():
    print(f"=== {name} ===")
    print(code)
```

---

### 2. 執行宏代碼

```python
reader = VBAReader("file.xlsm", use_win32com=True)

# 執行宏
result = reader.run_macro("FillExcelTemplate")

print(f"狀態: {result['status']}")
print(f"耗時: {result['duration']} 秒")

# 獲取生成的文件
files = reader.get_generated_files()
```

---

### 3. 智能執行

```python
result = reader.smart_run(
    macro_name="FillExcelTemplate",
    auto_save=True,
    auto_backup=True,
    timeout_seconds=300,
    retry_on_error=True,
    max_retries=3,
)
```

---

### 4. 命令行使用

```bash
# 列出模塊
python scripts/vba_reader.py file.xlsm --list

# 列出過程
python scripts/vba_reader.py file.xlsm --procs

# 執行宏
python scripts/vba_reader.py file.xlsm --run FillExcelTemplate
```

---

## 📋 核心 API

```python
class VBAReader:
    # 讀取功能
    def get_all_modules(self) -> Dict[str, str]
    def get_module(self, module_name: str) -> str
    def list_modules(self) -> List[str]
    def list_procedures(self) -> List[Dict]

    # 修改功能
    def update_module(self, module_name: str, code: str) -> bool
    def add_module(self, module_name: str, code: str) -> bool
    def delete_module(self, module_name: str) -> bool

    # ★執行功能（核心）
    def run_macro(self, macro_name: str) -> Dict
    def run_macro_with_params(self, macro_name: str, params: List) -> Dict
    def run_macro_monitored(self, macro_name: str, timeout_seconds: int = 300) -> Dict
    def run_macros_batch(self, macros: List[Tuple[str, List]]) -> List[Dict]
    def smart_run(self, macro_name: str, **kwargs) -> Dict
    def get_generated_files(self) -> List[str]
```

---

## 📋 技術方案

### win32com 模式（推薦）

- ✅ 可讀取和修改 VBA 代碼
- ✅ 可實時執行宏
- ✅ 支持 Excel 所有功能
- ⚠️ 需要安裝 Excel
- ⚠️ Windows only

### oletools 模式

- ✅ 不需要 Excel
- ✅ 跨平台
- ❌ 只能讀取，不能修改

---

## 📋 已知限制

| 限制 | 說明 | 解決方案 |
|:---|:---|:---|
| 需要 Excel | win32com 需要安裝 Excel | 使用 oletools 只讀模式 |
| Windows only | win32com 僅支持 Windows | 使用 oletools 跨平台 |
| 密碼保護 | 密碼保護的 VBA 無法讀取 | 需要密碼才能解鎖 |

---

## 📋 文件結構

```
vba-reader/
├── SKILL.md              # 技能文檔（10.3 KB）
├── README.md             # 本文件
├── scripts/
│   └── vba_reader.py     # 核心模塊（23 KB）
├── examples/
│   ├── read_vba.py       # 讀取示例（5.1 KB）
│   └── run_macro.py      # 執行示例（5.8 KB）
└── tests/
    └── test_reader.py    # 測試文件（2.9 KB）
```

---

## 📋 許可證

MIT License

---

## 📋 更新日誌

### v2.0.0 (2026-06-03)

- ✨ 新增執行宏功能（核心功能）
- ✨ 新增監控執行功能
- ✨ 新增結果處理功能
- ✨ 新增智能執行功能
- ✨ 新增批量執行功能
- 🔒 完成脫敏處理

### v1.0.0 (2026-06-03)

- ✅ 初始版本
- ✅ 讀取 VBA 代碼功能
- ✅ 解析模塊結構
- ✅ 實時修改功能

---

**VBA Macro Reader v2.0.0 - 讀取、修改、執行 VBA 宏代碼！** 🎉