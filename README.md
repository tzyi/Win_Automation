# Win_Automation

> 🤖 Windows 桌面應用程序 GUI 自動化工具

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## 📖 簡介

**Win_Automation** 是一個強大的 Python 自動化框架，用於自動化操作 Windows 桌面應用程序的圖形界面。特別適合重複性任務自動化，如批量文件處理、數據輸入、工作流自動化等。

本項目利用 Windows 原生的 **UI Automation (UIA)** 框架和 **pywinauto** 庫實現高效可靠的 UI 交互。

## ✨ 主要特性

- ✅ **簡單易用** - 基於配置文件的 JSON 驅動方式，無需複雜編程
- ✅ **跨應用支持** - 可操作任何使用 UIA 框架的 Windows 應用
- ✅ **靈活配置** - 通過 `exec.json` 管理所有 UI 元件信息
- ✅ **可擴展操作** - 支持點擊、輸入、拖拽等多種交互方式
- ✅ **智能延遲** - 自動處理應用響應時間

## 🚀 快速開始

### 前置要求

- Python 3.7 或更高版本
- Windows 10 / Windows 11
- 管理員權限（部分應用需要）

### 安裝

1. **克隆此倉庫**
   ```bash
   git clone https://github.com/yourusername/Win_Automation.git
   cd Win_Automation
   ```

2. **安裝依賴**
   ```bash
   pip install -r requirements.txt
   ```
   
   或直接安裝 pywinauto：
   ```bash
   pip install pywinauto
   ```

### 基本使用

#### 1️⃣ 探測 UI 元件

使用 `doc/test.py` 掃描目標應用的 UI 控件：

```python
from pywinauto import Application

app = Application(backend="uia").connect(title="你的應用標題")
dlg = app.window(title="你的應用標題")

# 列出所有控件信息
for ctrl in dlg.descendants():
    try:
        print(f"標題: {ctrl.window_text()}")
        print(f"控件類型: {ctrl.friendly_class_name()}")
        print(f"自動化 ID: {getattr(ctrl, 'automation_id', None)}")
        print("-" * 40)
    except Exception as e:
        print(f"錯誤: {e}")
```

#### 2️⃣ 配置 UI 元件

將要操作的 UI 元件信息添加到 `exec.json`：

```json
{
    "1": {
        "app": "應用標題",
        "Name": "控件標題",
        "ControlType": "Button"
    },
    "2": {
        "app": "應用標題",
        "Name": "另一個控件",
        "ControlType": "ListItem"
    }
}
```

#### 3️⃣ 執行自動化

運行生成的自動化腳本：

```bash
python run.py
```

### 完整範例

以下是一個實際的自動化流程示例（基於 GeoTag Pro 照片管理應用）：

```python
import time
from pywinauto import Application

# 步驟 1: 連接到 GeoTag Pro 並點擊「選擇照片」按鈕
app = Application(backend="uia").connect(title="GeoTag Pro")
dlg = app.window(title="GeoTag Pro")
dlg.child_window(title="folder_open 選擇照片", control_type="Text").click()

time.sleep(1)  # 等待「開啟」對話框

# 步驟 2: 在文件打開對話框中選擇文件
open_dlg = app.window(title="開啟")
open_dlg.child_window(title="IMG_4301.JPG", control_type="ListItem").click()

time.sleep(0.5)

# 步驟 3: 點擊「開啟」按鈕確認
open_dlg.child_window(title="開啟(O)", control_type="SplitButton").click()
```

## 📁 項目結構

```
Win_Automation/
├── README.md              # 本文件 - 快速參考指南
├── CLAUDE.md              # 詳細技術文檔
├── run.py                 # 自動生成的執行腳本
├── exec.json              # UI 元件配置文件
├── requirements.txt       # Python 依賴列表
└── doc/
    ├── prompts.md         # 自動化需求說明
    └── test.py            # UI 元件探測工具
```

### 文件說明

| 文件 | 用途 |
|------|------|
| `run.py` | 根據 `exec.json` 配置生成的自動化執行腳本 |
| `exec.json` | 定義要操作的 UI 元件（應用標題、控件名稱、控件類型） |
| `doc/test.py` | 用於掃描和列出應用中所有可用控件的調試工具 |
| `doc/prompts.md` | 記錄自動化任務需求和工作流說明 |

## 🎮 支持的操作

### 當前支持

- **左鍵點擊** - 激活按鈕、選擇項目等

### 可擴展操作

- 雙擊
- 右鍵點擊  
- 拖拽和放置
- 文本輸入
- 滾動
- 快捷鍵

## 🔧 高級用法

### 常見控件類型

```
Button          - 按鈕
Text            - 文本框
ListItem        - 列表項
SplitButton     - 分割按鈕
Window          - 窗口
ComboBox        - 組合框
CheckBox        - 複選框
Edit            - 編輯框
```

### 時間延遲管理

合理控制延遲時間以確保 UI 元件完全加載：

```python
import time

# 短延遲 (0.5s) - 快速響應的 UI 元素
time.sleep(0.5)

# 中延遲 (1s) - 對話框或窗口打開
time.sleep(1)

# 長延遲 (2s+) - 應用加載或複雜操作
time.sleep(2)
```

### 查找控件的多種方式

```python
# 按標題查找
ctrl = dlg.child_window(title="按鈕標題", control_type="Button")

# 按自動化 ID 查找
ctrl = dlg.child_window(automation_id="btnOK", control_type="Button")

# 使用正則表達式
import re
ctrl = dlg.child_window(title_re=r"確定.*", control_type="Button")

# 查找所有子元素
for ctrl in dlg.descendants():
    print(ctrl.window_text())
```

## 🐛 常見問題排查

| 問題 | 可能原因 | 解決方案 |
|------|--------|--------|
| 無法連接到應用 | 應用未運行或標題不匹配 | 檢查應用標題，可用 `test.py` 確認 |
| FindingElement timeout | 控件不存在或隱藏 | 增加延遲時間或確認控件屬性 |
| Access Denied | 權限不足 | 以管理員身份運行 Python 或 IDE |
| 操作無響應 | 應用未響應或被遮擋 | 保持應用在前景，檢查應用狀態 |
| 控件未找到 | 控件名稱或類型錯誤 | 使用 `test.py` 列出正確的屬性 |

## 📚 詳細文檔

- **[CLAUDE.md](CLAUDE.md)** - 完整技術文檔、最佳實踐和深度教程
- **[官方 pywinauto 文檔](https://pywinauto.readthedocs.io/)**
- **[Windows UI Automation](https://docs.microsoft.com/windows/win32/winauto/uiauto-intro)**

## 💡 使用範例

### 範例 1：批量文件處理

```python
import time
from pywinauto import Application

app = Application(backend="uia").connect(title="應用名稱")
dlg = app.window(title="應用名稱")

# 循環處理多個文件
files = ["文件1.txt", "文件2.txt", "文件3.txt"]
for file in files:
    dlg.child_window(title="開啟", control_type="Button").click()
    time.sleep(1)
    
    # 選擇文件
    dlg.child_window(title=file, control_type="ListItem").click()
    dlg.child_window(title="確定", control_type="Button").click()
    
    time.sleep(0.5)
```

### 範例 2：數據輸入工作流

```python
import time
from pywinauto import Application

app = Application(backend="uia").connect(title="表單應用")
dlg = app.window(title="表單應用")

# 輸入數據
dlg.child_window(control_type="Edit", top_level_only=False).type_keys("輸入內容")
time.sleep(0.3)

# 提交表單
dlg.child_window(title="提交", control_type="Button").click()
```

## 🎯 最佳實踐

1. **始終使用管理員權限** - 某些應用需要提升權限
2. **合理設置延遲** - 避免過短導致超時或過長導致緩慢
3. **錯誤處理** - 添加 try-catch 捕獲異常
4. **使用相對定位** - 優先使用標題而非座標
5. **保持應用前景** - 調試時將應用窗口置於前景
6. **記錄操作日誌** - 便於問題排查和維護

## 🤝 貢獻指南

歡迎提交改進建議！請遵循以下流程：

1. Fork 本倉庫
2. 創建特性分支 (`git checkout -b feature/新功能`)
3. 提交更改 (`git commit -am '添加新功能'`)
4. 推送到分支 (`git push origin feature/新功能`)
5. 提交 Pull Request

## 📝 授權

本項目採用 MIT 授權證。詳見 [LICENSE](LICENSE) 文件。

## 📞 支持

- 📖 查看 [CLAUDE.md](CLAUDE.md) 了解詳細技術文檔
- 🐛 提交 Issue 報告問題
- 💬 討論功能需求和改進建議

## 📌 注意事項

- ⚠️ 所有操作使用 UIA 後端，兼容現代 Windows 應用
- ⚠️ 某些應用可能需要特殊權限或特定配置
- ⚠️ 在生產環境使用前充分測試
- ⚠️ 仍在活躍開發中，API 可能變動

---

**Made with ❤️ for Windows automation enthusiasts**

