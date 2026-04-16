
# Win_Automation

> 🤖 Windows 桌面應用 GUI 自動化框架

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## 📖 專案簡介

**Win_Automation** 是一套基於 Python 與 pywinauto 的 Windows 桌面應用自動化框架，專為「低程式碼」自動化設計，所有操作流程皆由 JSON 配置檔驅動，無需修改程式即可快速調整自動化步驟。

主要特色：
- 支援多應用、多控件自動化（如 GeoTag Pro、檔案對話框、瀏覽器等）
- 完整 logging，執行進度與錯誤即時顯示於終端機
- 動態讀取 JSON 配置，流程調整零程式碼
- 內建多層降級控件搜尋，提升穩定性
- 支援多種 UI 操作（點擊、輸入、雙擊、右鍵等）

---

## 🚀 快速開始

### 1. 安裝環境


#### 安裝套件
```bash
pip install pywinauto
```

#### 安裝Inspect.exe (Windows SDK for Windows 11)
- [下載點](https://learn.microsoft.com/zh-tw/windows/apps/windows-sdk/downloads)
- 只安裝「Windows SDK for Windows 11」中的「Windows SDK for Desktop C++ Apps」部分


### 2. 編輯自動化流程

以 `case_01.json` 為例，每個步驟格式如下：

```json
{
  "1": {
    "app": "GeoTag Pro",
    "Name": "folder_open 選擇照片",
    "ControlType": "UIA_TextControlTypeId (0xC364)",
    "Action": "click_input()",
    "value": ""
  },
  ...
}
```

欄位說明：
- `app`：目標應用視窗標題
- `Name`：控件名稱（顯示文字）
- `ControlType`：控件類型（UIA_...ControlTypeId 或 pywinauto 名稱）
- `Action`：執行動作（如 click_input()、send_keys()、set_text() 等）
- `value`：輸入內容（如需）
- `Wait`：本步驟結束後等待秒數（可選）

### 3. 執行自動化

```bash
python run.py --case case_01.json
```

---

## 📁 專案結構

```
Win_Automation/
├── run.py              # 主自動化腳本（讀取 JSON 配置並執行）
├── case_01.json        # 自動化流程範例設定檔（可自訂多組）
├── CLAUDE.md           # 詳細技術文檔
├── README.md           # 快速參考指南
└── doc/
    └── prompts.md      # 需求說明與設計理念
```

---

## 🛠️ 核心設計理念

- 所有流程皆由 JSON 配置檔驅動，無需修改程式
- 控件搜尋採多層降級策略（title+type、title、正則模糊），失敗時自動列印所有可見控件協助排查
- 支援多種 UI 操作（click_input、send_keys、set_text 等）
- 完整 logging，執行進度、錯誤、控件搜尋策略皆即時顯示於終端機
- 每步可自訂 Wait 秒數，確保 UI 響應
- 執行失敗時自動停止流程，並顯示錯誤

---

## 🔍 常見問題排查

| 問題 | 可能原因 | 解決方案 |
|------|--------|--------|
| 連接不到應用 | 標題不符或應用未啟動 | 檢查 app 名稱、確保應用已開啟 |
| 找不到控件 | 名稱/類型錯誤或 UI 動態變化 | 用降級搜尋、列印控件清單協助排查 |
| 操作無反應 | 控件未顯示或延遲不足 | 增加 Wait 秒數，確保 UI 完全載入 |
| 權限不足 | 需管理員權限 | 以管理員身份執行 Python |

---

## 📚 進階說明

- 詳細設計理念、pywinauto 操作範例、控件搜尋策略、最佳實踐請參閱 [CLAUDE.md](CLAUDE.md)
- 參考 [pywinauto 官方文件](https://pywinauto.readthedocs.io/)
- 參考 [Windows UI Automation](https://learn.microsoft.com/windows/win32/winauto/uiauto-intro)

---

## 🤝 貢獻與維護

1. 測試新自動化流程，複製/調整 JSON 配置
2. 如需擴充新操作，於 run.py 增加對應 Action 處理
3. 更新本說明文件，確保內容與專案同步

---

**Made with ❤️ for Windows automation enthusiasts**

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

|方法|功能|備註|
|-|-|-|
|**click()**|發送滑鼠點擊訊息|不要求控制項可見|
|**click\_input()**|使用滑鼠事件進行點擊|需要控制項在螢幕上可見，更貼近實際使用者操作|
|**double\_click()**|雙擊|基於訊息|
|**double\_click\_input()**|雙擊輸入|基於滑鼠事件|
|**right\_click()**|右鍵點擊|基於訊息|
|**right\_click\_input()**|右鍵點擊輸入|基於滑鼠事件|
|**type\_keys(keys)**|發送鍵盤輸入|例如：`type\_keys("{TAB}")`, `type\_keys("{ENTER}")`|
|**send\_message(msg)**|發送 Windows 訊息|進階用法|
|**send\_message\_timeout()**|發送訊息並等待回應|可設定逾時時間|
|**set\_focus()**|設定焦點|-|
|**get\_focus()**|取得焦點|-|
|**set\_window\_text(text)**|設定視窗文字|-|
|**close()**|關閉視窗|-|
|**close\_click()**|點擊關閉|執行額外延遲|
|**drag\_mouse(dx, dy)**|拖動滑鼠|-|
|**move\_mouse(x, y)**|移動滑鼠|-|
|**press\_mouse()**|按下滑鼠按鍵|-|
|**release\_mouse()**|放開滑鼠按鍵|-|
|**press\_mouse\_input()**|按下滑鼠（基於事件）|-|
|**release\_mouse\_input()**|放開滑鼠（基於事件）|-|
|**maximize()**|最大化|-|
|**minimize()**|最小化|-|
|**restore()**|還原|-|
|**get\_show\_state()**|取得顯示狀態|-|
|**menu\_select(path)**|選擇菜單項目|-|
|**notify\_menu\_select()**|通知菜單選擇|-|
|**notify\_parent()**|通知父視窗|-|
|**draw\_outline()**|繪製輪廓|用於視覺化|
|**move\_window(x, y, width, height)**|移動與調整大小|-|
|**rectangle**|取得矩形座標|屬性：.left, .top, .right, .bottom|
|**children**|取得子視窗列表|-|


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

