
# Win_Automation - Windows GUI 自動化專案

## 📋 專案簡介

Win_Automation 是一套基於 Python 與 pywinauto 的 Windows 桌面應用自動化框架，專為「低程式碼」自動化設計，所有操作流程皆由 JSON 配置檔驅動，無需修改程式即可快速調整自動化步驟。

主要特色：
- 支援多應用、多控件自動化（如 GeoTag Pro、檔案對話框、瀏覽器等）
- 完整 logging，執行進度與錯誤即時顯示於終端機
- 動態讀取 JSON 配置，流程調整零程式碼
- 內建多層降級控件搜尋，提升穩定性
- 支援多種 UI 操作（點擊、輸入、雙擊、右鍵等）

---

## 📁 專案結構

```
Win_Automation/
├── run.py                 # 主自動化腳本（讀取 JSON 配置並執行）
├── inspect_tool.py        # UI 元素偵測工具（協助配置 JSON）
├── requirements.txt       # Python 依賴
├── CLAUDE.md              # 專案說明文件
├── README.md              # 專案簡介與快速說明
├── doc/
│   ├── prompts.md         # 需求說明與設計理念
│   └── inspect-spec.md    # inspect_tool.py 使用說明
├── settings/
│   ├── case_01.json       # 自動化流程範例設定檔 1
│   ├── case_02.json       # 自動化流程範例設定檔 2
│   └── case_03.json       # 自動化流程範例設定檔 3
└── img/                   # 專案截圖與示意圖
```

### 主要檔案說明

- `run.py`：主程式，負責讀取 JSON 配置、連接應用、尋找控件、執行操作、錯誤處理與 logging。
- `inspect_tool.py`：Tkinter UI 偵測工具，可即時偵測滑鼠下方的 UI 元素資訊，並協助快速生成配置檔案。
- `settings/*.json`：自動化流程配置檔案，定義每一步的應用、控件、操作、延遲等。
- `doc/`：相關文件與使用說明。

---

## 🚀 快速開始

### 1. 安裝環境

```bash
# 建立虛擬環境（建議）
python -m venv venv
venv\Scripts\activate  # Windows

# 安裝依賴（包含 pywinauto 及其他工具）
pip install -r requirements.txt
```

### 2. 準備 UI 偵測工具（可選）

若需快速生成配置檔案，可使用 `inspect_tool.py` 來偵測 UI 元素：

```bash
python inspect_tool.py
```

- 點選「開始偵測」，移動滑鼠到任意 UI 元件即時顯示資訊
- 按 `Ctrl+F1` 記錄元件，選擇動作並填入參數
- 點選「結束偵測」匯出 JSON 配置

### 3. 編輯自動化流程

以 `settings/case_01.json` 為例，每個步驟格式如下：

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
python run.py --case settings/case_01.json
```

或者使用其他配置檔：

```bash
python run.py --case settings/case_02.json
python run.py --case settings/case_03.json
```

> **提示**：如果配置檔放在專案根目錄（如 `exec.json`），可直接執行 `python run.py --case exec.json` 或 `python run.py`（預設值為 exec.json）

---

## 🛠️ 核心設計與最佳實踐

### pywinauto 操作範例

```python
from pywinauto import Application
app = Application(backend="uia").connect(title="應用標題")
dlg = app.window(title="應用標題")
dlg.child_window(title="控件名稱", control_type="Button").click_input()
```

### 應用快取機制

為提升效能，`run.py` 實現了應用快取機制：
- 首次連接應用時，將其緩存於記憶體中
- 後續步驟使用同一應用時，直接複用已連接的實例，避免重複 `connect()`
- 大幅提升執行效率，尤其在多步驟流程中

### 控件搜尋策略

1. 先以 title + control_type 精確搜尋
2. 只用 title（忽略類型）
3. title 正則模糊比對（容忍空白/特殊字元差異）
4. 失敗時自動列出所有可見控件協助除錯

### 支援的操作類型

- click_input()：左鍵物理點擊（需元素在螢幕可視區域）
- click()：左鍵程式化點擊（使用 UIA Invoke Pattern，不移動滑鼠）
- double_click_input()：雙擊
- right_click_input()：右鍵
- send_keys() / type_keys()：鍵盤輸入
- set_text()：直接設值（輸入框）

> click() 使用 UIA Invoke Pattern 程式化觸發，適用於網頁或捲動區域內的控件；click_input() 使用物理滑鼠點擊，需元素在螢幕可視區域

### 動態配置與 logging

- 所有流程皆由 JSON 配置檔驅動，無需修改程式
- 執行進度、錯誤、控件搜尋策略皆即時顯示於終端機

---

## ⏱️ 延遲與錯誤處理

- 每步可自訂 Wait 秒數，確保 UI 響應
- 控件搜尋失敗時自動降級多策略，並列印所有可見控件協助排查
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

## 📝 進階應用與擴充

1. 新增自動化流程：複製 case_01.json，依需求調整步驟
2. 支援多應用、多視窗、多控件複雜流程
3. 可擴充更多 Action（如拖曳、滾動、複合操作）
4. 參考 doc/prompts.md 設計理念，保持配置驅動、低程式碼

---

## 🔗 參考資源

- [pywinauto 官方文件](https://pywinauto.readthedocs.io/)
- [Windows UI Automation](https://learn.microsoft.com/windows/win32/winauto/uiauto-intro)
- [UIA 控件類型列表](https://learn.microsoft.com/windows/win32/winauto/uiauto-controltypesoverview)

---

## 📌 備註

- 建議於 Windows 10/11、Python 3.7+ 環境執行
- 某些應用需管理員權限
- 建議保持目標應用於前景視窗，便於觀察自動化過程
- 所有流程皆可由 JSON 配置檔快速調整，無需修改程式

---

## 🤝 貢献與維護

1. 測試新自動化流程，複製/調整 JSON 配置
2. 如需擴充新操作，於 run.py 增加對應 Action 處理
3. 更新本說明文件，確保內容與專案同步
