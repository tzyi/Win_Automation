# Win_Automation - Windows GUI 自動化項目

## 📋 項目概述

Win_Automation 是一個 Python 項目，用於自動化操作 Windows 應用程序的 GUI 界面，特別是針對 **GeoTag Pro** 應用的自動化流程。使用 **pywinauto** 庫與 Windows UI 自動化框架進行交互。

## 🎯 主要目的

- 自動化操作 Windows 桌面應用程序
- 通過 UI 自動化發送點擊、文本輸入等操作
- 減少重複性的手動操作
- 支持跨應用程序的自動化工作流

## 🛠️ 技術棧

- **Python**: 主要編程語言
- **pywinauto**: Windows UI 自動化庫
- **UIA (UI Automation)**: Windows 原生自動化框架

## 📁 項目結構

```
Win_Automation/
├── run.py              # 主要的自動化腳本（生成的執行代碼）
├── exec.json           # UI 元件配置文件，存儲要操作的控件信息
├── CLAUDE.md           # 本文件
└── doc/
    ├── prompts.md      # 用戶提示和需求說明
    └── test.py         # 測試腳本和示例代碼
```

### 文件詳細說明

#### `exec.json`
- 存儲項目中要自動化操作的 UI 元件信息
- 結構: 每個條目包含應用標題、控件名稱和控件類型
- 示例:
  ```json
  {
    "1": {
      "app": "GeoTag Pro",
      "Name": "folder_open 選擇照片",
      "ControlType": "UIA_TextControlTypeId (0xC364)"
    }
  }
  ```

#### `run.py`
- 自動生成的主執行腳本
- 根據 `exec.json` 中的配置，使用 pywinauto 執行自動化操作
- 典型操作流程:
  1. 連接到目標應用（如 GeoTag Pro）
  2. 獲取窗口和控件引用
  3. 執行操作（點擊、輸入等）
  4. 添加適當的延遲以等待 UI 響應

#### `doc/prompts.md`
- 存儲用戶的自動化需求和提示
- 可包含要執行的操作序列說明

#### `doc/test.py`
- 測試和調試腳本
- 包含發現 UI 元件屬性的示例代碼

## 🚀 快速開始

### 前置環境
```bash
pip install pywinauto
```

### 基本操作流程

1. **探測 UI 元件**
   - 使用 `test.py` 中的演示代碼掃描並打印應用中的控件信息
   - 取消註釋 for 循環以列出所有控件

2. **配置 `exec.json`**
   - 將需要操作的 UI 元件信息添加到配置文件
   - 包括應用標題、控件名稱和控件類型

3. **執行自動化**
   ```bash
   python run.py
   ```

## 💡 關鍵概念和最佳實踐

### pywinauto 基本用法

```python
from pywinauto import Application

# 連接到應用
app = Application(backend="uia").connect(title="應用標題")

# 獲取窗口
dlg = app.window(title="應用標題")

# 查找和操作控件
dlg.child_window(title="控件名稱", control_type="Button").click()
```

### 控件類型 (ControlType)

常見的 UIA 控件類型:
- `Button`: 按鈕
- `Text`: 文本框
- `ListItem`: 列表項
- `SplitButton`: 分割按鈕
- `Window`: 窗口
- `ComboBox`: 組合框

### 操作類型

當前實現的操作:
- **左鍵點擊** (click): 最常用的交互方式

可擴展的操作:
- 雙擊
- 右鍵點擊
- 拖拽
- 文本輸入
- 滾動

## ⏱️ 時間延遲處理

在自動化操作中，合理的延遲非常重要:
- 短延遲 (0.5s): 快速響應的 UI 元素
- 中延遲 (1s): 對話框或窗口打開等待
- 長延遲 (2s+): 應用程序加載或複雜操作

```python
import time
time.sleep(1)  # 等待 UI 響應
```

## 🔍 調試技巧

### 列出應用中的所有控件

```python
dlg = app.window(title="GeoTag Pro")
for ctrl in dlg.descendants():
    try:
        print(f"title: {ctrl.window_text()}")
        print(f"control_type: {ctrl.friendly_class_name()}")
        print(f"automation_id: {getattr(ctrl, 'automation_id', None)}")
    except Exception as e:
        print(f"Error: {e}")
```

### 常見問題排查

| 問題 | 可能原因 | 解決方案 |
|------|--------|--------|
| 無法連接到應用 | 應用未運行或標題不符 | 檢查 `Application.connect(title="...")` 中的標題 |
| 找不到控件 | 控件名稱或類型錯誤 | 使用上述列表方法確認正確的屬性 |
| 操作超時 | 延遲時間不足 | 增加 `time.sleep()` 的值 |
| 访问被拒 | 應用可能運行在不同的權限級別 | 以管理員身份運行 Python 腳本 |

## 📝 常見任務

### 添加新的自動化流程

1. 在 GeoTag Pro 中手動執行操作，記下每一步
2. 使用 `test.py` 掃描 UI 元件
3. 將元件信息添加到 `exec.json`
4. 在 `run.py` 中編寫相應的自動化代碼

### 修改現有流程

- 編輯 `exec.json` 中的控件信息
- 在 `run.py` 中調整操作順序和延遲
- 測試並驗證修改

## 🔗 相關資源

- [pywinauto 官方文檔](https://pywinauto.readthedocs.io/)
- [Windows UI Automation 文檔](https://docs.microsoft.com/windows/win32/winauto/uiauto-intro)
- [UIA 控件類型列表](https://docs.microsoft.com/windows/win32/winauto/uiauto-controltypesoverview)

## 📌 備註

- 所有操作都使用 UIA 後端，兼容大多數現代 Windows 應用
- 對話框或窗口打開時需要適當延遲
- 某些應用可能需要管理員權限才能進行 UI 自動化
- 在調試時，保持應用在前景窗口有助於觀察操作過程

## 🤝 貢獻指南

如要擴展此項目:
1. 測試新的自動化流程
2. 更新 `exec.json` 配置
3. 在 `run.py` 中實現新操作
4. 更新此文檔以反映新功能
