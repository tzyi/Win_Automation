"""
handler.py - Win_Automation 自訂 Handle 模組

本模組定義所有可在 JSON 配置中以 "handle" 欄位引用的自訂處理函數。
每個 handler 函數在執行實際 UI 操作前被呼叫，可對 value 進行預處理。

函數簽名規範：
    def handle_xxx(dlg, step: dict, value: str) -> str
        dlg   : 當前操作的 pywinauto 視窗物件
        step  : 目前步驟的完整配置 dict（來自 JSON）
        value : 原始 value 字串
        return: 處理後的 value 字串

在 JSON 配置中使用方式：
    {
        "handle": "check_file_name()"
    }
"""

import os
import re
import logging
from datetime import datetime

logger = logging.getLogger("WinAutomation")


# ============================================================
# 內部輔助工具
# ============================================================

def _build_known_folder_map() -> dict:
    """
    建立 Windows 本地化資料夾顯示名稱 → 實際路徑的對照表。
    同時支援中文（繁 / 簡）與英文顯示名稱。
    """
    home = os.path.expanduser("~")
    candidates = [
        # 繁中 / 簡中 / 英文顯示名稱 → 實際子目錄名
        ("下載",   "Downloads"),
        ("Downloads", "Downloads"),
        ("桌面",   "Desktop"),
        ("Desktop",  "Desktop"),
        ("文件",   "Documents"),
        ("我的文件", "Documents"),
        ("Documents", "Documents"),
        ("圖片",   "Pictures"),
        ("我的圖片", "Pictures"),
        ("Pictures", "Pictures"),
        ("音樂",   "Music"),
        ("我的音樂", "Music"),
        ("Music",   "Music"),
        ("影片",   "Videos"),
        ("我的影片", "Videos"),
        ("Videos",  "Videos"),
    ]
    mapping: dict = {}
    for display, subdir in candidates:
        real = os.path.join(home, subdir)
        if os.path.isdir(real):
            mapping[display] = real
    # 也包含 home 本身（以使用者名稱顯示）
    mapping[os.path.basename(home)] = home
    return mapping


# 模組載入時預先建立（只建一次）
_KNOWN_FOLDER_MAP: dict = _build_known_folder_map()


def _get_dialog_folder_path(dlg) -> str:
    """
    嘗試從檔案對話框（另存新檔 / 開啟）取得目前資料夾路徑。

    策略優先順序：
      1. 遍歷 ToolBar / SplitButton 控件，嘗試直接讀取完整路徑
      2. ToolBar / SplitButton 文字帶有 "位址:" / "Address:" 前綴
      3. ToolBar / SplitButton 文字匹配已知資料夾顯示名稱
      4. 遍歷所有控件，尋找文字是合法絕對路徑目錄者
      5. Fallback 到常見預設資料夾（Downloads 優先）
    """
    # 收集所有 ToolBar / SplitButton 文字，後面策略共用
    toolbar_texts: list = []
    try:
        for child in dlg.descendants():
            try:
                cls = child.friendly_class_name()
                if cls in ("ToolBar", "SplitButton"):
                    txt = child.window_text().strip()
                    if txt:
                        toolbar_texts.append(txt)
            except Exception:
                continue
    except Exception:
        pass

    # --- 策略 1 & 2：ToolBar / SplitButton 包含完整路徑 ---
    for txt in toolbar_texts:
        # 直接就是合法目錄路徑
        if os.path.isabs(txt) and os.path.isdir(txt):
            return txt
        # 帶有 "位址: PATH" 或 "Address: PATH" 前綴
        m = re.search(r'(?:位址|Address)\s*:\s*(.+)', txt)
        if m:
            p = m.group(1).strip()
            if os.path.isdir(p):
                return p
            # prefix 之後是顯示名稱
            if p in _KNOWN_FOLDER_MAP:
                return _KNOWN_FOLDER_MAP[p]

    # --- 策略 3：ToolBar / SplitButton 文字匹配已知顯示名稱 ---
    # 支援直接匹配或麵包屑最後一節（如 "PC > 下載" 取 "下載"）
    for txt in toolbar_texts:
        if txt in _KNOWN_FOLDER_MAP:
            return _KNOWN_FOLDER_MAP[txt]
        # 麵包屑格式：取 ">" 分隔後最後一段
        segments = [s.strip().lstrip("> ").strip() for s in re.split(r'[>\\]', txt) if s.strip().lstrip("> ").strip()]
        for seg in reversed(segments):
            if seg in _KNOWN_FOLDER_MAP:
                return _KNOWN_FOLDER_MAP[seg]

    # --- 策略 4：遍歷所有控件，找含 ':' 的合法路徑 ---
    try:
        for child in dlg.descendants():
            try:
                txt = child.window_text().strip()
                if txt and ':' in txt and os.path.isabs(txt) and os.path.isdir(txt):
                    return txt
            except Exception:
                continue
    except Exception:
        pass

    # --- 策略 5：Fallback 預設資料夾（Downloads 優先）---
    home = os.path.expanduser("~")
    for folder in [
        os.path.join(home, "Downloads"),
        os.path.join(home, "Desktop"),
        os.path.join(home, "Documents"),
        home,
    ]:
        if os.path.isdir(folder):
            return folder

    return None


# ============================================================
# Handler 函數定義
# ============================================================

def handle_check_file_name(dlg, step: dict, value: str) -> str:
    """
    【check_file_name()】檢查儲存路徑下是否已存在同名檔案。

    若有重複，在檔名後附加時間戳 (_YYYYMMDD_HHMMSS)，確保不覆蓋舊檔。
    同時嘗試從對話框的「存檔類型」下拉選單取得副檔名，自動補全。

    使用情境：
        在「另存新檔」對話框的檔名輸入控件執行 set_text / send_keys 前，
        先用此 handle 確認是否需要加上時間戳避免覆蓋。

    JSON 範例：
        {
            "app": "Notepad",
            "Name": "檔案名稱",
            "ControlType": "UIA_EditControlTypeId",
            "Action": "set_text()",
            "value": "MyReport",
            "handle": "check_file_name()"
        }

    參數：
        dlg   : 目前操作的 pywinauto 視窗物件（需為檔案對話框）
        step  : 步驟配置 dict
        value : 原始檔案名稱（不含或含副檔名皆可）

    回傳：
        str - 原始名稱（無衝突）或加上時間戳的新名稱
    """
    folder = _get_dialog_folder_path(dlg)
    if not folder:
        logger.warning("handle_check_file_name: 無法取得儲存路徑，跳過檔案名稱檢查")
        return value

    logger.info(f"handle_check_file_name: 偵測到儲存路徑 → {folder}")

    base, ext = os.path.splitext(value)

    # 嘗試從對話框取得存檔類型副檔名
    dialog_ext = ""
    try:
        for child in dlg.descendants():
            try:
                txt = child.window_text()
                m = re.search(r'\*\.(\w+)', txt)
                if m:
                    dialog_ext = '.' + m.group(1)
                    break
            except Exception:
                continue
    except Exception:
        pass

    # 決定要檢查的副檔名清單
    if ext:
        extensions = [ext]
    elif dialog_ext:
        extensions = [dialog_ext, '']
    else:
        extensions = ['.txt', '']

    # 檢查是否有同名檔案存在
    for e in extensions:
        candidate = os.path.join(folder, base + e)
        if os.path.exists(candidate):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_value = f"{base}_{timestamp}{e}" if e else f"{base}_{timestamp}"
            logger.info(f"handle_check_file_name: 檔案 '{base}{e}' 已存在，重命名為 → {new_value}")
            return new_value

    logger.info(f"handle_check_file_name: 檔案名稱 '{value}' 無重複，維持原名")
    return value


# ============================================================
# Handle 註冊表：在 JSON 配置中以字串引用，對應實際函數
# ============================================================
HANDLE_REGISTRY = {
    "check_file_name()": handle_check_file_name,
}
