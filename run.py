
import json
import logging
import time
import sys
import re
import os
import io
import argparse
from datetime import datetime
from pywinauto import Application


# ============================================================
# 修正 Windows 打包 .exe 後的 Unicode 編碼問題
# 預設終端機編碼為 cp950，無法處理 \u200b 等特殊 Unicode 字元
# ============================================================
if os.name == 'nt':
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)  # 設定輸出為 UTF-8
        ctypes.windll.kernel32.SetConsoleCP(65001)        # 設定輸入為 UTF-8
    except Exception:
        pass

if sys.stdout is not None:
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            # 退而求其次：只設定 errors='replace'，避免無法編碼時拋出例外
            try:
                sys.stdout.reconfigure(errors='replace')
            except Exception:
                pass
    elif hasattr(sys.stdout, 'buffer'):
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        except Exception:
            pass


# ============================================================
# Logging 設定：只輸出到 Terminal，不產生 log 檔案
# ============================================================
logger = logging.getLogger("WinAutomation")
logger.setLevel(logging.DEBUG)

# Console handler - 輸出到 Terminal
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO)
console_fmt = logging.Formatter(
    "[%(asctime)s] %(levelname)-7s %(message)s",
    datefmt="%H:%M:%S",
)
console_handler.setFormatter(console_fmt)

logger.addHandler(console_handler)


# ============================================================
# UIA ControlType 對照表
# ============================================================
CONTROL_TYPE_MAP = {
    "UIA_ButtonControlTypeId": "Button",
    "UIA_TextControlTypeId": "Text",
    "UIA_ListItemControlTypeId": "ListItem",
    "UIA_SplitButtonControlTypeId": "SplitButton",
    "UIA_EditControlTypeId": "Edit",
    "UIA_ComboBoxControlTypeId": "ComboBox",
    "UIA_CheckBoxControlTypeId": "CheckBox",
    "UIA_RadioButtonControlTypeId": "RadioButton",
    "UIA_MenuItemControlTypeId": "MenuItem",
    "UIA_TabItemControlTypeId": "TabItem",
    "UIA_TreeItemControlTypeId": "TreeItem",
    "UIA_HyperlinkControlTypeId": "Hyperlink",
    "UIA_ImageControlTypeId": "Image",
    "UIA_WindowControlTypeId": "Window",
    "UIA_PaneControlTypeId": "Pane",
    "UIA_GroupControlTypeId": "Group",
    "UIA_DataItemControlTypeId": "DataItem",
    "UIA_DocumentControlTypeId": "Document",
    "UIA_CustomControlTypeId": "Custom",
    "UIA_HeaderControlTypeId": "Header",
    "UIA_HeaderItemControlTypeId": "HeaderItem",
    "UIA_ListControlTypeId": "List",
    "UIA_MenuControlTypeId": "Menu",
    "UIA_MenuBarControlTypeId": "MenuBar",
    "UIA_ProgressBarControlTypeId": "ProgressBar",
    "UIA_ScrollBarControlTypeId": "ScrollBar",
    "UIA_SliderControlTypeId": "Slider",
    "UIA_SpinnerControlTypeId": "Spinner",
    "UIA_StatusBarControlTypeId": "StatusBar",
    "UIA_TableControlTypeId": "Table",
    "UIA_TabControlTypeId": "Tab",
    "UIA_ToolBarControlTypeId": "ToolBar",
    "UIA_ToolTipControlTypeId": "ToolTip",
    "UIA_TreeControlTypeId": "Tree",
}

# 操作之間的預設等待時間 (秒)
DEFAULT_WAIT = 0.3
CONNECT_TIMEOUT = 10
CONTROL_TIMEOUT = 10
QUICK_TIMEOUT = 2   # 多視窗遍歷降級時的短暫 timeout


# ============================================================
# Handle 機制：特定步驟可掛載自訂處理邏輯
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
    toolbar_texts: list[str] = []
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


def handle_check_file_name(dlg, step: dict, value: str) -> str:
    """
    檢查儲存路徑下是否已存在同名檔案。
    若有重複，在檔名後附加時間戳 (_YYYYMMDD_HHMMSS)。
    同時嘗試從對話框的「存檔類型」下拉選單取得副檔名。
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


# Handle 註冊表：在 JSON 配置中以字串引用，對應實際函數
HANDLE_REGISTRY = {
    "check_file_name()": handle_check_file_name,
}


def parse_control_type(control_type_str: str) -> str:
    """解析 UIA_XxxControlTypeId 格式，回傳 pywinauto 友好名稱"""
    match = re.match(r"(UIA_\w+ControlTypeId)", control_type_str)
    if match:
        uia_type = match.group(1)
        friendly = CONTROL_TYPE_MAP.get(uia_type)
        if friendly:
            return friendly
        logger.warning(f"未知的 ControlType: {uia_type}，將使用原始值")
    return control_type_str


def load_config(config_path: str = None) -> dict:
    """讀取 exec.json 設定檔"""
    if config_path is None:
        config_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "exec.json"
        )

    logger.info(f"讀取設定檔: {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    logger.info(f"共載入 {len(config)} 個步驟")
    return config


def execute_action(ctrl, action_str: str, control_name: str, value: str = ""):
    """
    根據 Action 字串執行對應操作。
    click() 使用 UIA Invoke Pattern（程式化觸發，不依賴螢幕座標）。
    click_input() 使用物理滑鼠點擊（需元素在螢幕可視區域）。
    若 value 不為空，則優先作為 type_keys / set_text / send_keys 的輸入內容。
    """
    action_lower = action_str.strip().lower()

    if action_lower in ("click", "click()"):
        # 使用 UIA Invoke Pattern，不移動滑鼠，適用於不在螢幕可視區域的控件
        try:
            ctrl.invoke()
        except Exception:
            # 若 Invoke Pattern 不支援，降級為先捲動再物理點擊
            logger.warning(f"invoke() 不支援，降級為 scroll + click_input(): '{control_name}'")
            try:
                ctrl.scroll_into_view()
            except Exception:
                pass
            ctrl.click_input()
    elif action_lower in ("click_input", "click_input()"):
        # 物理滑鼠點擊，先嘗試捲動到可視區域
        try:
            ctrl.scroll_into_view()
        except Exception:
            pass
        ctrl.click_input()
    elif action_lower in ("double_click", "double_click()", "double_click_input", "double_click_input()"):
        ctrl.double_click_input()
    elif action_lower in ("right_click", "right_click_input", "right_click_input()"):
        ctrl.right_click_input()
    elif action_lower in ("send_keys", "send_keys()"):
        # send_keys() 必須搭配 value 使用
        if not value:
            logger.error(f"send_keys() 需要提供 value，但 value 為空: '{control_name}'")
            raise ValueError(f"send_keys() 需要提供 value: '{control_name}'")
        ctrl.type_keys(value, with_spaces=True)
    elif action_lower.startswith("type_keys("):
        # 若有 value，優先使用 value；否則解析 action 字串內嵌文字
        if value:
            ctrl.type_keys(value, with_spaces=True)
        else:
            match = re.match(r"type_keys\(['\"](.+)['\"]\)", action_str.strip())
            if match:
                keys = match.group(1)
                ctrl.type_keys(keys, with_spaces=True)
            else:
                logger.error(f"無法解析 type_keys 參數: {action_str}")
                raise ValueError(f"無法解析 type_keys 參數: {action_str}")
    elif action_lower.startswith("set_text("):
        # 若有 value，優先使用 value；否則解析 action 字串內嵌文字
        if value:
            ctrl.set_text(value)
        else:
            match = re.match(r"set_text\(['\"](.+)['\"]\)", action_str.strip())
            if match:
                text = match.group(1)
                ctrl.set_text(text)
            else:
                logger.error(f"無法解析 set_text 參數: {action_str}")
                raise ValueError(f"無法解析 set_text 參數: {action_str}")
    else:
        logger.warning(f"未知的 Action '{action_str}'，預設使用 click_input()")
        ctrl.click_input()


def quick_find_in_window(dlg, control_name: str, control_type: str) -> object:
    """
    快速在指定視窗中搜尋控件（短暫 timeout），用於多視窗遍歷降級。
    找不到時拋出 RuntimeError。
    """
    if not control_name:
        ctrl = dlg.child_window(control_type=control_type)
        ctrl.wait("exists", timeout=QUICK_TIMEOUT)
        return ctrl

    # 策略 1：精確搜尋
    try:
        ctrl = dlg.child_window(title=control_name, control_type=control_type)
        ctrl.wait("exists", timeout=QUICK_TIMEOUT)
        return ctrl
    except Exception:
        pass

    # 策略 2：僅用 title
    try:
        ctrl = dlg.child_window(title=control_name)
        ctrl.wait("exists", timeout=QUICK_TIMEOUT)
        return ctrl
    except Exception:
        pass

    # 策略 3：遍歷 descendants 直接比對
    try:
        for child in dlg.descendants():
            try:
                if child.window_text() == control_name:
                    return child
            except Exception:
                continue
    except Exception:
        pass

    raise RuntimeError(f"quick_find: 找不到控件 '{control_name}'")


def find_control(dlg, control_name: str, control_type: str, step_label: str):
    """
    多層降級搜尋控件：
      1. title + control_type  (精確)
      2. title only            (忽略類型)
      3. title_re 正則局部匹配 (容忍空白/特殊字元差異)
      4. 遍歷 descendants() 直接比對 (適用大型頁面 / 深層巢狀控件)
      5. control_type only     (Name 為空或前四策略均失敗時，僅依類型搜尋)
    找不到時印出所有可見控件清單以利除錯。
    """

    # --- Name 為空：直接用 control_type 搜尋，跳過 title 相關策略 ---
    if not control_name:
        try:
            ctrl = dlg.child_window(control_type=control_type)
            ctrl.wait("exists", timeout=CONTROL_TIMEOUT)
            logger.info(f"{step_label} 已找到控件 [類型搜尋]: (type={control_type})")
            return ctrl
        except Exception:
            logger.warning(f"{step_label} 類型搜尋失敗，嘗試遍歷所有子控件 ...")
        # 遍歷 descendants 僅比對 control_type
        try:
            for child in dlg.descendants():
                try:
                    ct = child.friendly_class_name()
                    if ct == control_type:
                        logger.warning(f"{step_label} 已找到控件 [遍歷類型]: (type={ct})")
                        return child
                except Exception:
                    continue
        except Exception:
            pass
        # 跳到最後的錯誤輸出
    else:
        # --- 策略 1：精確搜尋 ---
        try:
            ctrl = dlg.child_window(title=control_name, control_type=control_type)
            ctrl.wait("exists", timeout=CONTROL_TIMEOUT)
            logger.info(f"{step_label} 已找到控件 [精確]: '{control_name}' ({control_type})")
            return ctrl
        except Exception:
            logger.warning(f"{step_label} 精確搜尋失敗，嘗試僅用 title ...")

        # --- 策略 2：僅用 title，不限類型 ---
        try:
            ctrl = dlg.child_window(title=control_name)
            ctrl.wait("exists", timeout=CONTROL_TIMEOUT)
            logger.warning(f"{step_label} 已找到控件 [僅 title]: '{control_name}'（控件類型可能不同）")
            return ctrl
        except Exception:
            logger.warning(f"{step_label} 僅 title 搜尋失敗，嘗試正則局部匹配 ...")

        # --- 策略 3：正則局部匹配（逸脫特殊字元後部分匹配）---
        try:
            pattern = re.escape(control_name)
            ctrl = dlg.child_window(title_re=pattern)
            ctrl.wait("exists", timeout=CONTROL_TIMEOUT)
            logger.warning(f"{step_label} 已找到控件 [正則]: '{control_name}'")
            return ctrl
        except Exception:
            logger.warning(f"{step_label} 正則搜尋失敗，嘗試遍歷所有子控件 ...")

        # --- 策略 4：遍歷 descendants() 直接比對 ---
        try:
            best_match = None
            for child in dlg.descendants():
                try:
                    txt = child.window_text()
                    if txt == control_name:
                        ct = child.friendly_class_name()
                        if ct == control_type:
                            logger.warning(f"{step_label} 已找到控件 [遍歷精確]: '{control_name}' ({ct})")
                            return child
                        elif best_match is None:
                            best_match = child
                except Exception:
                    continue
            if best_match is not None:
                ct = best_match.friendly_class_name()
                logger.warning(f"{step_label} 已找到控件 [遍歷]: '{control_name}' ({ct})")
                return best_match
        except Exception:
            pass

        # --- 策略 5：僅用 control_type（最終降級）---
        if control_type:
            try:
                ctrl = dlg.child_window(control_type=control_type)
                ctrl.wait("exists", timeout=CONTROL_TIMEOUT)
                logger.warning(f"{step_label} 已找到控件 [僅類型]: (type={control_type})，title 不符但仍嘗試")
                return ctrl
            except Exception:
                pass
            try:
                for child in dlg.descendants():
                    try:
                        ct = child.friendly_class_name()
                        if ct == control_type:
                            logger.warning(f"{step_label} 已找到控件 [遍歷僅類型]: (type={ct})")
                            return child
                    except Exception:
                        continue
            except Exception:
                pass

    # --- 全部失敗：印出目前可見控件清單協助除錯 ---
    logger.error(f"{step_label} 四種搜尋策略均失敗，列出目前所有可見控件：")
    try:
        for i, child in enumerate(dlg.descendants()):
            try:
                txt = child.window_text()
                ct = child.friendly_class_name()
                logger.error(f"  [{i:03d}] type={ct:20s} | title='{txt}'")
            except Exception:
                pass
    except Exception as dump_err:
        logger.error(f"  (無法列出子控件: {dump_err})")

    raise RuntimeError(f"找不到控件: '{control_name}' (type={control_type})")


def execute_steps(config: dict):
    """依序執行 exec.json 中定義的所有步驟"""
    sorted_keys = sorted(config.keys(), key=lambda x: int(x))
    total = len(sorted_keys)

    logger.info("=" * 60)
    logger.info(f"開始執行自動化流程，共 {total} 個步驟")
    logger.info("=" * 60)

    # 快取已連線的 app，避免相同視窗重複 connect()
    app_cache: dict = {}

    for idx, key in enumerate(sorted_keys, start=1):
        step = config[key]
        app_title = step["app"]
        control_name = step["Name"]
        control_type = parse_control_type(step["ControlType"])
        action = step.get("Action", "click_input()")
        value = step.get("value", "")
        wait_time = float(step.get("Wait", DEFAULT_WAIT))
        step_label = f"[步驟 {idx}/{total}]"

        logger.info("-" * 60)
        logger.info(f"{step_label} 開始執行")
        logger.info(f"  應用程式 : {app_title}")
        logger.info(f"  控件名稱 : {control_name}")
        logger.info(f"  控件類型 : {control_type}")
        logger.info(f"  執行動作 : {action}")

        try:
            # 連接到目標應用程式（快取複用，避免重複 connect）
            if app_title not in app_cache:
                logger.info(f"{step_label} 正在連接應用程式: {app_title} ...")
                app = Application(backend="uia").connect(
                    title=app_title, timeout=CONNECT_TIMEOUT
                )
                app_cache[app_title] = app
                logger.info(f"{step_label} 已連接到: {app_title}")
            else:
                app = app_cache[app_title]
                logger.info(f"{step_label} 複用已連線的應用程式: {app_title}")

            # 取得視窗：先精確比對標題，若失敗則局部匹配（如記事本輸入後標題加 *），
            # 最終降級為 top_window()
            dlg = app.window(title=app_title)
            try:
                dlg.wrapper_object()
            except Exception:
                try:
                    dlg = app.window(title_re=f".*{re.escape(app_title)}.*")
                    dlg.wrapper_object()
                    logger.info(f"{step_label} 視窗標題已變更，已透過局部匹配重新找到視窗")
                except Exception:
                    dlg = app.top_window()
                    logger.info(f"{step_label} 使用 top_window() 作為視窗降級")

            # 尋找目標控件（多層降級搜尋）
            logger.info(f"{step_label} 正在尋找控件: '{control_name}' ...")
            try:
                ctrl = find_control(dlg, control_name, control_type, step_label)
            except RuntimeError:
                # 主視窗找不到（如自動完成下拉視窗遮蔽），嘗試遍歷應用程式所有視窗
                logger.warning(f"{step_label} 主視窗找不到控件，嘗試遍歷所有視窗 ...")
                ctrl = None
                try:
                    for alt_win in app.windows():
                        try:
                            if alt_win.handle == dlg.handle:
                                continue
                            ctrl = quick_find_in_window(alt_win, control_name, control_type)
                            dlg = alt_win
                            logger.info(f"{step_label} 已在替代視窗找到控件: '{control_name}'")
                            break
                        except Exception:
                            continue
                except Exception as enum_err:
                    logger.warning(f"{step_label} 遍歷視窗時發生錯誤: {enum_err}")
                if ctrl is None:
                    logger.error(f"{step_label} 所有視窗均找不到控件: '{control_name}' (type={control_type})")
                    raise RuntimeError(f"找不到控件: '{control_name}' (type={control_type})")

            # 處理 handle（在執行動作前呼叫自訂邏輯，可修改 value）
            handle = step.get("handle", "")
            if handle:
                if handle in HANDLE_REGISTRY:
                    logger.info(f"{step_label} 執行 handle: {handle}")
                    value = HANDLE_REGISTRY[handle](dlg, step, value)
                    logger.info(f"{step_label} handle 處理後 value = '{value}'")
                else:
                    logger.warning(f"{step_label} 未知的 handle: '{handle}'，已跳過")

            # 執行操作
            logger.info(f"{step_label} 正在執行: {action} -> '{control_name}' (value='{value}')")
            execute_action(ctrl, action, control_name, value)

            logger.info(f"{step_label} 完成")

            # 等待 UI 響應
            if idx < total:
                logger.info(f"{step_label} 等待 {wait_time} 秒...")
                time.sleep(wait_time)

        except Exception as e:
            logger.error(f"{step_label} 執行失敗: {e}")
            logger.error("停止自動化流程。")
            sys.exit(1)

    logger.info("=" * 60)
    logger.info(f"全部 {total} 個步驟執行完畢!")
    logger.info("=" * 60)


def main():
    parser = argparse.ArgumentParser(description="Win_Automation - Windows GUI 自動化腳本")
    parser.add_argument("--case", type=str, default="exec.json", help="指定要執行的 JSON 設定檔 (預設: exec.json)")
    args = parser.parse_args()

    start_time = time.time()
    logger.info("Win_Automation 啟動")
    config = load_config(args.case)
    execute_steps(config)
    elapsed = time.time() - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)
    logger.info(f"總共花費時間: {minutes:02d}分{seconds:02d}秒")
    logger.info("Win_Automation 結束")


if __name__ == "__main__":
    main()
