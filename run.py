
import json
import logging
import time
import sys
import re
import os
import io
import argparse
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
            ctrl = find_control(dlg, control_name, control_type, step_label)

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
