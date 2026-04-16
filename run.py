import json
import logging
import time
import sys
import re
import os
from pywinauto import Application


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
CONTROL_TIMEOUT = 3


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


def execute_action(ctrl, action_str: str, control_name: str):
    """
    根據 Action 字串執行對應操作。
    所有 click 類操作統一使用 click_input()。
    """
    action_lower = action_str.strip().lower()

    if action_lower in ("click", "click()", "click_input", "click_input()"):
        ctrl.click_input()
    elif action_lower in ("double_click", "double_click_input", "double_click_input()"):
        ctrl.double_click_input()
    elif action_lower in ("right_click", "right_click_input", "right_click_input()"):
        ctrl.right_click_input()
    elif action_lower.startswith("type_keys("):
        # 支援 type_keys('文字') 格式
        match = re.match(r"type_keys\(['\"](.+)['\"]\)", action_str.strip())
        if match:
            keys = match.group(1)
            ctrl.type_keys(keys, with_spaces=True)
        else:
            logger.error(f"無法解析 type_keys 參數: {action_str}")
            raise ValueError(f"無法解析 type_keys 參數: {action_str}")
    elif action_lower.startswith("set_text("):
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
    找不到時印出所有可見控件清單以利除錯。
    """

    # --- 策略 1：精確搜尋 ---
    try:
        ctrl = dlg.child_window(title=control_name, control_type=control_type)
        ctrl.wait("visible", timeout=CONTROL_TIMEOUT)
        logger.info(f"{step_label} 已找到控件 [精確]: '{control_name}' ({control_type})")
        return ctrl
    except Exception:
        logger.warning(f"{step_label} 精確搜尋失敗，嘗試僅用 title ...")

    # --- 策略 2：僅用 title，不限類型 ---
    try:
        ctrl = dlg.child_window(title=control_name)
        ctrl.wait("visible", timeout=CONTROL_TIMEOUT)
        logger.warning(f"{step_label} 已找到控件 [僅 title]: '{control_name}'（控件類型可能不同）")
        return ctrl
    except Exception:
        logger.warning(f"{step_label} 僅 title 搜尋失敗，嘗試正則局部匹配 ...")

    # --- 策略 3：正則局部匹配（逸脫特殊字元後部分匹配）---
    try:
        pattern = re.escape(control_name)
        ctrl = dlg.child_window(title_re=pattern)
        ctrl.wait("visible", timeout=CONTROL_TIMEOUT)
        logger.warning(f"{step_label} 已找到控件 [正則]: '{control_name}'")
        return ctrl
    except Exception:
        pass

    # --- 全部失敗：印出目前可見控件清單協助除錯 ---
    logger.error(f"{step_label} 三種搜尋策略均失敗，列出目前所有可見控件：")
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
            dlg = app.window(title=app_title)

            # 尋找目標控件（多層降級搜尋）
            logger.info(f"{step_label} 正在尋找控件: '{control_name}' ...")
            ctrl = find_control(dlg, control_name, control_type, step_label)

            # 執行操作
            logger.info(f"{step_label} 正在執行: click_input() -> '{control_name}'")
            execute_action(ctrl, action, control_name)

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
    start_time = time.time()
    logger.info("Win_Automation 啟動")
    config = load_config()
    execute_steps(config)
    elapsed = time.time() - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)
    logger.info(f"總共花費時間: {minutes:02d}分{seconds:02d}秒")
    logger.info("Win_Automation 結束")


if __name__ == "__main__":
    main()
