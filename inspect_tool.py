"""
WinAutomation UI 元素偵測工具 (inspect_tool.py)
====================================
類似 inspect.exe 的 Tkinter GUI 小工具，
可即時偵測滑鼠下方的 UI 元素資訊，並匯出 set.json。

使用步驟：
  1. 點選「開始偵測」→ 移動滑鼠到任意元件，即時顯示資訊
  2. 按 Ctrl+F1 記錄該元件，選擇 Action / 填入 value
  3. 點選「結束偵測」→ 匯出 set.json
"""

import ctypes
import ctypes.wintypes
import json
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# COM (隨 pywinauto 一起安裝)
import comtypes
import comtypes.client

# 自訂 handler 模組（取得已註冊的 handler 清單）
try:
    from handler import HANDLE_REGISTRY as _HANDLE_REGISTRY
except Exception:
    _HANDLE_REGISTRY = {}


# ============================================================
# UIA ControlType ID → 顯示名稱 對照表
# ============================================================
CONTROL_TYPE_ID_MAP = {
    50000: ("UIA_ButtonControlTypeId", "0xC350"),
    50001: ("UIA_CalendarControlTypeId", "0xC351"),
    50002: ("UIA_CheckBoxControlTypeId", "0xC352"),
    50003: ("UIA_ComboBoxControlTypeId", "0xC353"),
    50004: ("UIA_EditControlTypeId", "0xC354"),
    50005: ("UIA_HyperlinkControlTypeId", "0xC355"),
    50006: ("UIA_ImageControlTypeId", "0xC356"),
    50007: ("UIA_ListItemControlTypeId", "0xC357"),
    50008: ("UIA_ListControlTypeId", "0xC358"),
    50009: ("UIA_MenuControlTypeId", "0xC359"),
    50010: ("UIA_MenuBarControlTypeId", "0xC35A"),
    50011: ("UIA_MenuItemControlTypeId", "0xC35B"),
    50012: ("UIA_ProgressBarControlTypeId", "0xC35C"),
    50013: ("UIA_RadioButtonControlTypeId", "0xC35D"),
    50014: ("UIA_ScrollBarControlTypeId", "0xC35E"),
    50015: ("UIA_SliderControlTypeId", "0xC35F"),
    50016: ("UIA_SpinnerControlTypeId", "0xC360"),
    50017: ("UIA_StatusBarControlTypeId", "0xC361"),
    50018: ("UIA_TabControlTypeId", "0xC362"),
    50019: ("UIA_TabItemControlTypeId", "0xC363"),
    50020: ("UIA_TextControlTypeId", "0xC364"),
    50021: ("UIA_ToolBarControlTypeId", "0xC365"),
    50022: ("UIA_ToolTipControlTypeId", "0xC366"),
    50023: ("UIA_TreeControlTypeId", "0xC367"),
    50024: ("UIA_TreeItemControlTypeId", "0xC368"),
    50025: ("UIA_CustomControlTypeId", "0xC369"),
    50026: ("UIA_GroupControlTypeId", "0xC36A"),
    50027: ("UIA_ThumbControlTypeId", "0xC36B"),
    50028: ("UIA_DataGridControlTypeId", "0xC36C"),
    50029: ("UIA_DataItemControlTypeId", "0xC36D"),
    50030: ("UIA_DocumentControlTypeId", "0xC36E"),
    50031: ("UIA_SplitButtonControlTypeId", "0xC36F"),
    50032: ("UIA_WindowControlTypeId", "0xC370"),
    50033: ("UIA_PaneControlTypeId", "0xC371"),
    50034: ("UIA_HeaderControlTypeId", "0xC372"),
    50035: ("UIA_HeaderItemControlTypeId", "0xC373"),
    50036: ("UIA_TableControlTypeId", "0xC374"),
    50037: ("UIA_TitleBarControlTypeId", "0xC375"),
    50038: ("UIA_SeparatorControlTypeId", "0xC376"),
    50039: ("UIA_SemanticZoomControlTypeId", "0xC377"),
    50040: ("UIA_AppBarControlTypeId", "0xC378"),
}

# Action 下拉選項 (顯示文字含中文說明, 實際儲存值)
_ACTION_OPTIONS_RAW = [
    # --- 點擊 ---
    ("click()",              "click()  發送滑鼠點擊訊息"),
    ("click_input()",        "click_input()  使用滑鼠事件進行點擊"),
    ("double_click()",       "double_click()  雙擊"),
    ("double_click_input()", "double_click_input()  雙擊輸入"),
    ("right_click()",        "right_click()  右鍵點擊"),
    ("right_click_input()",  "right_click_input()  右鍵點擊輸入"),
    # --- 鍵盤 ---
    ("send_keys()",          "send_keys()  發送鍵盤事件"),
    ("type_keys()",          "type_keys()  發送鍵盤輸入"),
    ("set_text()",           "set_text()  直接設定文字"),
    # --- 訊息 ---
    ("send_message()",       "send_message()  發送 Windows 訊息"),
    ("send_message_timeout()", "send_message_timeout()  發送訊息並等待回應"),
    # --- 焦點 / 文字 ---
    ("set_focus()",          "set_focus()  設定焦點"),
    ("get_focus()",          "get_focus()  取得焦點"),
    ("set_window_text()",    "set_window_text()  設定視窗文字"),
    # --- 視窗控制 ---
    ("close()",              "close()  關閉視窗"),
    ("close_click()",        "close_click()  點擊關閉"),
    ("maximize()",           "maximize()  最大化"),
    ("minimize()",           "minimize()  最小化"),
    ("restore()",            "restore()  還原"),
    ("get_show_state()",     "get_show_state()  取得顯示狀態"),
    ("move_window()",        "move_window()  移動與調整大小"),
    # --- 滑鼠移動 / 拖曳 ---
    ("drag_mouse()",         "drag_mouse()  拖動滑鼠"),
    ("move_mouse()",         "move_mouse()  移動滑鼠"),
    ("press_mouse()",        "press_mouse()  按下滑鼠按鍵"),
    ("release_mouse()",      "release_mouse()  放開滑鼠按鍵"),
    ("press_mouse_input()",  "press_mouse_input()  按下滑鼠（基於事件）"),
    ("release_mouse_input()", "release_mouse_input()  放開滑鼠（基於事件）"),
    # --- 選單 ---
    ("menu_select()",        "menu_select()  選擇菜單項目"),
    ("notify_menu_select()", "notify_menu_select()  通知菜單選擇"),
    ("notify_parent()",      "notify_parent()  通知父視窗"),
    # --- 其他 ---
    ("draw_outline()",       "draw_outline()  繪製輪廓"),
    ("rectangle",            "rectangle  取得矩形座標（屬性）"),
    ("children",             "children  取得子視窗列表（屬性）"),
]
ACTION_OPTIONS = [display for _, display in _ACTION_OPTIONS_RAW]
ACTION_VALUE_MAP = {display: value for value, display in _ACTION_OPTIONS_RAW}

# Handler 下拉選項
HANDLER_OPTIONS = list(_HANDLE_REGISTRY.keys())

# 偵測輪詢間隔 (毫秒)
POLL_INTERVAL_MS = 150


# ============================================================
# 格式化 ControlType
# ============================================================
def format_control_type(control_type_id: int) -> str:
    """將 ControlType ID 轉為 'UIA_XxxControlTypeId (0xXXXX)' 格式"""
    entry = CONTROL_TYPE_ID_MAP.get(control_type_id)
    if entry:
        return f"{entry[0]} ({entry[1]})"
    return f"Unknown ({hex(control_type_id)})"


# ============================================================
# UIA 偵測引擎
# ============================================================
class UIAInspector:
    """透過 COM 直接存取 UI Automation，偵測滑鼠下方元素"""

    def __init__(self):
        comtypes.CoInitialize()
        self._load_uia()

    def _load_uia(self):
        comtypes.client.GetModule("UIAutomationCore.dll")
        from comtypes.gen.UIAutomationClient import (
            IUIAutomation,
            CUIAutomation,
            tagPOINT,
        )
        self._tagPOINT = tagPOINT
        self._uia = comtypes.CoCreateInstance(
            CUIAutomation._reg_clsid_,
            interface=IUIAutomation,
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
        )
        self._root = self._uia.GetRootElement()
        self._walker = self._uia.ControlViewWalker

    # --------------------------------------------------------
    def get_element_info(self, x: int, y: int) -> dict | None:
        """回傳 {app, Name, ControlType, AutomationId, ClassName} 或 None"""
        try:
            pt = self._tagPOINT(x, y)
            elem = self._uia.ElementFromPoint(pt)
            if elem is None:
                return None

            name = elem.CurrentName or ""
            control_type_id = elem.CurrentControlType
            control_type_str = format_control_type(control_type_id)
            app_name = self._get_top_window_name(elem)
            auto_id = elem.CurrentAutomationId or ""
            class_name = elem.CurrentClassName or ""

            return {
                "app": app_name,
                "Name": name,
                "ControlType": control_type_str,
                "AutomationId": auto_id,
                "ClassName": class_name,
                "_elem": elem,  # 保留元素參考，供詳細資訊使用
            }
        except Exception:
            return None

    # --------------------------------------------------------
    def get_detailed_info(self, elem) -> str:
        """從 UIA 元素取得完整屬性，回傳格式化文字"""
        lines = []

        def _safe(func, default=""):
            try:
                v = func()
                return v if v is not None else default
            except Exception:
                return default

        # --- 基本屬性 ---
        lines.append(f"Name:\t\"{_safe(lambda: elem.CurrentName)}\"")
        ct_id = _safe(lambda: elem.CurrentControlType, 0)
        lines.append(f"ControlType:\t{format_control_type(ct_id) if ct_id else 'Unknown'}")
        lines.append(f"LocalizedControlType:\t\"{_safe(lambda: elem.CurrentLocalizedControlType)}\"")

        # BoundingRectangle
        try:
            rect = elem.CurrentBoundingRectangle
            lines.append(f"BoundingRectangle:\t{{l:{rect.left} t:{rect.top} r:{rect.right} b:{rect.bottom}}}")
        except Exception:
            lines.append("BoundingRectangle:\t")

        lines.append(f"IsEnabled:\t{str(_safe(lambda: elem.CurrentIsEnabled, False)).lower()}")
        lines.append(f"IsOffscreen:\t{str(_safe(lambda: elem.CurrentIsOffscreen, False)).lower()}")
        lines.append(f"IsKeyboardFocusable:\t{str(_safe(lambda: elem.CurrentIsKeyboardFocusable, False)).lower()}")
        lines.append(f"HasKeyboardFocus:\t{str(_safe(lambda: elem.CurrentHasKeyboardFocus, False)).lower()}")
        lines.append(f"ProcessId:\t{_safe(lambda: elem.CurrentProcessId, 0)}")

        # RuntimeId
        try:
            rid = elem.GetRuntimeId()
            if rid:
                hex_parts = [f"{x:X}" for x in rid]
                lines.append(f"RuntimeId:\t[{'.'.join(hex_parts)}]")
            else:
                lines.append("RuntimeId:\t")
        except Exception:
            lines.append("RuntimeId:\t")

        lines.append(f"AutomationId:\t\"{_safe(lambda: elem.CurrentAutomationId)}\"")
        lines.append(f"FrameworkId:\t\"{_safe(lambda: elem.CurrentFrameworkId)}\"")
        lines.append(f"ClassName:\t\"{_safe(lambda: elem.CurrentClassName)}\"")
        lines.append(f"IsControlElement:\t{str(_safe(lambda: elem.CurrentIsControlElement, False)).lower()}")
        lines.append(f"IsContentElement:\t{str(_safe(lambda: elem.CurrentIsContentElement, False)).lower()}")
        lines.append(f"ProviderDescription:\t\"{_safe(lambda: elem.CurrentProviderDescription)}\"")

        # 額外屬性（可能不支援）
        lines.append(f"IsPeripheral:\t{str(_safe(lambda: elem.CurrentIsPeripheral, False)).lower()}")
        lines.append(f"AriaRole:\t\"{_safe(lambda: elem.CurrentAriaRole)}\"")
        lines.append(f"AriaProperties:\t\"{_safe(lambda: elem.CurrentAriaProperties)}\"")
        lines.append(f"IsPassword:\t{str(_safe(lambda: elem.CurrentIsPassword, False)).lower()}")
        lines.append(f"IsRequiredForForm:\t{str(_safe(lambda: elem.CurrentIsRequiredForForm, False)).lower()}")
        lines.append(f"IsDataValidForForm:\t{str(_safe(lambda: elem.CurrentIsDataValidForForm, False)).lower()}")
        lines.append(f"HelpText:\t\"{_safe(lambda: elem.CurrentHelpText)}\"")

        # ClickablePoint
        try:
            pt = elem.GetClickablePoint()
            if pt and hasattr(pt, '__len__') and len(pt) == 2:
                clickable, point = pt
                if clickable:
                    lines.append(f"ClickablePoint:\t{{x:{int(point.x)} y:{int(point.y)}}}")
                else:
                    lines.append("ClickablePoint:\t")
            else:
                lines.append("ClickablePoint:\t")
        except Exception:
            lines.append("ClickablePoint:\t")

        lines.append(f"Culture:\t{_safe(lambda: elem.CurrentCulture, 0)}")
        lines.append(f"Orientation:\t{_safe(lambda: elem.CurrentOrientation, 0)}")
        lines.append(f"FullDescription:\t\"{_safe(lambda: elem.CurrentFullDescription)}\"")
        lines.append(f"IsDialog:\t{str(_safe(lambda: elem.CurrentIsDialog, False)).lower()}")

        # --- Pattern 可用性 ---
        lines.append("")
        lines.append("--- Pattern Availability ---")
        pattern_props = [
            ("IsInvokePatternAvailable", 10000),
            ("IsSelectionPatternAvailable", 10001),
            ("IsValuePatternAvailable", 10002),
            ("IsRangeValuePatternAvailable", 10003),
            ("IsScrollPatternAvailable", 10004),
            ("IsExpandCollapsePatternAvailable", 10005),
            ("IsGridPatternAvailable", 10006),
            ("IsGridItemPatternAvailable", 10007),
            ("IsMultipleViewPatternAvailable", 10008),
            ("IsWindowPatternAvailable", 10009),
            ("IsSelectionItemPatternAvailable", 10010),
            ("IsDockPatternAvailable", 10011),
            ("IsTablePatternAvailable", 10012),
            ("IsTableItemPatternAvailable", 10013),
            ("IsTextPatternAvailable", 10014),
            ("IsTogglePatternAvailable", 10015),
            ("IsTransformPatternAvailable", 10016),
            ("IsScrollItemPatternAvailable", 10017),
            ("IsLegacyIAccessiblePatternAvailable", 10018),
            ("IsItemContainerPatternAvailable", 10019),
            ("IsSynchronizedInputPatternAvailable", 10020),
            ("IsAnnotationPatternAvailable", 10023),
            ("IsTextPattern2Available", 10024),
            ("IsStylesPatternAvailable", 10025),
            ("IsSpreadsheetPatternAvailable", 10026),
            ("IsSpreadsheetItemPatternAvailable", 10027),
            ("IsTransform2PatternAvailable", 10028),
            ("IsTextEditPatternAvailable", 10029),
            ("IsCustomNavigationPatternAvailable", 10033),
            ("IsSelectionPattern2Available", 10034),
            ("IsDragPatternAvailable", 10035),
            ("IsDropTargetPatternAvailable", 10036),
            ("IsObjectModelPatternAvailable", 10022),
            ("IsTextChildPatternAvailable", 10030),
            ("IsVirtualizedItemPatternAvailable", 10020),
        ]
        seen = set()
        for pname, pid in pattern_props:
            if pname in seen:
                continue
            seen.add(pname)
            try:
                val = elem.GetCurrentPropertyValue(pid)
                lines.append(f"{pname}:\t{str(bool(val)).lower()}")
            except Exception:
                lines.append(f"{pname}:\t")

        # --- 樹狀結構 ---
        lines.append("")
        lines.append("--- Tree Navigation ---")
        for label, getter in [
            ("FirstChild", lambda: self._walker.GetFirstChildElement(elem)),
            ("LastChild", lambda: self._walker.GetLastChildElement(elem)),
            ("Next", lambda: self._walker.GetNextSiblingElement(elem)),
            ("Previous", lambda: self._walker.GetPreviousSiblingElement(elem)),
        ]:
            try:
                sibling = getter()
                if sibling:
                    sname = sibling.CurrentName or ""
                    stype = format_control_type(sibling.CurrentControlType)
                    short_type = stype.split("_")[1].replace("ControlTypeId", "").strip() if "_" in stype else stype
                    lines.append(f"{label}:\t\"{sname}\" {short_type}")
                else:
                    lines.append(f"{label}:\t[null]")
            except Exception:
                lines.append(f"{label}:\t[null]")

        # --- Ancestors ---
        lines.append("")
        lines.append("--- Ancestors ---")
        try:
            current = elem
            for _ in range(32):
                parent = self._walker.GetParentElement(current)
                if parent is None:
                    break
                if self._uia.CompareElements(parent, self._root):
                    lines.append("\t[ No Parent ]")
                    break
                pname = parent.CurrentName or ""
                ptype = format_control_type(parent.CurrentControlType)
                short_type = ptype.split("_")[1].replace("ControlTypeId", "").strip() if "_" in ptype else ptype
                lines.append(f"\t\"{pname}\" {short_type}")
                current = parent
        except Exception:
            pass

        return "\n".join(lines)

    # --------------------------------------------------------
    def _get_top_window_name(self, element) -> str:
        """沿 UIA 樹往上走，找到頂層視窗名稱"""
        try:
            current = element
            for _ in range(64):  # 防止無窮迴圈
                parent = self._walker.GetParentElement(current)
                if parent is None:
                    break
                if self._uia.CompareElements(parent, self._root):
                    # current 就是頂層視窗
                    return current.CurrentName or ""
                current = parent
            return current.CurrentName or ""
        except Exception:
            return ""


# ============================================================
# 全域快捷鍵監聽 (Ctrl+F1)
# ============================================================
class HotkeyListener:
    """使用 Windows RegisterHotKey API 監聽多組快捷鍵"""

    _MOD_CONTROL = 0x0002
    _VK_F1 = 0x70
    _VK_F2 = 0x71
    _WM_HOTKEY = 0x0312
    _ID_F1 = 1
    _ID_F2 = 2

    def __init__(self, callback_f1, callback_f2=None):
        self._callback_f1 = callback_f1
        self._callback_f2 = callback_f2
        self._running = True
        self._thread = threading.Thread(target=self._listen, daemon=True)
        self._thread.start()

    def _listen(self):
        user32 = ctypes.windll.user32
        ok1 = user32.RegisterHotKey(None, self._ID_F1, self._MOD_CONTROL, self._VK_F1)
        ok2 = user32.RegisterHotKey(None, self._ID_F2, self._MOD_CONTROL, self._VK_F2)
        if not ok1 and not ok2:
            return

        msg = ctypes.wintypes.MSG()
        while self._running:
            ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if ret <= 0:
                break
            if msg.message == self._WM_HOTKEY:
                if msg.wParam == self._ID_F1 and self._callback_f1:
                    self._callback_f1()
                elif msg.wParam == self._ID_F2 and self._callback_f2:
                    self._callback_f2()

        if ok1:
            user32.UnregisterHotKey(None, self._ID_F1)
        if ok2:
            user32.UnregisterHotKey(None, self._ID_F2)

    def stop(self):
        self._running = False
        # 發送 WM_QUIT 讓 GetMessageW 返回
        if self._thread.is_alive():
            ctypes.windll.user32.PostThreadMessageW(
                self._thread.native_id, 0x0012, 0, 0  # WM_QUIT
            )


# ============================================================
# 詳細資訊對話框
# ============================================================
class DetailDialog(tk.Toplevel):
    """顯示 UI 元素的完整 UIA 屬性"""

    def __init__(self, parent, detail_text: str):
        super().__init__(parent)
        self.title("元件詳細資訊")
        self.resizable(True, True)
        self.grab_set()
        self.transient(parent)

        # 文字區域 + 捲軸
        text_frame = ttk.Frame(self, padding=8)
        text_frame.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(text_frame, orient="vertical")
        self._text = tk.Text(
            text_frame,
            wrap="none",
            width=90,
            height=35,
            font=("Consolas", 10),
            yscrollcommand=scrollbar.set,
        )
        h_scrollbar = ttk.Scrollbar(text_frame, orient="horizontal", command=self._text.xview)
        self._text.configure(xscrollcommand=h_scrollbar.set)
        scrollbar.configure(command=self._text.yview)

        self._text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        self._text.insert("1.0", detail_text)
        self._text.configure(state="disabled")

        # 按鈕列
        btn_frame = ttk.Frame(self, padding=8)
        btn_frame.pack(fill="x")

        ttk.Button(btn_frame, text="複製全部", command=self._copy_all, width=12, bootstyle=INFO).pack(
            side="left", padx=4
        )
        ttk.Button(btn_frame, text="關閉", command=self.destroy, width=10, bootstyle=SECONDARY).pack(
            side="right", padx=4
        )

        # 置中
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = parent.winfo_x() + (parent.winfo_width() - w) // 2
        y = parent.winfo_y() + (parent.winfo_height() - h) // 2
        self.geometry(f"+{x}+{y}")

    def _copy_all(self):
        self.clipboard_clear()
        self.clipboard_append(self._text.get("1.0", "end-1c"))


# ============================================================
# 記錄元件對話框
# ============================================================
class CaptureDialog(tk.Toplevel):
    """彈出對話框：顯示元件資訊、讓使用者選 Action 與 value"""

    def __init__(self, parent, element_info: dict):
        super().__init__(parent)
        self.title("記錄 UI 元件")
        self.resizable(False, False)
        self.grab_set()
        self.transient(parent)

        self.result = None  # 儲存回傳結果

        pad = dict(padx=8, pady=4)

        # --- 元件資訊（唯讀）---
        info_frame = ttk.Labelframe(self, text="元件資訊", padding=8, bootstyle=INFO)
        info_frame.pack(fill="x", **pad)

        for i, (label, key) in enumerate(
            [("窗格名稱", "app"), ("Name", "Name"), ("ControlType", "ControlType"),
             ("AutomationId", "AutomationId"), ("ClassName", "ClassName")]
        ):
            ttk.Label(info_frame, text=f"{label}:").grid(
                row=i, column=0, sticky="w", padx=(0, 6), pady=2
            )
            entry = ttk.Entry(info_frame, width=50)
            entry.insert(0, element_info.get(key, ""))
            entry.configure(state="readonly")
            entry.grid(row=i, column=1, sticky="ew", pady=2)

        info_frame.columnconfigure(1, weight=1)

        # --- Action ---
        action_frame = ttk.Labelframe(self, text="Action", padding=8, bootstyle=PRIMARY)
        action_frame.pack(fill="x", **pad)

        self._action_var = tk.StringVar(value=ACTION_OPTIONS[0])
        action_combo = ttk.Combobox(
            action_frame,
            textvariable=self._action_var,
            values=ACTION_OPTIONS,
            state="readonly",
            width=50,
        )
        action_combo.pack(anchor="w", fill="x")

        # --- Value ---
        value_frame = ttk.Labelframe(self, text="Value（選填）", padding=8, bootstyle=PRIMARY)
        value_frame.pack(fill="x", **pad)

        self._value_entry = ttk.Entry(value_frame, width=50)
        self._value_entry.pack(fill="x")

        # --- Wait ---
        wait_frame = ttk.Labelframe(self, text="Wait 等待秒數（選填）", padding=8, bootstyle=PRIMARY)
        wait_frame.pack(fill="x", **pad)

        self._wait_entry = ttk.Entry(wait_frame, width=20)
        self._wait_entry.pack(anchor="w")

        # --- found_index ---
        fi_frame = ttk.Labelframe(self, text="found_index 同類型第 N 個（選填，從 0 開始）", padding=8, bootstyle=PRIMARY)
        fi_frame.pack(fill="x", **pad)

        self._fi_entry = ttk.Entry(fi_frame, width=20)
        self._fi_entry.pack(anchor="w")

        # --- Handler ---
        handler_frame = ttk.Labelframe(self, text="Handler（選填）", padding=8, bootstyle=PRIMARY)
        handler_frame.pack(fill="x", **pad)

        self._handler_var = tk.StringVar(value="")
        handler_combo = ttk.Combobox(
            handler_frame,
            textvariable=self._handler_var,
            values=HANDLER_OPTIONS,
            state="readonly",
            width=50,
        )
        handler_combo.pack(anchor="w", fill="x")

        # --- 按鈕 ---
        btn_frame = ttk.Frame(self, padding=8)
        btn_frame.pack(fill="x")

        ttk.Button(btn_frame, text="確定", command=self._on_ok, width=10, bootstyle=SUCCESS).pack(
            side="right", padx=4
        )
        ttk.Button(btn_frame, text="取消", command=self._on_cancel, width=10, bootstyle=SECONDARY).pack(
            side="right", padx=4
        )

        # 置中顯示
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = parent.winfo_x() + (parent.winfo_width() - w) // 2
        y = parent.winfo_y() + (parent.winfo_height() - h) // 2
        self.geometry(f"+{x}+{y}")

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _on_ok(self):
        selected = self._action_var.get()
        self.result = {
            "Action": ACTION_VALUE_MAP.get(selected, selected),
            "value": self._value_entry.get(),
        }
        wait_str = self._wait_entry.get().strip()
        if wait_str:
            try:
                self.result["Wait"] = float(wait_str)
            except ValueError:
                pass
        fi_str = self._fi_entry.get().strip()
        if fi_str:
            try:
                self.result["found_index"] = int(fi_str)
            except ValueError:
                pass
        handler_str = self._handler_var.get().strip()
        if handler_str:
            self.result["handler"] = handler_str
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


# ============================================================
# 修改元件對話框
# ============================================================
# 反向查找：儲存值 → 顯示文字
_VALUE_TO_DISPLAY = {v: k for k, v in ACTION_VALUE_MAP.items()}


class EditDialog(tk.Toplevel):
    """彈出對話框：修改已記錄元件的所有欄位"""

    def __init__(self, parent, element_info: dict):
        super().__init__(parent)
        self.title("修改 UI 元件")
        self.resizable(False, False)
        self.grab_set()
        self.transient(parent)

        self.result = None

        pad = dict(padx=8, pady=4)

        # --- 元件資訊（可編輯）---
        info_frame = ttk.Labelframe(self, text="元件資訊", padding=8, bootstyle=INFO)
        info_frame.pack(fill="x", **pad)

        self._entries = {}
        for i, (label, key) in enumerate(
            [("窗格名稱", "app"), ("Name", "Name"), ("ControlType", "ControlType"),
             ("AutomationId", "AutomationId"), ("ClassName", "ClassName")]
        ):
            ttk.Label(info_frame, text=f"{label}:").grid(
                row=i, column=0, sticky="w", padx=(0, 6), pady=2
            )
            entry = ttk.Entry(info_frame, width=50)
            entry.insert(0, element_info.get(key, ""))
            entry.grid(row=i, column=1, sticky="ew", pady=2)
            self._entries[key] = entry

        info_frame.columnconfigure(1, weight=1)

        # --- Action ---
        action_frame = ttk.Labelframe(self, text="Action", padding=8, bootstyle=PRIMARY)
        action_frame.pack(fill="x", **pad)

        current_action = element_info.get("Action", "")
        current_display = _VALUE_TO_DISPLAY.get(current_action, current_action)
        default_display = (
            current_display if current_display in ACTION_OPTIONS else ACTION_OPTIONS[0]
        )

        self._action_var = tk.StringVar(value=default_display)
        action_combo = ttk.Combobox(
            action_frame,
            textvariable=self._action_var,
            values=ACTION_OPTIONS,
            state="readonly",
            width=50,
        )
        action_combo.pack(anchor="w", fill="x")

        # --- Value ---
        value_frame = ttk.Labelframe(self, text="Value（選填）", padding=8, bootstyle=PRIMARY)
        value_frame.pack(fill="x", **pad)

        self._value_entry = ttk.Entry(value_frame, width=50)
        self._value_entry.insert(0, element_info.get("value", ""))
        self._value_entry.pack(fill="x")

        # --- Wait ---
        wait_frame = ttk.Labelframe(self, text="Wait 等待秒數（選填）", padding=8, bootstyle=PRIMARY)
        wait_frame.pack(fill="x", **pad)

        self._wait_entry = ttk.Entry(wait_frame, width=20)
        wait_val = element_info.get("Wait", "")
        if wait_val != "":
            self._wait_entry.insert(0, str(wait_val))
        self._wait_entry.pack(anchor="w")

        # --- found_index ---
        fi_frame = ttk.Labelframe(self, text="found_index 同類型第 N 個（選填，從 0 開始）", padding=8, bootstyle=PRIMARY)
        fi_frame.pack(fill="x", **pad)

        self._fi_entry = ttk.Entry(fi_frame, width=20)
        fi_val = element_info.get("found_index", "")
        if fi_val != "":
            self._fi_entry.insert(0, str(fi_val))
        self._fi_entry.pack(anchor="w")

        # --- Handler ---
        handler_frame = ttk.Labelframe(self, text="Handler（選填）", padding=8, bootstyle=PRIMARY)
        handler_frame.pack(fill="x", **pad)

        current_handler = element_info.get("handler", "")
        if current_handler not in HANDLER_OPTIONS:
            current_handler = ""
        self._handler_var = tk.StringVar(value=current_handler)
        handler_combo = ttk.Combobox(
            handler_frame,
            textvariable=self._handler_var,
            values=HANDLER_OPTIONS,
            state="readonly",
            width=50,
        )
        handler_combo.pack(anchor="w", fill="x")

        # --- 按鈕 ---
        btn_frame = ttk.Frame(self, padding=8)
        btn_frame.pack(fill="x")

        ttk.Button(btn_frame, text="確定", command=self._on_ok, width=10, bootstyle=SUCCESS).pack(
            side="right", padx=4
        )
        ttk.Button(btn_frame, text="取消", command=self._on_cancel, width=10, bootstyle=SECONDARY).pack(
            side="right", padx=4
        )

        # 置中顯示
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = parent.winfo_x() + (parent.winfo_width() - w) // 2
        y = parent.winfo_y() + (parent.winfo_height() - h) // 2
        self.geometry(f"+{x}+{y}")

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _on_ok(self):
        selected = self._action_var.get()
        result = {
            "app": self._entries["app"].get(),
            "Name": self._entries["Name"].get(),
            "ControlType": self._entries["ControlType"].get(),
            "AutomationId": self._entries["AutomationId"].get(),
            "ClassName": self._entries["ClassName"].get(),
            "Action": ACTION_VALUE_MAP.get(selected, selected),
            "value": self._value_entry.get(),
        }
        wait_str = self._wait_entry.get().strip()
        if wait_str:
            try:
                result["Wait"] = float(wait_str)
            except ValueError:
                pass
        fi_str = self._fi_entry.get().strip()
        if fi_str:
            try:
                result["found_index"] = int(fi_str)
            except ValueError:
                pass
        handler_str = self._handler_var.get().strip()
        if handler_str:
            result["handler"] = handler_str
        self.result = result
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


# ============================================================
# 主要 GUI 應用程式
# ============================================================
class InspectApp:
    def __init__(self, root):
        self.root = root

        self._inspector = UIAInspector()
        self._is_inspecting = False
        self._poll_id = None
        self._hotkey_listener: HotkeyListener | None = None
        self._captured: list[dict] = []  # 已記錄的元件
        self._current_info: dict | None = None  # 目前滑鼠下方的元件資訊
        self._runner_process = None  # 執行中的子程序

        self._build_gui()

    # ========================================================
    # GUI 建構
    # ========================================================
    def _build_gui(self):
        # --- 分頁 Notebook ---
        self._notebook = ttk.Notebook(self.root)
        self._notebook.pack(fill="both", expand=True)

        # --- 分頁 1: 偵測工具 ---
        inspect_frame = ttk.Frame(self._notebook)
        self._notebook.add(inspect_frame, text=" 偵測工具 ")
        self._build_inspect_tab(inspect_frame)

        # --- 分頁 2: 執行工具 ---
        runner_frame = ttk.Frame(self._notebook)
        self._notebook.add(runner_frame, text=" 執行工具 ")
        self._build_runner_tab(runner_frame)

    # ========================================================
    # 偵測工具分頁
    # ========================================================
    def _build_inspect_tab(self, parent):
        """建構「偵測工具」分頁"""
        # --- 頂部按鈕列 ---
        btn_bar = ttk.Frame(parent, padding=6)
        btn_bar.pack(fill="x")

        self._btn_start = ttk.Button(
            btn_bar, text="開始偵測", command=self._start_inspect, bootstyle=SUCCESS
        )
        self._btn_start.pack(side="left", padx=4)

        self._btn_stop = ttk.Button(
            btn_bar, text="結束偵測", command=self._stop_inspect, state="disabled", bootstyle=DANGER
        )
        self._btn_stop.pack(side="left", padx=4)

        ttk.Separator(btn_bar, orient="vertical").pack(side="left", fill="y", padx=6)

        ttk.Button(
            btn_bar, text="載入設定檔", command=self._load_config, bootstyle=(INFO, OUTLINE)
        ).pack(side="left", padx=4)

        ttk.Label(btn_bar, text="快捷鍵: Ctrl+F1 記錄 | Ctrl+F2 詳細資訊", bootstyle=SECONDARY).pack(
            side="right", padx=8
        )

        ttk.Separator(parent, orient="horizontal").pack(fill="x")

        # --- 即時元件資訊區 ---
        info_frame = ttk.Labelframe(
            parent, text="當前滑鼠下方元件", padding=10, bootstyle=INFO
        )
        info_frame.pack(fill="x", padx=8, pady=(8, 4))

        self._var_app = tk.StringVar(value="—")
        self._var_name = tk.StringVar(value="—")
        self._var_type = tk.StringVar(value="—")
        self._var_auto_id = tk.StringVar(value="—")
        self._var_class_name = tk.StringVar(value="—")

        for i, (label, var) in enumerate(
            [
                ("窗格名稱", self._var_app),
                ("Name", self._var_name),
                ("ControlType", self._var_type),
                ("AutomationId", self._var_auto_id),
                ("ClassName", self._var_class_name),
            ]
        ):
            ttk.Label(info_frame, text=f"{label}:", width=12, anchor="w").grid(
                row=i, column=0, sticky="w", pady=2
            )
            ttk.Label(
                info_frame,
                textvariable=var,
                relief="sunken",
                padding=(6, 2),
                anchor="w",
            ).grid(row=i, column=1, sticky="ew", pady=2)

        info_frame.columnconfigure(1, weight=1)

        # 「更多資訊」按鈕
        self._btn_detail = ttk.Button(
            info_frame, text="更多資訊", command=self._show_detail,
            state="disabled", bootstyle=(INFO, OUTLINE), width=10,
        )
        self._btn_detail.grid(row=5, column=1, sticky="e", pady=(4, 0))

        ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=4)

        # --- 已記錄元件列表 ---
        list_frame = ttk.Labelframe(
            parent, text="已記錄的 UI 元件", padding=6, bootstyle=PRIMARY
        )
        list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 4))

        columns = ("step", "app", "name", "type", "auto_id", "class_name", "action", "value", "handler")
        self._tree = ttk.Treeview(
            list_frame, columns=columns, show="headings", selectmode="browse"
        )

        col_cfg = [
            ("step", "步驟", 45),
            ("app", "窗格名稱", 140),
            ("name", "Name", 140),
            ("type", "ControlType", 140),
            ("auto_id", "AutomationId", 120),
            ("class_name", "ClassName", 100),
            ("action", "Action", 110),
            ("value", "value", 90),
            ("handler", "Handler", 100),
        ]
        for col_id, heading, width in col_cfg:
            self._tree.heading(col_id, text=heading, anchor="w")
            self._tree.column(col_id, width=width, minwidth=40, anchor="w")

        scrollbar = ttk.Scrollbar(
            list_frame, orient="vertical", command=self._tree.yview
        )
        self._tree.configure(yscrollcommand=scrollbar.set)
        self._tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 右鍵選單（修改 / 刪除）
        self._ctx_menu = tk.Menu(self._tree, tearoff=0)
        self._ctx_menu.add_command(label="修改", command=self._edit_selected)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="刪除", command=self._delete_selected)
        self._tree.bind("<Button-3>", self._show_context_menu)
        self._tree.bind("<Double-1>", lambda e: self._edit_selected())

        # 拖曳排序（左鍵長按拖曳以調換順序）
        self._drag_source = None
        self._drag_target_prev = None
        self._tree.tag_configure("drag_target", background="#2d5a8e")
        self._tree.bind("<ButtonPress-1>", self._drag_start)
        self._tree.bind("<B1-Motion>", self._drag_motion)
        self._tree.bind("<ButtonRelease-1>", self._drag_drop)

        # --- 操作按鈕區 ---
        btn_frame = ttk.Frame(parent, padding=6)
        btn_frame.pack(fill="x", padx=8, pady=4)

        ttk.Button(btn_frame, text="全部清除", command=self._clear_all, bootstyle=(DANGER, OUTLINE)).pack(
            side="right", padx=4
        )

        ttk.Button(btn_frame, text="儲存設定檔", command=self._export_json, bootstyle=(SUCCESS, OUTLINE)).pack(
            side="right", padx=4
        )

        # --- 底部狀態列 ---
        self._status_var = tk.StringVar(value="就緒")
        status_bar = ttk.Label(
            parent,
            textvariable=self._status_var,
            relief="sunken",
            anchor="w",
            padding=(8, 2),
        )
        status_bar.pack(fill="x", side="bottom")

    # ========================================================
    # 執行工具分頁
    # ========================================================
    def _build_runner_tab(self, parent):
        """建構「執行工具」分頁"""
        # --- 設定檔選擇區 ---
        config_frame = ttk.Labelframe(parent, text="設定檔", padding=8, bootstyle=INFO)
        config_frame.pack(fill="x", padx=8, pady=(8, 4))

        ttk.Label(config_frame, text="設定檔路徑:").pack(side="left", padx=(0, 4))

        self._runner_config_var = tk.StringVar()
        config_entry = ttk.Entry(
            config_frame, textvariable=self._runner_config_var, width=50
        )
        config_entry.pack(side="left", fill="x", expand=True, padx=4)

        ttk.Button(
            config_frame, text="瀏覽...", command=self._browse_runner_config, bootstyle=(INFO, OUTLINE)
        ).pack(side="left", padx=4)

        # --- 執行控制區 ---
        exec_frame = ttk.Frame(parent, padding=6)
        exec_frame.pack(fill="x", padx=8)

        self._btn_execute = ttk.Button(
            exec_frame, text="▶ 執行自動化", command=self._run_automation, bootstyle=SUCCESS
        )
        self._btn_execute.pack(side="left", padx=4)

        self._btn_stop_runner = ttk.Button(
            exec_frame, text="■ 中斷執行", command=self._stop_automation,
            state="disabled", bootstyle=DANGER
        )
        self._btn_stop_runner.pack(side="left", padx=4)

        self._runner_status_var = tk.StringVar(value="就緒")
        ttk.Label(
            exec_frame, textvariable=self._runner_status_var, bootstyle=SECONDARY
        ).pack(side="left", padx=8)

        ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=4)

        # --- Log 輸出區 ---
        log_frame = ttk.Labelframe(parent, text="執行記錄", padding=6, bootstyle=PRIMARY)
        log_frame.pack(fill="both", expand=True, padx=8, pady=(0, 4))

        self._log_text = tk.Text(
            log_frame, wrap="word", state="disabled",
            font=("Consolas", 9), bg="#1e1e1e", fg="#cccccc",
            insertbackground="#cccccc",
        )
        log_scrollbar = ttk.Scrollbar(
            log_frame, orient="vertical", command=self._log_text.yview
        )
        self._log_text.configure(yscrollcommand=log_scrollbar.set)
        self._log_text.pack(side="left", fill="both", expand=True)
        log_scrollbar.pack(side="right", fill="y")

        # --- 底部按鈕區 ---
        bottom_frame = ttk.Frame(parent, padding=6)
        bottom_frame.pack(fill="x", padx=8, pady=4)

        ttk.Button(
            bottom_frame, text="清除記錄", command=self._clear_log, bootstyle=(DANGER, OUTLINE)
        ).pack(side="right", padx=4)

    # ========================================================
    # 開始 / 結束偵測
    # ========================================================
    def _start_inspect(self):
        if self._is_inspecting:
            return

        self._is_inspecting = True
        self._btn_start.configure(state="disabled")
        self._btn_stop.configure(state="normal")

        # 視窗置頂
        self.root.attributes("-topmost", True)

        # 啟動快捷鍵監聽
        self._hotkey_listener = HotkeyListener(self._on_hotkey, self._on_hotkey_detail)

        # 開始輪詢
        self._status_var.set("偵測中… 移動滑鼠到目標元件，Ctrl+F1 記錄 | Ctrl+F2 查看詳細")
        self._poll_mouse()

    def _stop_inspect(self):
        if not self._is_inspecting:
            return

        self._is_inspecting = False
        self._btn_start.configure(state="normal")
        self._btn_stop.configure(state="disabled")
        self.root.attributes("-topmost", False)

        # 停止輪詢
        if self._poll_id is not None:
            self.root.after_cancel(self._poll_id)
            self._poll_id = None

        # 停止快捷鍵
        if self._hotkey_listener:
            self._hotkey_listener.stop()
            self._hotkey_listener = None

        # 清空即時顯示
        self._var_app.set("—")
        self._var_name.set("—")
        self._var_type.set("—")
        self._var_auto_id.set("—")
        self._var_class_name.set("—")
        self._btn_detail.configure(state="disabled")

        self._status_var.set("偵測已結束")

    # ========================================================
    # 滑鼠輪詢
    # ========================================================
    def _poll_mouse(self):
        if not self._is_inspecting:
            return

        try:
            pt = ctypes.wintypes.POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
            info = self._inspector.get_element_info(pt.x, pt.y)

            if info:
                self._current_info = info
                self._var_app.set(info["app"])
                self._var_name.set(info["Name"])
                self._var_type.set(info["ControlType"])
                self._var_auto_id.set(info.get("AutomationId", ""))
                self._var_class_name.set(info.get("ClassName", ""))
                self._btn_detail.configure(state="normal")
            else:
                self._current_info = None
                self._var_app.set("（無法偵測）")
                self._var_name.set("—")
                self._var_type.set("—")
                self._var_auto_id.set("—")
                self._var_class_name.set("—")
                self._btn_detail.configure(state="disabled")
        except Exception:
            pass

        self._poll_id = self.root.after(POLL_INTERVAL_MS, self._poll_mouse)

    # ========================================================
    # 更多資訊
    # ========================================================
    def _show_detail(self):
        info = self._current_info
        if not info or "_elem" not in info:
            return
        # 暂停偵測輪詢，避免元素變動
        was_inspecting = self._is_inspecting
        if self._poll_id is not None:
            self.root.after_cancel(self._poll_id)
            self._poll_id = None
        try:
            detail_text = self._inspector.get_detailed_info(info["_elem"])
        except Exception as e:
            detail_text = f"無法取得詳細資訊：{e}"
        dlg = DetailDialog(self.root, detail_text)
        self.root.wait_window(dlg)
        # 恢復偵測輪詢
        if was_inspecting and self._is_inspecting:
            self._poll_mouse()

    # ========================================================
    # 快捷鍵回呼 (Ctrl+F1) — 從背景執行緒安全呼叫
    # ========================================================
    def _on_hotkey(self):
        self.root.after(0, self._capture_element)

    def _on_hotkey_detail(self):
        self.root.after(0, self._show_detail)

    def _capture_element(self):
        if not self._is_inspecting or self._current_info is None:
            return

        info = dict(self._current_info)  # 複製一份

        dlg = CaptureDialog(self.root, info)
        self.root.wait_window(dlg)

        if dlg.result is None:
            return  # 使用者取消

        info["Action"] = dlg.result["Action"]
        info["value"] = dlg.result["value"]
        if "Wait" in dlg.result:
            info["Wait"] = dlg.result["Wait"]
        if "found_index" in dlg.result:
            info["found_index"] = dlg.result["found_index"]
        if "handler" in dlg.result:
            info["handler"] = dlg.result["handler"]
        # 移除內部用的 _elem 參考，不需存入記錄
        info.pop("_elem", None)
        self._captured.append(info)

        step_num = len(self._captured)
        self._tree.insert(
            "",
            "end",
            iid=str(step_num),
            values=(
                step_num,
                info["app"],
                info["Name"],
                info["ControlType"],
                info.get("AutomationId", ""),
                info.get("ClassName", ""),
                info["Action"],
                info["value"],
                info.get("handler", ""),
            ),
        )
        self._status_var.set(f"已記錄第 {step_num} 個元件")

    # ========================================================
    # 右鍵選單：刪除
    # ========================================================
    def _show_context_menu(self, event):
        item = self._tree.identify_row(event.y)
        if item:
            self._tree.selection_set(item)
            self._ctx_menu.tk_popup(event.x_root, event.y_root)

    def _edit_selected(self):
        selected = self._tree.selection()
        if not selected:
            return

        item_id = selected[0]
        idx = int(item_id) - 1
        if not (0 <= idx < len(self._captured)):
            return

        info = dict(self._captured[idx])
        dlg = EditDialog(self.root, info)
        self.root.wait_window(dlg)

        if dlg.result is None:
            return

        self._captured[idx] = dlg.result
        self._rebuild_tree()
        self._status_var.set(f"已修改第 {idx + 1} 個元件")

    def _delete_selected(self):
        selected = self._tree.selection()
        if not selected:
            return

        item_id = selected[0]
        idx = int(item_id) - 1
        if 0 <= idx < len(self._captured):
            self._captured.pop(idx)

        # 重建 Treeview（因為步驟編號需重新排列）
        self._rebuild_tree()
        self._status_var.set(f"已刪除，目前共 {len(self._captured)} 個元件")

    def _rebuild_tree(self):
        self._tree.delete(*self._tree.get_children())
        for i, info in enumerate(self._captured, start=1):
            self._tree.insert(
                "",
                "end",
                iid=str(i),
                values=(
                    i,
                    info["app"],
                    info["Name"],
                    info["ControlType"],
                    info.get("AutomationId", ""),
                    info.get("ClassName", ""),
                    info["Action"],
                    info["value"],
                    info.get("handler", ""),
                ),
            )

    # ========================================================
    # 拖曳排序
    # ========================================================
    def _drag_start(self, event):
        """記錄被拖曳的列"""
        item = self._tree.identify_row(event.y)
        if item:
            self._drag_source = item
            self._drag_target_prev = None
            self._tree.configure(cursor="size_ns")

    def _drag_motion(self, event):
        """拖曳中：高亮目標列"""
        if not self._drag_source:
            return
        target = self._tree.identify_row(event.y)
        # 取消前一個目標的高亮
        if self._drag_target_prev and self._drag_target_prev in self._tree.get_children():
            self._tree.item(self._drag_target_prev, tags=())
        # 高亮新目標列
        if target and target != self._drag_source:
            self._tree.item(target, tags=("drag_target",))
            self._drag_target_prev = target
        else:
            self._drag_target_prev = None

    def _drag_drop(self, event):
        """放開：執行重排"""
        self._tree.configure(cursor="")
        # 清除高亮
        if self._drag_target_prev and self._drag_target_prev in self._tree.get_children():
            self._tree.item(self._drag_target_prev, tags=())
        self._drag_target_prev = None

        src = self._drag_source
        self._drag_source = None
        if not src:
            return

        target = self._tree.identify_row(event.y)
        if not target or target == src:
            return

        src_idx = int(src) - 1
        dst_idx = int(target) - 1
        if not (0 <= src_idx < len(self._captured) and 0 <= dst_idx < len(self._captured)):
            return

        item = self._captured.pop(src_idx)
        self._captured.insert(dst_idx, item)
        self._rebuild_tree()
        # 重新選中被移動的列
        self._tree.selection_set(str(dst_idx + 1))
        self._tree.see(str(dst_idx + 1))
        self._status_var.set(f"已將第 {src_idx + 1} 步移到第 {dst_idx + 1} 步")

    def _clear_all(self):
        """清除所有已記錄的 UI 元件"""
        if not self._captured:
            messagebox.showinfo("提示", "目前沒有已記錄的元件")
            return

        result = messagebox.askyesno(
            "確認清除",
            f"確定要清除所有 {len(self._captured)} 個已記錄的 UI 元件嗎？",
        )
        if result:
            self._captured.clear()
            self._rebuild_tree()
            self._status_var.set("已清除所有元件")

    # ========================================================
    # 載入設定檔
    # ========================================================
    def _load_config(self):
        """從 JSON 設定檔載入 UI 元件列表"""
        if getattr(sys, 'frozen', False):
            default_dir = os.path.dirname(sys.executable)
        else:
            default_dir = os.path.dirname(os.path.abspath(__file__))
        filepath = filedialog.askopenfilename(
            title="選擇設定檔",
            initialdir=default_dir,
            defaultextension=".json",
            filetypes=[("JSON 檔案", "*.json"), ("所有檔案", "*.*")],
        )
        if not filepath:
            return

        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("讀取失敗", f"無法讀取設定檔：\n{e}")
            return

        # 解析 JSON（格式：{"1": {...}, "2": {...}}）
        items = []
        try:
            for key in sorted(data.keys(), key=lambda k: int(k)):
                entry = data[key]
                item = {
                    "app": entry.get("app", ""),
                    "Name": entry.get("Name", ""),
                    "ControlType": entry.get("ControlType", ""),
                    "AutomationId": entry.get("AutomationId", ""),
                    "ClassName": entry.get("ClassName", ""),
                    "Action": entry.get("Action", ""),
                    "value": entry.get("value", ""),
                }
                if "Wait" in entry:
                    item["Wait"] = entry["Wait"]
                if "found_index" in entry:
                    item["found_index"] = entry["found_index"]
                if "handler" in entry:
                    item["handler"] = entry["handler"]
                items.append(item)
        except Exception as e:
            messagebox.showerror("格式錯誤", f"設定檔格式不符：\n{e}")
            return

        if not items:
            messagebox.showinfo("提示", "設定檔中沒有任何元件")
            return

        # 若已有資料，詢問取代或追加
        if self._captured:
            answer = messagebox.askyesnocancel(
                "載入方式",
                f"目前已有 {len(self._captured)} 個元件。\n\n"
                f"「是」= 取代全部\n「否」= 追加到末尾\n「取消」= 放棄載入",
            )
            if answer is None:
                return
            if answer:  # 取代
                self._captured.clear()

        self._captured.extend(items)
        self._rebuild_tree()
        self._status_var.set(
            f"已載入 {len(items)} 個元件 ← {os.path.basename(filepath)}"
        )

    # ========================================================
    # 匯出 JSON
    # ========================================================
    def _export_json(self):
        if getattr(sys, 'frozen', False):
            default_dir = os.path.dirname(sys.executable)
        else:
            default_dir = os.path.dirname(os.path.abspath(__file__))
        filepath = filedialog.asksaveasfilename(
            title="匯出 set.json",
            initialdir=default_dir,
            initialfile="set.json",
            defaultextension=".json",
            filetypes=[("JSON 檔案", "*.json"), ("所有檔案", "*.*")],
        )
        if not filepath:
            self._status_var.set("已取消匯出")
            return

        data = {}
        for i, info in enumerate(self._captured, start=1):
            entry = {
                "app": info["app"],
                "Name": info["Name"],
                "ControlType": info["ControlType"],
                "AutomationId": info.get("AutomationId", ""),
                "ClassName": info.get("ClassName", ""),
                "Action": info["Action"],
                "value": info["value"],
            }
            if "Wait" in info:
                entry["Wait"] = info["Wait"]
            if "found_index" in info:
                entry["found_index"] = info["found_index"]
            if "handler" in info:
                entry["handler"] = info["handler"]
            data[str(i)] = entry

        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

        self._status_var.set(f"已匯出: {filepath}")
        messagebox.showinfo("匯出成功", f"已成功匯出 {len(self._captured)} 個元件至:\n{filepath}")

    # ========================================================
    # 執行工具：瀏覽 / 執行 / Log
    # ========================================================
    def _browse_runner_config(self):
        """瀏覽並選擇設定檔"""
        if getattr(sys, 'frozen', False):
            default_dir = os.path.dirname(sys.executable)
        else:
            default_dir = os.path.dirname(os.path.abspath(__file__))
        filepath = filedialog.askopenfilename(
            title="選擇設定檔",
            initialdir=default_dir,
            defaultextension=".json",
            filetypes=[("JSON 檔案", "*.json"), ("所有檔案", "*.*")],
        )
        if filepath:
            self._runner_config_var.set(filepath)

    def _run_automation(self):
        """在背景執行 run.py"""
        config_path = self._runner_config_var.get().strip()
        if not config_path:
            messagebox.showwarning("提示", "請先選擇設定檔")
            return
        if not os.path.isfile(config_path):
            messagebox.showerror("錯誤", f"找不到設定檔:\n{config_path}")
            return

        # 鎖定執行按鈕，啟用中斷按鈕
        self._btn_execute.configure(state="disabled")
        self._btn_stop_runner.configure(state="normal")
        self._runner_status_var.set("執行中...")

        # 清空 Log
        self._log_text.configure(state="normal")
        self._log_text.delete("1.0", "end")
        self._log_text.configure(state="disabled")

        # 啟動子程序
        # 傳遞 PYTHONIOENCODING=utf-8 給子程序，避免 cp950 編碼錯誤
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        if getattr(sys, 'frozen', False):
            # 打包後：使用同目錄下的 run.exe
            base_dir = os.path.dirname(sys.executable)
            run_exe = os.path.join(base_dir, "run.exe")
            cmd = [run_exe, "--case", config_path]
        else:
            run_py = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "run.py"
            )
            cmd = [sys.executable, run_py, "--case", config_path]
        self._runner_process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            cwd=os.path.dirname(os.path.abspath(__file__)),
            env=env,
        )

        # 背景執行緒讀取輸出
        thread = threading.Thread(target=self._read_process_output, daemon=True)
        thread.start()

    def _read_process_output(self):
        """背景執行緒：逐行讀取子程序輸出"""
        try:
            for line in self._runner_process.stdout:
                self.root.after(0, self._append_log, line)
            self._runner_process.wait()
        except Exception as e:
            self.root.after(0, self._append_log, f"\n錯誤: {e}\n")
        returncode = self._runner_process.returncode
        self.root.after(0, self._on_process_done, returncode)

    def _append_log(self, text):
        """將文字附加到 Log 區域"""
        self._log_text.configure(state="normal")
        self._log_text.insert("end", text)
        self._log_text.see("end")
        self._log_text.configure(state="disabled")

    def _stop_automation(self):
        """中斷執行中的自動化程序"""
        if self._runner_process is not None:
            self._runner_process.terminate()
            self._btn_stop_runner.configure(state="disabled")
            self._runner_status_var.set("正在中斷...")

    def _on_process_done(self, returncode):
        """子程序結束時的回呼"""
        self._btn_execute.configure(state="normal")
        self._btn_stop_runner.configure(state="disabled")
        self._runner_process = None

        if returncode == 0:
            self._runner_status_var.set("執行完成")
            self._append_log("\n===== 執行完成 =====\n")
            messagebox.showinfo("完成", "自動化流程已執行完成！")
        else:
            self._runner_status_var.set(f"執行失敗 (代碼: {returncode})")
            self._append_log(f"\n===== 執行失敗 (代碼: {returncode}) =====\n")
            messagebox.showerror(
                "錯誤",
                f"自動化流程執行失敗！\n返回代碼: {returncode}\n\n請查看執行記錄了解詳情。",
            )

    def _clear_log(self):
        """清除 Log 輸出區"""
        self._log_text.configure(state="normal")
        self._log_text.delete("1.0", "end")
        self._log_text.configure(state="disabled")

    # ========================================================
    # 視窗關閉
    # ========================================================
    def on_close(self):
        if self._is_inspecting:
            self._is_inspecting = False
            if self._poll_id is not None:
                self.root.after_cancel(self._poll_id)
            if self._hotkey_listener:
                self._hotkey_listener.stop()
        if self._runner_process is not None:
            self._runner_process.terminate()
        self.root.destroy()


# ============================================================
# 程式進入點
# ============================================================
def main():
    root = ttk.Window(themename="darkly", title="WinAutomation", size=(1200, 900), minsize=(1200, 900))
    app = InspectApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
