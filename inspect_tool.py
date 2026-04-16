"""
UI 元素偵測工具 (inspect_tool.py)
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
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# COM (隨 pywinauto 一起安裝)
import comtypes
import comtypes.client


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
        """回傳 {app, Name, ControlType} 或 None"""
        try:
            pt = self._tagPOINT(x, y)
            elem = self._uia.ElementFromPoint(pt)
            if elem is None:
                return None

            name = elem.CurrentName or ""
            control_type_id = elem.CurrentControlType
            control_type_str = format_control_type(control_type_id)
            app_name = self._get_top_window_name(elem)

            return {
                "app": app_name,
                "Name": name,
                "ControlType": control_type_str,
            }
        except Exception:
            return None

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
    """使用 Windows RegisterHotKey API 監聽 Ctrl+F1"""

    _MOD_CONTROL = 0x0002
    _VK_F1 = 0x70
    _WM_HOTKEY = 0x0312
    _HOTKEY_ID = 1

    def __init__(self, callback):
        self._callback = callback
        self._running = True
        self._thread = threading.Thread(target=self._listen, daemon=True)
        self._thread.start()

    def _listen(self):
        user32 = ctypes.windll.user32
        ok = user32.RegisterHotKey(
            None, self._HOTKEY_ID, self._MOD_CONTROL, self._VK_F1
        )
        if not ok:
            return

        msg = ctypes.wintypes.MSG()
        while self._running:
            ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if ret <= 0:
                break
            if msg.message == self._WM_HOTKEY and msg.wParam == self._HOTKEY_ID:
                self._callback()

        user32.UnregisterHotKey(None, self._HOTKEY_ID)

    def stop(self):
        self._running = False
        # 發送 WM_QUIT 讓 GetMessageW 返回
        if self._thread.is_alive():
            ctypes.windll.user32.PostThreadMessageW(
                self._thread.native_id, 0x0012, 0, 0  # WM_QUIT
            )


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
        info_frame = ttk.LabelFrame(self, text="元件資訊", padding=8)
        info_frame.pack(fill="x", **pad)

        for i, (label, key) in enumerate(
            [("窗格名稱", "app"), ("Name", "Name"), ("ControlType", "ControlType")]
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
        action_frame = ttk.LabelFrame(self, text="Action", padding=8)
        action_frame.pack(fill="x", **pad)

        self._action_var = tk.StringVar(value=ACTION_OPTIONS[0])
        action_combo = ttk.Combobox(
            action_frame,
            textvariable=self._action_var,
            values=ACTION_OPTIONS,
            state="readonly",
            width=30,
        )
        action_combo.pack(anchor="w")

        # --- Value ---
        value_frame = ttk.LabelFrame(self, text="Value（選填）", padding=8)
        value_frame.pack(fill="x", **pad)

        self._value_entry = ttk.Entry(value_frame, width=50)
        self._value_entry.pack(fill="x")

        # --- 按鈕 ---
        btn_frame = ttk.Frame(self, padding=8)
        btn_frame.pack(fill="x")

        ttk.Button(btn_frame, text="確定", command=self._on_ok, width=10).pack(
            side="right", padx=4
        )
        ttk.Button(btn_frame, text="取消", command=self._on_cancel, width=10).pack(
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
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


# ============================================================
# 主要 GUI 應用程式
# ============================================================
class InspectApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("UI 元素偵測工具")
        self.root.geometry("780x560")
        self.root.minsize(640, 440)

        self._inspector = UIAInspector()
        self._is_inspecting = False
        self._poll_id = None
        self._hotkey_listener: HotkeyListener | None = None
        self._captured: list[dict] = []  # 已記錄的元件
        self._current_info: dict | None = None  # 目前滑鼠下方的元件資訊

        self._build_gui()

    # ========================================================
    # GUI 建構
    # ========================================================
    def _build_gui(self):
        # --- 頂部按鈕列 ---
        btn_bar = ttk.Frame(self.root, padding=6)
        btn_bar.pack(fill="x")

        self._btn_start = ttk.Button(
            btn_bar, text="開始偵測", command=self._start_inspect
        )
        self._btn_start.pack(side="left", padx=4)

        self._btn_stop = ttk.Button(
            btn_bar, text="結束偵測", command=self._stop_inspect, state="disabled"
        )
        self._btn_stop.pack(side="left", padx=4)

        ttk.Label(btn_bar, text="快捷鍵: Ctrl+F1 記錄元件", foreground="gray").pack(
            side="right", padx=8
        )

        ttk.Separator(self.root, orient="horizontal").pack(fill="x")

        # --- 即時元件資訊區 ---
        info_frame = ttk.LabelFrame(
            self.root, text="當前滑鼠下方元件", padding=10
        )
        info_frame.pack(fill="x", padx=8, pady=(8, 4))

        self._var_app = tk.StringVar(value="—")
        self._var_name = tk.StringVar(value="—")
        self._var_type = tk.StringVar(value="—")

        for i, (label, var) in enumerate(
            [
                ("窗格名稱", self._var_app),
                ("Name", self._var_name),
                ("ControlType", self._var_type),
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

        ttk.Separator(self.root, orient="horizontal").pack(fill="x", pady=4)

        # --- 已記錄元件列表 ---
        list_frame = ttk.LabelFrame(
            self.root, text="已記錄的 UI 元件", padding=6
        )
        list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        columns = ("step", "app", "name", "type", "action", "value")
        self._tree = ttk.Treeview(
            list_frame, columns=columns, show="headings", selectmode="browse"
        )

        col_cfg = [
            ("step", "步驟", 45),
            ("app", "窗格名稱", 160),
            ("name", "Name", 160),
            ("type", "ControlType", 160),
            ("action", "Action", 120),
            ("value", "value", 100),
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

        # 右鍵選單（刪除）
        self._ctx_menu = tk.Menu(self._tree, tearoff=0)
        self._ctx_menu.add_command(label="刪除", command=self._delete_selected)
        self._tree.bind("<Button-3>", self._show_context_menu)

        # --- 操作按鈕區 ---
        btn_frame = ttk.Frame(self.root, padding=6)
        btn_frame.pack(fill="x", padx=8, pady=4)

        ttk.Button(btn_frame, text="全部清除", command=self._clear_all).pack(
            side="right", padx=4
        )

        # --- 底部狀態列 ---
        self._status_var = tk.StringVar(value="就緒")
        status_bar = ttk.Label(
            self.root,
            textvariable=self._status_var,
            relief="sunken",
            anchor="w",
            padding=(8, 2),
        )
        status_bar.pack(fill="x", side="bottom")

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
        self._hotkey_listener = HotkeyListener(self._on_hotkey)

        # 開始輪詢
        self._status_var.set("偵測中… 移動滑鼠到目標元件，Ctrl+F1 記錄")
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

        # 匯出 JSON
        if self._captured:
            self._export_json()
        else:
            self._status_var.set("未記錄任何元件，不產出檔案")

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
            else:
                self._current_info = None
                self._var_app.set("（無法偵測）")
                self._var_name.set("—")
                self._var_type.set("—")
        except Exception:
            pass

        self._poll_id = self.root.after(POLL_INTERVAL_MS, self._poll_mouse)

    # ========================================================
    # 快捷鍵回呼 (Ctrl+F1) — 從背景執行緒安全呼叫
    # ========================================================
    def _on_hotkey(self):
        self.root.after(0, self._capture_element)

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
                info["Action"],
                info["value"],
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
                    info["Action"],
                    info["value"],
                ),
            )

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
    # 匯出 JSON
    # ========================================================
    def _export_json(self):
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
            data[str(i)] = {
                "app": info["app"],
                "Name": info["Name"],
                "ControlType": info["ControlType"],
                "Action": info["Action"],
                "value": info["value"],
            }

        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

        self._status_var.set(f"已匯出: {filepath}")
        messagebox.showinfo("匯出成功", f"已成功匯出 {len(self._captured)} 個元件至:\n{filepath}")

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
        self.root.destroy()


# ============================================================
# 程式進入點
# ============================================================
def main():
    root = tk.Tk()
    app = InspectApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
