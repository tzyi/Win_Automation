# 功能

幫我寫一個可以抓取畫面中視窗元件UI資訊的小工具

像是inspect.exe可以抓取到UI元件的窗格名稱、Name、ControlType

最後可以產出set.json



- set.json

```json

{
    "1": {
        "app": "GeoTag Pro",
        "Name": "folder_open 選擇照片",
        "ControlType": "UIA_TextControlTypeId (0xC364)",
        "Action": "click_input()",
        "value": ""
    },
    "2": {
        "app": "開啟",
        "Name": "IMG_4301.JPG",
        "ControlType": "UIA_ListItemControlTypeId (0xC357)",
        "Action": "click_input()",
        "value": ""
    },
    "3": {
        "app": "開啟",
        "Name": "開啟(O)",
        "ControlType": "UIA_SplitButtonControlTypeId (0xC36F)",
        "Action": "click_input()",
        "value": ""
    },
    "4": {
        "app": "GeoTag Pro",
        "Name": "清除所有照片",
        "ControlType": "UIA_ButtonControlTypeId (0xC350)",
        "Action": "click_input()",
        "value": ""
    },
    "5": {
        "app": "WinDesk - 新分頁 - Google Chrome",
        "Name": "網址與搜尋列",
        "ControlType": "UIA_EditControlTypeId (0xC354)",
        "Action": "send_keys()",
        "value": "github{ENTER}"
    }
}
```


# 技術
- 請使用python的Tkinter來開發小工具的GUI界面



# 使用步驟

1. 小工具開啟後，使用者點選”開始偵測”按鈕，使用者就可以移動滑鼠鼠標到某元件上，小工具就會即刻同步顯示該UI元件的窗格名稱、Name、ControlType

2. 當使用者想記錄該UI元件，可以使用快捷鍵Ctrl+F1儲存該UI元件資訊，並可讓使用者選擇該UI元件的Action(參考下面的Action參考)以及value(如上面set.json)

3. 等所有UI元件資訊記錄完後，使用者點選”結束偵測”按鈕，就可以立刻產出set.json







# Action參考

| **方法** | **功能** | **備註** | 
|---|---|---|
| **click()** | 發送滑鼠點擊訊息 | 不要求控制項可見 | 
| **click_input()** | 使用滑鼠事件進行點擊 | 需要控制項在螢幕上可見，更貼近實際使用者操作 | 
| **double_click()** | 雙擊 | 基於訊息 | 
| **double_click_input()** | 雙擊輸入 | 基於滑鼠事件 | 
| **right_click()** | 右鍵點擊 | 基於訊息 | 
| **right_click_input()** | 右鍵點擊輸入 | 基於滑鼠事件 | 
| **type_keys(keys)** | 發送鍵盤輸入 | 例如：`type\_keys("{TAB}")`, `type\_keys("{ENTER}")` | 
| **send_message(msg)** | 發送 Windows 訊息 | 進階用法 | 
| **send_message_timeout()** | 發送訊息並等待回應 | 可設定逾時時間 | 
| **set_focus()** | 設定焦點 | \- | 
| **get_focus()** | 取得焦點 | \- | 
| **set_window_text(text)** | 設定視窗文字 | \- | 
| **close()** | 關閉視窗 | \- | 
| **close_click()** | 點擊關閉 | 執行額外延遲 | 
| **drag_mouse(dx, dy)** | 拖動滑鼠 | \- | 
| **move_mouse(x, y)** | 移動滑鼠 | \- | 
| **press_mouse()** | 按下滑鼠按鍵 | \- | 
| **release_mouse()** | 放開滑鼠按鍵 | \- | 
| **press_mouse_input()** | 按下滑鼠（基於事件） | \- | 
| **release_mouse_input()** | 放開滑鼠（基於事件） | \- | 
| **maximize()** | 最大化 | \- | 
| **minimize()** | 最小化 | \- | 
| **restore()** | 還原 | \- | 
| **get_show_state()** | 取得顯示狀態 | \- | 
| **menu_select(path)** | 選擇菜單項目 | \- | 
| **notify_menu_select()** | 通知菜單選擇 | \- | 
| **notify_parent()** | 通知父視窗 | \- | 
| **draw_outline()** | 繪製輪廓 | 用於視覺化 | 
| **move_window(x, y, width, height)** | 移動與調整大小 | \- | 
| **rectangle** | 取得矩形座標 | 屬性：.left, .top, .right, .bottom | 
| **children** | 取得子視窗列表 | \- | 

