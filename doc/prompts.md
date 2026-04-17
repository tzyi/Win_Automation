# Prompts

```
@file:exec.json 裡面包含我想要操作的UI元件資訊
以及想對按鈕執行的動作
想透過pywinauto去操作UI元件

我希望python可以動態讀取exec.json裡面的資訊
我之後只需要修改exec.json裡面的資訊就可以達到我要的自動化操作

利用上面資訊幫我寫.py
並且要加上Logging功能
並且執行到哪個步驟就要print在terminal上面
所有的click請使用click_input()
```


```
請幫我抓取UI元件的邏輯更改
請依照下面抓取優先順序來抓取UI元件

1. AutomationId（auto_id）
2. ClassName（class_name）+ ControlType
3. Name（title）+ ControlType + 父層容器限定
4. found_index（同類型第 N 個）
```