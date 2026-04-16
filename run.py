import time
from pywinauto import Application

# Step 1: GeoTag Pro - 左鍵點擊「folder_open 選擇照片」
app = Application(backend="uia").connect(title="GeoTag Pro")
dlg = app.window(title="GeoTag Pro")
dlg.child_window(title="folder_open 選擇照片", control_type="Text").click()

time.sleep(1)  # 等待「開啟」對話框出現

# Step 2: 開啟對話框 - 左鍵點擊清單項目「IMG_4301.JPG」
open_dlg = app.window(title="開啟")
open_dlg.child_window(title="IMG_4301.JPG", control_type="ListItem").click()

time.sleep(0.5)

# Step 3: 開啟對話框 - 左鍵點擊「開啟(O)」
open_dlg.child_window(title="開啟(O)", control_type="SplitButton").click()
