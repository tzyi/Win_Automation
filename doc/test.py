from pywinauto import Application

app = Application(backend="uia").connect(title="GeoTag Pro")

dlg = app.window(title="GeoTag Pro")

# 列印所有子控制項資訊，協助確認屬性
# for ctrl in dlg.descendants():
#     try:
#         print(f"title: {ctrl.window_text()}, control_type: {ctrl.friendly_class_name()}, automation_id: {getattr(ctrl, 'automation_id', lambda: None)()}")
#     except Exception as e:
#         print(f"Error reading control: {e}")

dlg.child_window(
    title="folder_open 選擇照片",
    control_type="Button"
).click()