@echo off
chcp 65001 >nul
echo ========================================
echo   WinAutomation 打包腳本
echo ========================================
echo.

REM 啟動虛擬環境
call .venv\Scripts\activate.bat

echo [1/3] 打包 run.exe (自動化引擎)...
pyinstaller --noconfirm --clean ^
    --onefile ^
    --console ^
    --name run ^
    --hidden-import=comtypes.gen.UIAutomationClient ^
    --hidden-import=comtypes.gen.stdole ^
    --hidden-import=comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0 ^
    --hidden-import=comtypes.gen._944DE083_8FB8_45CF_BCB7_C477ACB2F897_0_1_0 ^
    --hidden-import=comtypes.stream ^
    --hidden-import=pywinauto.controls.uiawrapper ^
    --hidden-import=pywinauto.controls.uia_controls ^
    run.py

if %ERRORLEVEL% neq 0 (
    echo [錯誤] run.exe 打包失敗！
    pause
    exit /b 1
)

echo.
echo [2/3] 打包 WinAutomation.exe (主程式 GUI)...
pyinstaller --noconfirm --clean ^
    --onefile ^
    --windowed ^
    --name WinAutomation ^
    --hidden-import=comtypes.gen.UIAutomationClient ^
    --hidden-import=comtypes.gen.stdole ^
    --hidden-import=comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0 ^
    --hidden-import=comtypes.gen._944DE083_8FB8_45CF_BCB7_C477ACB2F897_0_1_0 ^
    --hidden-import=comtypes.stream ^
    inspect_tool.py

if %ERRORLEVEL% neq 0 (
    echo [錯誤] WinAutomation.exe 打包失敗！
    pause
    exit /b 1
)

echo.
echo [3/3] 整理輸出檔案...

REM 確認輸出目錄
if not exist "dist" (
    echo [錯誤] dist 目錄不存在！
    pause
    exit /b 1
)

echo.
echo ========================================
echo   打包完成！
echo   輸出位置: dist\WinAutomation.exe
echo              dist\run.exe
echo.
echo   部署時請將兩個 exe 放在同一目錄，
echo   設定檔 (*.json) 也放在同一目錄即可。
echo ========================================
pause
