@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo Log檔執行時間分析工具 (當前資料夾)
echo ===================================
echo.

REM 獲取批次檔所在的資料夾
set "SCRIPT_DIR=%~dp0"

REM 使用相對路徑的Python (在當前環境中)
set PYTHON=python

REM 確認Python腳本文件是否存在
set PYTHON_SCRIPT="%SCRIPT_DIR%simple_log_analyzer.py"
if not exist %PYTHON_SCRIPT% (
    echo 錯誤: 找不到分析腳本 simple_log_analyzer.py
    echo 請確保批次檔與 simple_log_analyzer.py 在同一個資料夾
    pause
    exit /b 1
)

REM 安裝必要的庫
echo 檢查必要的Python庫...
%PYTHON% -m pip install pandas numpy xlsxwriter
if %ERRORLEVEL% neq 0 (
    echo 安裝必要的Python庫時出錯
    pause
    exit /b 1
)

echo.
echo 請輸入要搜尋的前綴字串(完全匹配，包含空格):
echo 例如要搜尋 "bbbbaaa 1 ms" 中的數字1，請輸入 "bbbbaaa "
set /p prefix="> "

if "!prefix!"=="" (
    echo 錯誤: 必須提供前綴字串
    pause
    exit /b 1
)

echo.
echo 請輸入要搜尋的後綴字串(完全匹配，包含空格):
echo 例如要搜尋 "bbbbaaa 1 ms" 中的數字1，請輸入 " ms"
set /p suffix="> "

if "!suffix!"=="" (
    echo 錯誤: 必須提供後綴字串
    pause
    exit /b 1
)

REM 設置參數，使用當前資料夾，保留前綴後綴中的空格
set params="." "!prefix!" "!suffix!"

echo.
echo 執行分析...
%PYTHON% %PYTHON_SCRIPT% !params!

if %ERRORLEVEL% neq 0 (
    echo.
    echo 分析過程中出錯
) else (
    echo.
    echo 分析完成！
)

pause
exit /b 0