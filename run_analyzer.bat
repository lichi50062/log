@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo.
echo ===============================================
echo           Log執行時間分析工具
echo ===============================================
echo.

REM 檢查執行檔
if not exist "LogAnalyzer.exe" (
    echo 錯誤: 找不到 LogAnalyzer.exe
    echo 請確保 LogAnalyzer.exe 在當前目錄
    pause
    exit /b 1
)

echo 已找到: LogAnalyzer.exe
echo.

REM 顯示目錄檔案
echo 當前目錄的檔案:
dir /b | findstr /v /i "\.exe$ \.bat$"
echo.

echo 常見搜尋範例:
echo   格式: "執行時間: 123 ms"   前綴="執行時間: " 後綴=" ms"
echo   格式: "耗時 456 毫秒"      前綴="耗時 "      後綴=" 毫秒"  
echo   格式: "duration: 789ms"   前綴="duration: " 後綴="ms"
echo.

:get_prefix
set /p prefix="請輸入前綴: "
if "!prefix!"=="" goto get_prefix

:get_suffix
set /p suffix="請輸入後綴: "
if "!suffix!"=="" goto get_suffix

echo.
echo 開始分析...
echo 前綴: "!prefix!"
echo 後綴: "!suffix!"
echo.

LogAnalyzer.exe . "!prefix!" "!suffix!"

if !ERRORLEVEL! equ 0 (
    echo.
    echo 分析成功完成!
    echo 請檢查產生的Excel檔案
) else (
    echo.
    echo 分析失敗，錯誤代碼: !ERRORLEVEL!
)

echo.
pause