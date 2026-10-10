@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion
title pptx-slim - PPT 瘦身

echo ════════════════════════════════════════════════
echo    PPT 瘦身  ^(pptx-slim^)
echo ════════════════════════════════════════════════
echo.

rem ---- 1. 找 Python ----------------------------------------------------
set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY ( where py >nul 2>&1 && set "PY=py" )
if not defined PY (
    echo [X] 没找到 Python。
    echo.
    echo     请到 https://www.python.org/downloads/ 下载安装。
    echo     安装第一屏务必勾选  "Add python.exe to PATH"
    echo     装完关掉这个窗口，重新双击一次。
    echo.
    pause
    exit /b 1
)
echo [1/3] 找到 Python

rem ---- 2. 检查 Pillow --------------------------------------------------
%PY% -c "import PIL" >nul 2>&1
if errorlevel 1 (
    echo [2/3] 首次运行，正在安装 Pillow ^(需要联网，约 10 秒^)...
    %PY% -m pip install --user --quiet Pillow
    %PY% -c "import PIL" >nul 2>&1
    if errorlevel 1 (
        echo [X] Pillow 装不上。请把上面的报错截图发给我。
        pause
        exit /b 1
    )
) else (
    echo [2/3] 依赖齐全
)

rem ---- 3. 问参数、开跑 -------------------------------------------------
set "FILE=%~1"
if "%FILE%"=="" (
    echo.
    set /p "FILE=把 .pptx 文件拖进这个窗口，然后按回车: "
)
set "FILE=%FILE:"=%"

if not exist "%FILE%" (
    echo [X] 找不到这个文件: %FILE%
    pause
    exit /b 1
)

echo.
set "MB="
set /p "MB=目标大小 MB ^(直接回车 = 5^): "
if "%MB%"=="" set "MB=5"

echo.
echo [3/3] 开始压缩...
echo ────────────────────────────────────────────────
%PY% "%~dp0..\pptx_slim.py" "%FILE%" %MB%
echo ────────────────────────────────────────────────
echo.
echo 结果保存在原文件旁边，名字后面带 " (compressed)"
echo.
pause
