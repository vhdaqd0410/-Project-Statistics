@echo off
chcp 65001 >nul
echo ==================================================
echo  AI后期剪辑提成工具 - 打包单文件 exe（通用版）
echo ==================================================
echo.

cd /d "%~dp0"

set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

echo [1/4] 检查 PyInstaller...
"%PY%" -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo   未安装，正在安装 PyInstaller（使用清华镜像）...
    "%PY%" -m pip install pyinstaller -q -i https://pypi.tuna.tsinghua.edu.cn/simple
    if errorlevel 1 (
        "%PY%" -m pip install pyinstaller -q
    )
) else (
    echo   PyInstaller 已安装。
)
echo.

echo [2/4] 清理旧打包产物...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo   已清理 build/ 与 dist/
echo.

echo [3/4] 开始打包（单文件、无控制台窗口）...
"%PY%" -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "提成工具" ^
    --add-data "src;src" ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import matplotlib ^
    --hidden-import PIL ^
    --collect-all matplotlib ^
    ai_commission_gui.py
if errorlevel 1 (
    echo ❌ 打包失败！请查看上方错误信息。
    pause
    exit /b 1
)
echo   打包完成。
echo.

echo [4/4] 组装发布目录...
if exist "提成工具-便携版" rmdir /s /q "提成工具-便携版"
mkdir "提成工具-便携版"
mkdir "提成工具-便携版\data"
copy /y "dist\提成工具.exe" "提成工具-便携版\" >nul
copy /y "release\config.json" "提成工具-便携版\config.json" >nul
copy /y "release\data\AI后期剪辑提成一组模板.xlsx" "提成工具-便携版\data\" >nul
copy /y "release\data\一组AI项目-8月.xlsx" "提成工具-便携版\data\" >nul
echo.

echo ==================================================
echo   打包完成！发布目录: 提成工具-便携版\
echo   把整个"提成工具-便携版"文件夹拷到目标电脑即可。
echo ==================================================
pause
