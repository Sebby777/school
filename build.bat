@echo off
chcp 65001 > nul

echo 正在安装所需的Python库...
pip install pyinstaller pandas openpyxl

if errorlevel 1 (
    echo 安装库时出现错误，请检查Python和pip是否正确安装
    pause
    exit /b 1
)

echo 正在打包应用程序...
pyinstaller --onefile --windowed --name "配码处理工具" --add-data "码数表.xlsx;." --add-data "格式模版.xlsx;." --icon=NONE head.py

if errorlevel 1 (
    echo 打包过程中出现错误
    pause
    exit /b 1
)

echo 打包完成！可执行文件位于 dist 文件夹中
echo.
echo 按任意键退出...
pause > nul