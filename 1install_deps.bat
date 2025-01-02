@echo off
chcp 65001 >nul

echo 正在安装依赖库...
pip3 install openpyxl==3.0.10 Pillow==9.5.0 lunar-python==1.3.12 requests==2.28.2

echo.
echo 验证安装...
python -c "import openpyxl; import PIL; import lunar_python; import requests; print('所有依赖库安装成功！')"

if %errorlevel% equ 0 (
    echo.
    echo 安装完成！
) else (
    echo.
    echo 安装失败，请检查错误信息。
)

echo.
echo 按任意键退出...
pause >nul 