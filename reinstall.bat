@echo off
chcp 65001 >nul

echo 正在卸载现有的依赖库...
pip3 uninstall -y openpyxl Pillow lunar-python requests pywin32

echo.
echo 正在安装指定版本的依赖库...

echo 安装 openpyxl...
pip3 install openpyxl==3.0.10
if %errorlevel% neq 0 goto :error

echo 安装 Pillow...
pip3 install Pillow==9.5.0
if %errorlevel% neq 0 goto :error

echo 安装 lunar-python...
pip3 install lunar-python==1.3.12
if %errorlevel% neq 0 goto :error

echo 安装 requests...
pip3 install requests==2.28.2
if %errorlevel% neq 0 goto :error

echo 安装 pywin32...
pip3 install pywin32==306
if %errorlevel% neq 0 goto :error

echo.
echo 验证安装...
python -c "import openpyxl; import PIL; import lunar_python; import requests; import win32com.client; print('所有库安装成功！')"
if %errorlevel% neq 0 goto :verify_error

goto :success

:error
echo.
echo 安装过程中出现错误，请检查错误信息。
goto :end

:verify_error
echo.
echo 库已安装但验证失败，可能需要重启命令行或IDE。
goto :end

:success
echo.
echo 所有依赖库安装成功！

:end
echo.
echo 按任意键退出...
pause >nul 