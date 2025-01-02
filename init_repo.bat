@echo off
chcp 65001 >nul

echo 初始化Git仓库...
git init

echo.
echo 添加远程仓库...
git remote add origin https://github.com/zghk/chinese_calendar.git

echo.
echo 创建文档图片目录...
mkdir docs\images

echo.
echo 添加所有文件...
git add .

echo.
echo 提交更改...
git commit -m "Initial commit: Chinese Calendar Generator
echo.
echo 切换到main分支...
git branch -M main

echo.
echo 推送到GitHub...
git push -u origin main

echo.
echo 完成！
echo 请检查 https://github.com/zghk/chinese_calendar
pause