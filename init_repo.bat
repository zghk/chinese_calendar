@echo off
chcp 65001 >nul

echo ��ʼ��Git�ֿ�...
git init

echo.
echo ���Զ�ֿ̲�...
git remote add origin https://github.com/zghk/chinese_calendar.git

echo.
echo �����ĵ�ͼƬĿ¼...
mkdir docs\images

echo.
echo ��������ļ�...
git add .

echo.
echo �ύ����...
git commit -m "Initial commit: Chinese Calendar Generator
echo.
echo �л���main��֧...
git branch -M main

echo.
echo ���͵�GitHub...
git push -u origin main

echo.
echo ��ɣ�
echo ���� https://github.com/zghk/chinese_calendar
pause