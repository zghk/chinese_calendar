@echo off
chcp 65001 >nul

set /p year="请输入要生成的年份（如：2024）: "
python chinese_calendar.py --year %year%
pause 