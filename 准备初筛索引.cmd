@echo off
chcp 65001 >nul
call "%~dp0查找Python.cmd"
if errorlevel 1 exit /b 1
"%附注Python%" -B -X utf8 "%~dp0结构初筛\准备索引.py"
set "准备退出码=%errorlevel%"
pause
exit /b %准备退出码%
