@echo off
chcp 65001 >nul
call "%~dp0查找Python.cmd"
if errorlevel 1 exit /b 1
"%附注Python%" -B -X utf8 "%~dp0验证公开版.py"
set "验证退出码=%errorlevel%"
pause
exit /b %验证退出码%
