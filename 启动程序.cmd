@echo off
chcp 65001 >nul
cd /d "%~dp0"
call "%~dp0查找Python.cmd"
if errorlevel 1 exit /b 1
"%附注Python%" -B -X utf8 "%~dp0启动程序.pyw"
