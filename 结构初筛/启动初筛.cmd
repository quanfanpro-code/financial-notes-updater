@echo off
chcp 65001 >nul
call "%~dp0..\查找Python.cmd"
if errorlevel 1 exit /b 1
"%附注Python%" -B -X utf8 "%~dp0运行初筛.py" %*
set "初筛退出码=%errorlevel%"
if "%~1"=="" pause
exit /b %初筛退出码%
