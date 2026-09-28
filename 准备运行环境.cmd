@echo off
chcp 65001 >nul
cd /d "%~dp0"
call "%~dp0查找Python.cmd"
if errorlevel 1 exit /b 1
echo 正在为选定的本机 Python 安装本项目所需组件。
"%附注Python%" -m pip install -r "%~dp0requirements.txt"
set "准备退出码=%errorlevel%"
if not "%准备退出码%"=="0" goto 完成
echo Python组件准备完成。结构初筛和纯Excel操作现在可以使用。
echo 需要提取或回写Word时，请另行双击“准备Word功能.cmd”。
:完成
pause
exit /b %准备退出码%
