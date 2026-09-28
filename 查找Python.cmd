@echo off
chcp 65001 >nul
set "附注Python=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
if exist "%附注Python%" exit /b 0
set "附注Python=%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe"
if exist "%附注Python%" exit /b 0
set "附注Python="
echo 请在弹出的窗口中选择已安装的 Python 3.14 程序 python.exe。
for /f "usebackq delims=" %%P in (`powershell.exe -NoProfile -STA -Command "[Console]::OutputEncoding=[Text.Encoding]::UTF8; Add-Type -AssemblyName System.Windows.Forms; $d=New-Object System.Windows.Forms.OpenFileDialog; $d.Title='选择本机 Python 3.14 的 python.exe'; $d.Filter='Python程序 (python.exe)|python.exe'; if($d.ShowDialog() -eq 'OK'){Write-Output $d.FileName}"`) do set "附注Python=%%P"
if not defined 附注Python exit /b 1
if not exist "%附注Python%" exit /b 1
exit /b 0
