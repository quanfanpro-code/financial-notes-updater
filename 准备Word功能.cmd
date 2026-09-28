@echo off
chcp 65001 >nul
set "附注Dotnet=%ProgramFiles%\dotnet\dotnet.exe"
if not exist "%附注Dotnet%" (
  echo 请先从微软官网安装 .NET 10 开发工具包，再运行本文件。
  echo 下载页面：https://dotnet.microsoft.com/zh-cn/download/dotnet/10.0
  pause
  exit /b 1
)
if exist "%~dp0Word桥接\发布\WordBridge.exe" (
  echo Word功能已准备好。为保护已有成果，本入口不覆盖现有程序。
  pause
  exit /b 0
)
"%附注Dotnet%" publish "%~dp0Word桥接\Word桥接.csproj" -c Release --no-self-contained -o "%~dp0Word桥接\发布"
set "准备退出码=%errorlevel%"
pause
exit /b %准备退出码%
