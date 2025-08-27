@echo off
chcp 65001 >nul
title YY工具 - .NET Framework 检测工具

echo ========================================
echo YY工具 v3.0 - 运行环境检测
echo ========================================
echo.

:: 检测.NET Framework版本
echo [1/3] 检测.NET Framework版本...
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release >nul 2>&1
if %errorlevel% equ 0 (
    for /f "tokens=3" %%i in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release') do set release=%%i
    if !release! geq 528040 (
        echo ✓ 已安装.NET Framework 4.8或更高版本
        echo   版本号: !release!
        goto :check_excel
    ) else (
        echo ✗ 检测到.NET Framework 4.0，但版本过低
        echo   当前版本: !release!
        echo   需要版本: 528040或更高
        goto :install_dotnet
    )
) else (
    echo ✗ 未检测到.NET Framework 4.0或更高版本
    goto :install_dotnet
)

:check_excel
echo.
echo [2/3] 检测Office环境...
reg query "HKEY_CLASSES_ROOT\Excel.Application" >nul 2>&1
if %errorlevel% equ 0 (
    echo ✓ 检测到Excel
) else (
    echo ⚠ 未检测到Excel，但可能安装了WPS
)

reg query "HKEY_CLASSES_ROOT\WPS.Application" >nul 2>&1
if %errorlevel% equ 0 (
    echo ✓ 检测到WPS
) else (
    echo ⚠ 未检测到WPS
)

echo.
echo [3/3] 环境检测完成
echo ✓ 运行环境满足要求
echo.
echo 现在可以运行YY工具了！
echo.
pause
exit /b 0

:install_dotnet
echo.
echo ========================================
echo 需要安装.NET Framework 4.8
echo ========================================
echo.
echo 请选择安装方式:
echo.
echo 1. 自动下载并安装 (推荐)
echo 2. 手动下载安装
echo 3. 退出
echo.
set /p choice="请输入选择 (1-3): "

if "%choice%"=="1" goto :auto_install
if "%choice%"=="2" goto :manual_install
if "%choice%"=="3" goto :exit
goto :install_dotnet

:auto_install
echo.
echo 正在下载.NET Framework 4.8...
echo 这可能需要几分钟时间，请耐心等待...
echo.

:: 下载.NET Framework 4.8
powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://go.microsoft.com/fwlink/?LinkId=2085150' -OutFile 'ndp48-web.exe'}"
if %errorlevel% neq 0 (
    echo 下载失败，请检查网络连接
    goto :manual_install
)

echo 下载完成，正在安装...
ndp48-web.exe /quiet /norestart
if %errorlevel% equ 0 (
    echo.
    echo ✓ .NET Framework 4.8 安装成功！
    echo 请重启计算机后再次运行此脚本
    echo.
    del ndp48-web.exe
    pause
    exit /b 0
) else (
    echo.
    echo ✗ 安装失败，请尝试手动安装
    del ndp48-web.exe
    goto :manual_install
)

:manual_install
echo.
echo ========================================
echo 手动安装.NET Framework 4.8
echo ========================================
echo.
echo 请按以下步骤操作:
echo.
echo 1. 访问微软官方下载页面:
echo    https://dotnet.microsoft.com/download/dotnet-framework/net48
echo.
echo 2. 下载 "Runtime" 版本
echo.
echo 3. 运行下载的安装程序
echo.
echo 4. 安装完成后重启计算机
echo.
echo 5. 重新运行此脚本检测环境
echo.
pause
exit /b 0

:exit
echo.
echo 感谢使用YY工具！
echo.
pause
exit /b 0