@echo off
echo ========================================
echo YYTools 启动测试脚本
echo ========================================
echo.

echo 正在编译项目...
cd YYTools

echo.
echo 选择启动方式：
echo 1. 最小化启动（推荐先测试）
echo 2. 紧急启动
echo 3. 正常启动
echo 4. 退出
echo.

set /p choice="请输入选择 (1-4): "

if "%choice%"=="1" goto minimal
if "%choice%"=="2" goto emergency
if "%choice%"=="3" goto normal
if "%choice%"=="4" goto exit
goto invalid

:minimal
echo.
echo 正在使用最小化启动方式...
echo 这将显示详细的启动过程信息
echo.
pause
goto compile

:emergency
echo.
echo 正在使用紧急启动方式...
echo 跳过所有复杂功能，只启动基本窗体
echo.
pause
goto compile

:normal
echo.
echo 正在使用正常启动方式...
echo 包含所有功能，但可能不稳定
echo.
pause
goto compile

:invalid
echo.
echo 无效选择，请重新运行脚本
pause
exit /b 1

:compile
echo.
echo 正在编译项目...
msbuild YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /verbosity:minimal

if %errorlevel% neq 0 (
    echo.
    echo 编译失败！请检查错误信息
    pause
    exit /b 1
)

echo.
echo 编译成功！正在测试启动...

if "%choice%"=="1" (
    echo 使用最小化启动方式...
    copy Program_Minimal.cs Program.cs >nul
    msbuild YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /verbosity:minimal
    if %errorlevel% equ 0 (
        echo 最小化版本编译成功，正在启动...
        start YYTools.exe
    )
) else if "%choice%"=="2" (
    echo 使用紧急启动方式...
    copy Program_Emergency.cs Program.cs >nul
    msbuild YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /verbosity:minimal
    if %errorlevel% equ 0 (
        echo 紧急版本编译成功，正在启动...
        start YYTools.exe
    )
) else if "%choice%"=="3" (
    echo 使用正常启动方式...
    copy Program.cs.bak Program.cs >nul 2>nul
    if not exist "Program.cs.bak" (
        echo 警告：未找到备份文件，使用当前版本
    )
    msbuild YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /verbosity:minimal
    if %errorlevel% equ 0 (
        echo 正常版本编译成功，正在启动...
        start YYTools.exe
    )
)

echo.
echo 启动测试完成！
echo 如果程序仍然无法启动，请查看错误信息
echo 错误信息会保存到 startup_error*.log 文件
echo.
pause

:exit
echo 退出测试脚本