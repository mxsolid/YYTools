@echo off
chcp 65001 >nul

echo ========================================
echo 编译YYTools测试程序
echo ========================================

REM 设置.NET Framework路径
set FRAMEWORK_PATH=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
if not exist "%FRAMEWORK_PATH%\csc.exe" (
    set FRAMEWORK_PATH=%WINDIR%\Microsoft.NET\Framework\v4.0.30319
)

if not exist "%FRAMEWORK_PATH%\csc.exe" (
    echo 错误：未找到.NET Framework 4.0编译器
    pause
    exit /b 1
)

echo 找到编译器: %FRAMEWORK_PATH%\csc.exe

REM 创建输出目录
if not exist "TestApp" mkdir "TestApp"
set OUTPUT_DIR=bin\Debug

echo 开始编译测试程序...

REM 编译带界面的测试EXE
"%FRAMEWORK_PATH%\csc.exe" ^
    /target:winexe ^
    /out:"%OUTPUT_DIR%\YYToolsTest.exe" ^
    /reference:System.dll ^
    /reference:System.Windows.Forms.dll ^
    /reference:System.Drawing.dll ^
    TestApp\Program.cs TestApp\MainForm.cs

if %ERRORLEVEL% neq 0 (
    echo.
    echo 编译失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo 编译成功！
echo 输出文件: %OUTPUT_DIR%\YYToolsTest.exe
echo ========================================

if exist "%OUTPUT_DIR%\YYToolsTest.exe" (
    echo 测试EXE已生成:
    dir "%OUTPUT_DIR%\YYToolsTest.exe" | find "YYToolsTest.exe"
    echo.
    echo 使用方法:
    echo 1. 界面模式 (默认): %OUTPUT_DIR%\YYToolsTest.exe
    echo 2. 控制台模式: %OUTPUT_DIR%\YYToolsTest.exe console
) else (
    echo 错误：EXE文件未生成
)

pause 