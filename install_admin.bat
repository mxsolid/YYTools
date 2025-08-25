@echo off
chcp 65001 >nul

echo ========================================
echo YYTools 管理员安装脚本
echo ========================================

REM 检查管理员权限
net session >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo 错误：需要管理员权限
    echo 请右键点击此脚本，选择"以管理员身份运行"
    pause
    exit /b 1
)

echo 已检测到管理员权限

REM 设置.NET Framework路径
set FRAMEWORK_PATH=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
if not exist "%FRAMEWORK_PATH%\csc.exe" (
    set FRAMEWORK_PATH=%WINDIR%\Microsoft.NET\Framework\v4.0.30319
)

if not exist "%FRAMEWORK_PATH%\csc.exe" (
    echo 错误：未找到.NET Framework 4.0编译器
    echo 请确保已安装.NET Framework 4.0或更高版本
    pause
    exit /b 1
)

echo 找到编译器: %FRAMEWORK_PATH%\csc.exe

REM 设置输出目录
set OUTPUT_DIR=bin\Debug
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

echo 开始编译 YYTools...

REM 编译主文件
"%FRAMEWORK_PATH%\csc.exe" ^
    /target:library ^
    /out:"%OUTPUT_DIR%\YYTools.dll" ^
    /reference:System.dll ^
    /reference:System.Windows.Forms.dll ^
    /reference:System.Drawing.dll ^
    YYTools\ExcelAddin.cs YYTools\ColumnSelectionForm.cs

if %ERRORLEVEL% neq 0 (
    echo.
    echo ========================================
    echo 编译失败！请检查语法错误
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo 编译成功！
echo ========================================

REM 检查文件是否生成
if exist "%OUTPUT_DIR%\YYTools.dll" (
    echo DLL文件已生成:
    dir "%OUTPUT_DIR%\YYTools.dll" | find "YYTools.dll"
) else (
    echo 错误：DLL文件未生成
    pause
    exit /b 1
)

echo.
echo 正在注册COM组件...

REM 先尝试取消注册旧版本
"%FRAMEWORK_PATH%\RegAsm.exe" "%OUTPUT_DIR%\YYTools.dll" /u /silent >nul 2>&1

REM 注册新版本
"%FRAMEWORK_PATH%\RegAsm.exe" "%OUTPUT_DIR%\YYTools.dll" /codebase /tlb

if %ERRORLEVEL% neq 0 (
    echo.
    echo 错误：COM注册失败
    echo 可能的原因：
    echo 1. 程序集未签名
    echo 2. 系统权限不足
    echo 3. 防病毒软件阻止
    pause
    exit /b 1
) else (
    echo COM组件注册成功！
)

echo.
echo ========================================
echo 安装完成！
echo.
echo 现在可以：
echo 1. 打开WPS表格或Excel
echo 2. 查看是否出现"YY工具"菜单
echo 3. 或运行VBA宏测试功能
echo.
echo 测试VBA代码：
echo Sub 测试YYTools()
echo     CreateObject("YYTools.ExcelAddin").GetApplicationInfo()
echo End Sub
echo ========================================

pause 