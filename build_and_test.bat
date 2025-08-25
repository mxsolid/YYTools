@echo off
chcp 65001 >nul

echo ========================================
echo YYTools 编译和测试脚本
echo ========================================

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

echo 开始编译 ExcelAddin.cs...

REM 编译主文件，不依赖Office Interop（运行时通过COM调用）
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
echo 输出文件: %OUTPUT_DIR%\YYTools.dll
echo ========================================

REM 检查文件是否生成
if exist "%OUTPUT_DIR%\YYTools.dll" (
    echo DLL文件已生成，大小: 
    dir "%OUTPUT_DIR%\YYTools.dll" | find "YYTools.dll"
) else (
    echo 警告：DLL文件未找到
)

echo.
echo 正在注册COM组件...

REM 注册COM组件（需要管理员权限）
"%FRAMEWORK_PATH%\RegAsm.exe" "%OUTPUT_DIR%\YYTools.dll" /codebase /tlb

if %ERRORLEVEL% neq 0 (
    echo.
    echo 警告：COM注册失败，可能需要管理员权限
    echo 请以管理员身份运行此脚本
) else (
    echo COM组件注册成功！
)

echo.
echo ========================================
echo 构建完成！
echo.
echo 如需在WPS中使用，请确保：
echo 1. 以管理员身份运行过此脚本
echo 2. WPS表格已安装
echo 3. 在WPS中启用宏和COM加载项
echo ========================================

pause 