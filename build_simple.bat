@echo off
chcp 65001 >nul
echo =====================================
echo YY工具 简化编译脚本 v2.1
echo =====================================

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
if not exist "YYTools\bin\Debug" mkdir "YYTools\bin\Debug"

echo 开始编译YYTools.dll...

REM 编译主DLL
"%FRAMEWORK_PATH%\csc.exe" ^
    /target:library ^
    /out:"YYTools\bin\Debug\YYTools.dll" ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    /reference:System.Data.dll ^
    /reference:System.Drawing.dll ^
    /reference:System.Windows.Forms.dll ^
    /reference:System.Xml.dll ^
    /reference:"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE" ^
    /reference:"C:\Windows\Microsoft.NET\assembly\GAC_MSIL\office\v4.0_15.0.0.0__71e9bce111e9429c\office.dll" ^
    YYTools\*.cs

if %ERRORLEVEL% neq 0 (
    echo.
    echo 编译失败！尝试不使用Office引用...
    
    "%FRAMEWORK_PATH%\csc.exe" ^
        /target:library ^
        /out:"YYTools\bin\Debug\YYTools.dll" ^
        /reference:System.dll ^
        /reference:System.Core.dll ^
        /reference:System.Data.dll ^
        /reference:System.Drawing.dll ^
        /reference:System.Windows.Forms.dll ^
        /reference:System.Xml.dll ^
        YYTools\ExcelHelper.cs ^
        YYTools\MatchService.cs ^
        YYTools\ColumnSelectionForm.cs ^
        YYTools\MatchForm.cs ^
        YYTools\SettingsForm.cs ^
        YYTools\Properties\AssemblyInfo.cs
        
    if %ERRORLEVEL% neq 0 (
        echo 编译仍然失败！
        pause
        exit /b 1
    )
)

echo.
echo ========================================
echo 编译成功！
echo ========================================

if exist "YYTools\bin\Debug\YYTools.dll" (
    echo DLL文件已生成:
    dir "YYTools\bin\Debug\YYTools.dll" | find "YYTools.dll"
) else (
    echo 错误：DLL文件未生成
)

echo.
echo 接下来可以：
echo 1. 以管理员身份运行: install_admin.bat
echo 2. 测试功能: bin\Debug\YYToolsTest.exe
echo.

pause 