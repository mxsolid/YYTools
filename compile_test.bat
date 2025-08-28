@echo off
chcp 65001 >nul
echo ========================================
echo YYTools 编译测试脚本
echo ========================================
echo.

echo 正在检查编译环境...
where msbuild >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到 MSBuild
    echo 请确保已安装 Visual Studio 或 .NET Framework SDK
    pause
    exit /b 1
)

echo 找到 MSBuild，开始编译...
echo.

cd YYTools
msbuild YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /verbosity:minimal

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo 编译成功！
    echo ========================================
    echo.
    echo 输出文件: YYTools\bin\Debug\YYTools.exe
    echo.
    echo 所有编译错误已修复！
) else (
    echo.
    echo ========================================
    echo 编译失败！
    echo ========================================
    echo.
    echo 请检查上述错误信息并修复
)

echo.
pause 