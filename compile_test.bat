@echo off
echo ========================================
echo YYTools 编译测试脚本
echo ========================================
echo.

echo 正在检查编译环境...
where msbuild >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到 MSBuild，请确保已安装 Visual Studio 或 .NET Framework SDK
    pause
    exit /b 1
)

echo 正在清理旧的编译文件...
if exist "YYTools\bin" rmdir /s /q "YYTools\bin"
if exist "YYTools\obj" rmdir /s /q "YYTools\obj"

echo.
echo 正在编译项目...
cd YYTools
msbuild YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /verbosity:minimal

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo 编译成功！
    echo ========================================
    echo.
    echo 输出文件位置: YYTools\bin\Debug\YYTools.exe
    echo.
    
    if exist "bin\Debug\YYTools.exe" (
        echo 正在验证输出文件...
        echo 文件大小: 
        dir "bin\Debug\YYTools.exe" | findstr "YYTools.exe"
        echo.
        echo 编译测试完成！
    ) else (
        echo 警告: 未找到输出文件
    )
) else (
    echo.
    echo ========================================
    echo 编译失败！
    echo ========================================
    echo.
    echo 请检查错误信息并修复问题
)

echo.
pause 