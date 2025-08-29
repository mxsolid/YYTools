@echo off
chcp 65001 >nul
echo ========================================
echo YY工具 DPI优化版本编译测试
echo ========================================
echo.

echo 正在检查环境...
if not exist "YYTools\YYTools.csproj" (
    echo 错误: 找不到项目文件 YYTools\YYTools.csproj
    pause
    exit /b 1
)

echo 正在清理之前的编译结果...
if exist "YYTools\bin" rmdir /s /q "YYTools\bin"
if exist "YYTools\obj" rmdir /s /q "YYTools\obj"

echo.
echo 正在编译项目...
cd YYTools

echo 使用 MSBuild 编译...
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /verbosity:minimal

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo 编译成功！
    echo ========================================
    echo.
    echo 编译输出位置: YYTools\bin\Debug\
    echo.
    echo 正在检查生成的文件...
    if exist "bin\Debug\YYTools.exe" (
        echo ✓ YYTools.exe 已生成
        echo ✓ 文件大小: 
        dir "bin\Debug\YYTools.exe" | find "YYTools.exe"
    ) else (
        echo ✗ YYTools.exe 未找到
    )
    
    if exist "bin\Debug\YYTools.dll" (
        echo ✓ YYTools.dll 已生成
        echo ✓ 文件大小:
        dir "bin\Debug\YYTools.dll" | find "YYTools.dll"
    ) else (
        echo ✗ YYTools.dll 未找到
    )
    
    echo.
    echo 编译测试完成！
    
) else (
    echo.
    echo ========================================
    echo 编译失败！
    echo ========================================
    echo.
    echo 请检查错误信息并修复问题。
    echo.
    pause
    exit /b 1
)

echo.
echo 按任意键退出...
pause >nul 