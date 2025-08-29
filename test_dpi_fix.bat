@echo off
chcp 65001 >nul
echo ========================================
echo DPI修复测试程序
echo ========================================
echo.

echo 正在编译DPI测试程序...
echo.

REM 尝试使用csc编译器
if exist "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe" (
    echo 使用 .NET Framework 4.0 编译器...
    "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe" /target:winexe /reference:"C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Windows.Forms.dll" /reference:"C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Drawing.dll" /reference:"C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorlib.dll" "DPI_TEST_PROGRAM.cs" /out:DPI_TEST.exe
    goto :check_result
)

REM 尝试使用csc编译器 (64位)
if exist "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" (
    echo 使用 .NET Framework 4.0 64位编译器...
    "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /target:winexe /reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Windows.Forms.dll" /reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Drawing.dll" /reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.dll" "DPI_TEST_PROGRAM.cs" /out:DPI_TEST.exe
    goto :check_result
)

REM 尝试使用dotnet
if command -v dotnet >nul 2>&1 (
    echo 使用 .NET Core 编译器...
    dotnet new console --force
    copy "DPI_TEST_PROGRAM.cs" "Program.cs"
    dotnet build
    if exist "bin\Debug\net6.0\DPI_TEST_PROGRAM.exe" (
        copy "bin\Debug\net6.0\DPI_TEST_PROGRAM.exe" "DPI_TEST.exe"
        goto :check_result
    )
)

echo 错误: 未找到可用的编译器
echo 请确保安装了 .NET Framework 或 .NET Core
pause
exit /b 1

:check_result
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo 编译成功！
    echo ========================================
    echo.
    echo 正在检查生成的文件...
    if exist "DPI_TEST.exe" (
        echo ✓ DPI_TEST.exe 已生成
        echo ✓ 文件大小:
        dir "DPI_TEST.exe" | find "DPI_TEST.exe"
        echo.
        echo 现在可以运行 DPI_TEST.exe 来测试DPI修复效果
        echo.
        echo 测试步骤:
        echo 1. 运行 DPI_TEST.exe
        echo 2. 点击"测试DPI"按钮查看字体大小
        echo 3. 检查界面是否清晰且字体大小合适
        echo.
        echo 如果测试成功，说明DPI修复方案有效
        echo 如果仍有问题，请检查日志输出
    ) else (
        echo ✗ DPI_TEST.exe 未找到
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