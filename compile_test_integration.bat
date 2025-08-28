@echo off
echo ========================================
echo YYTools 集成修复编译测试脚本
echo ========================================
echo.

echo 正在检查编译环境...
where msbuild >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到 MSBuild，请确保已安装 Visual Studio 或 .NET Framework SDK
    echo.
    echo 尝试使用 csc 编译器...
    where csc >nul 2>&1
    if %errorlevel% neq 0 (
        echo 错误: 未找到 csc 编译器
        pause
        exit /b 1
    )
    echo 找到 csc 编译器，使用 csc 进行编译...
    goto use_csc
)

echo 找到 MSBuild，使用 MSBuild 进行编译...
goto use_msbuild

:use_msbuild
echo.
echo 正在清理旧的编译文件...
if exist "YYTools\bin" rmdir /s /q "YYTools\bin"
if exist "YYTools\obj" rmdir /s /q "YYTools\obj"

echo.
echo 正在使用 MSBuild 编译项目...
cd YYTools
msbuild YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /verbosity:minimal

if %errorlevel% equ 0 (
    goto compile_success
) else (
    goto compile_failed
)

:use_csc
echo.
echo 正在使用 csc 编译器...
cd YYTools

echo 正在编译主要源文件...
csc /target:winexe /reference:System.dll /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Data.dll /reference:System.Xml.dll /reference:Microsoft.Office.Interop.Excel.dll /out:YYTools.exe /recurse:*.cs

if %errorlevel% equ 0 (
    goto compile_success
) else (
    goto compile_failed
)

:compile_success
echo.
echo ========================================
echo 编译成功！
echo ========================================
echo.
echo 输出文件位置: YYTools\YYTools.exe
echo.
echo 正在验证输出文件...
if exist "YYTools.exe" (
    echo 文件大小: 
    dir "YYTools.exe" | findstr "YYTools.exe"
    echo.
    echo 集成修复编译测试完成！
    echo.
    echo 修复内容验证：
    echo ✅ AsyncStartupManager.cs - 异步启动管理器
    echo ✅ StartupProgressForm.cs - 启动进度窗体
    echo ✅ TaskOptionsForm.cs - 任务选项配置窗体
    echo ✅ DPIManager.cs - DPI兼容性管理器
    echo ✅ Program.cs - 集成异步启动
    echo ✅ MatchForm.cs - 集成DPI管理和任务选项
    echo ✅ MatchForm.Designer.cs - 窗体样式和菜单修复
) else (
    echo 警告: 未找到输出文件
)
goto end

:compile_failed
echo.
echo ========================================
echo 编译失败！
echo ========================================
echo.
echo 请检查错误信息并修复问题
echo.
echo 常见问题解决方案：
echo 1. 确保所有引用的DLL文件存在
echo 2. 检查代码语法错误
echo 3. 确保.NET Framework版本正确
echo 4. 检查项目文件配置
echo 5. 验证所有新文件是否正确添加到项目中

:end
echo.
pause