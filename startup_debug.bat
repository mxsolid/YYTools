@echo off
echo ========================================
echo YYTools 启动问题诊断脚本
echo ========================================
echo.

echo 正在检查系统环境...
echo.

echo 1. 检查.NET Framework版本...
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Version 2>nul
if %errorlevel% equ 0 (
    for /f "tokens=3" %%i in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Version 2^>nul') do (
        echo 当前.NET Framework版本: %%i
    )
) else (
    echo 警告: 未找到.NET Framework 4.0或更高版本
)

echo.
echo 2. 检查系统架构...
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    echo 系统架构: 64位
) else (
    echo 系统架构: 32位
)

echo.
echo 3. 检查Windows版本...
ver
echo.

echo 4. 检查可用内存...
wmic computersystem get TotalPhysicalMemory /format:value 2>nul | find "="
echo.

echo 5. 检查磁盘空间...
for %%i in (C:) do (
    echo C盘可用空间:
    dir C:\ | find "可用字节"
)
echo.

echo 6. 检查日志目录...
if exist "%APPDATA%\YYTools\Logs" (
    echo 日志目录存在: %APPDATA%\YYTools\Logs
    dir "%APPDATA%\YYTools\Logs" /b
) else (
    echo 日志目录不存在: %APPDATA%\YYTools\Logs
)

echo.
echo 7. 检查临时目录...
if exist "%TEMP%\YYTools" (
    echo 临时目录存在: %TEMP%\YYTools
) else (
    echo 临时目录不存在: %TEMP%\YYTools
)

echo.
echo 8. 检查当前目录...
echo 当前目录: %CD%
if exist "YYTools.exe" (
    echo 找到YYTools.exe
    dir "YYTools.exe" | find "YYTools.exe"
) else (
    echo 未找到YYTools.exe
)

echo.
echo 9. 尝试创建测试日志...
echo 测试日志写入时间: %date% %time% > test_log.txt
if exist "test_log.txt" (
    echo 测试日志创建成功
    del test_log.txt
) else (
    echo 测试日志创建失败
)

echo.
echo 10. 检查Excel相关组件...
reg query "HKEY_CLASSES_ROOT\Excel.Application" >nul 2>&1
if %errorlevel% equ 0 (
    echo Excel组件注册正常
) else (
    echo 警告: Excel组件未注册
)

echo.
echo ========================================
echo 诊断完成
echo ========================================
echo.
echo 如果程序仍然无法启动，请：
echo 1. 检查上述信息中的警告项
echo 2. 尝试以管理员身份运行
echo 3. 检查杀毒软件是否阻止程序运行
echo 4. 查看Windows事件查看器中的错误信息
echo.
pause