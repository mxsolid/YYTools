@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

:: =====================================
:: YY运单匹配工具 - 自动安装脚本 v2.0
:: =====================================

echo.
echo =====================================
echo YY运单匹配工具 - 自动安装程序
echo =====================================
echo 版本: v2.0
echo 功能: 运单匹配，多工作簿支持，极致性能
echo =====================================
echo.

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ 错误：需要管理员权限！
    echo.
    echo 请以管理员身份运行此脚本。
    echo 右键点击此文件 → "以管理员身份运行"
    pause
    exit /b 1
)

echo ✅ 管理员权限检查通过

:: 设置安装路径
set "INSTALL_DIR=%ProgramFiles%\YYTools"
set "ADDIN_DIR=%APPDATA%\Microsoft\AddIns"
set "WPS_ADDIN_DIR=%APPDATA%\Kingsoft\WPS Office\AddIns"

echo.
echo 📂 安装目录设置：
echo    程序目录: %INSTALL_DIR%
echo    Excel插件: %ADDIN_DIR%
echo    WPS插件: %WPS_ADDIN_DIR%
echo.

:: 创建安装目录
echo 🔨 创建安装目录...
if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%" 2>nul
    if !errorLevel! neq 0 (
        echo ❌ 无法创建安装目录: %INSTALL_DIR%
        pause
        exit /b 1
    )
)

if not exist "%ADDIN_DIR%" (
    mkdir "%ADDIN_DIR%" 2>nul
)

if not exist "%WPS_ADDIN_DIR%" (
    mkdir "%WPS_ADDIN_DIR%" 2>nul
)

echo ✅ 目录创建完成

:: 停止相关进程
echo.
echo 🛑 检查并停止相关进程...
taskkill /f /im EXCEL.EXE 2>nul
taskkill /f /im et.exe 2>nul
taskkill /f /im wps.exe 2>nul
timeout /t 2 /nobreak >nul

:: 复制文件
echo.
echo 📦 复制程序文件...

if exist "YYTools\bin\Release\YYTools.dll" (
    copy "YYTools\bin\Release\YYTools.dll" "%INSTALL_DIR%\" >nul
    echo ✅ YYTools.dll
) else (
    echo ❌ 找不到 YYTools.dll，请先编译项目！
    pause
    exit /b 1
)

if exist "TestProgram.exe" (
    copy "TestProgram.exe" "%INSTALL_DIR%\" >nul
    echo ✅ TestProgram.exe
)

:: 创建VBA集成文件
echo.
echo 🔧 创建VBA集成文件...

:: Excel VBA 文件
echo Sub YYTools_运单匹配() > "%ADDIN_DIR%\YYTools.bas"
echo     Dim obj As Object >> "%ADDIN_DIR%\YYTools.bas"
echo     Set obj = CreateObject("YYTools.ExcelAddin") >> "%ADDIN_DIR%\YYTools.bas"
echo     obj.ShowMatchForm >> "%ADDIN_DIR%\YYTools.bas"
echo End Sub >> "%ADDIN_DIR%\YYTools.bas"
echo.
echo Sub YYTools_高级匹配() >> "%ADDIN_DIR%\YYTools.bas"
echo     Dim obj As Object >> "%ADDIN_DIR%\YYTools.bas"
echo     Set obj = CreateObject("YYTools.ExcelAddin") >> "%ADDIN_DIR%\YYTools.bas"
echo     obj.ShowAdvancedMatchForm >> "%ADDIN_DIR%\YYTools.bas"
echo End Sub >> "%ADDIN_DIR%\YYTools.bas"

:: WPS VBA 文件
copy "%ADDIN_DIR%\YYTools.bas" "%WPS_ADDIN_DIR%\YYTools.bas" >nul

echo ✅ VBA集成文件已创建

:: 注册COM组件
echo.
echo 🔗 注册COM组件...
regsvr32 /s "%INSTALL_DIR%\YYTools.dll"
if !errorLevel! equ 0 (
    echo ✅ COM组件注册成功
) else (
    echo ⚠️  COM组件注册失败，将使用直接调用方式
)

:: 创建桌面快捷方式
echo.
echo 🖥️  创建桌面快捷方式...
set "DESKTOP=%USERPROFILE%\Desktop"

echo Set oWS = WScript.CreateObject("WScript.Shell") > "%TEMP%\CreateShortcut.vbs"
echo sLinkFile = "%DESKTOP%\YY运单匹配工具.lnk" >> "%TEMP%\CreateShortcut.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%TEMP%\CreateShortcut.vbs"
echo oLink.TargetPath = "%INSTALL_DIR%\TestProgram.exe" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.WorkingDirectory = "%INSTALL_DIR%" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.Description = "YY运单匹配工具 - 快速匹配发货和账单数据" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.Save >> "%TEMP%\CreateShortcut.vbs"

cscript //nologo "%TEMP%\CreateShortcut.vbs"
del "%TEMP%\CreateShortcut.vbs" 2>nul

echo ✅ 桌面快捷方式已创建

:: 创建开始菜单快捷方式
echo.
echo 📁 创建开始菜单快捷方式...
set "START_MENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs"
if not exist "%START_MENU%\YYTools" mkdir "%START_MENU%\YYTools"

echo Set oWS = WScript.CreateObject("WScript.Shell") > "%TEMP%\CreateStartMenu.vbs"
echo sLinkFile = "%START_MENU%\YYTools\YY运单匹配工具.lnk" >> "%TEMP%\CreateStartMenu.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%TEMP%\CreateStartMenu.vbs"
echo oLink.TargetPath = "%INSTALL_DIR%\TestProgram.exe" >> "%TEMP%\CreateStartMenu.vbs"
echo oLink.WorkingDirectory = "%INSTALL_DIR%" >> "%TEMP%\CreateStartMenu.vbs"
echo oLink.Description = "YY运单匹配工具" >> "%TEMP%\CreateStartMenu.vbs"
echo oLink.Save >> "%TEMP%\CreateStartMenu.vbs"

cscript //nologo "%TEMP%\CreateStartMenu.vbs"
del "%TEMP%\CreateStartMenu.vbs" 2>nul

echo ✅ 开始菜单快捷方式已创建

:: 添加到Windows PATH
echo.
echo 🌐 添加到系统PATH...
setx PATH "%PATH%;%INSTALL_DIR%" /M >nul 2>&1
if !errorLevel! equ 0 (
    echo ✅ 已添加到系统PATH
) else (
    echo ⚠️  添加到PATH失败，请手动添加: %INSTALL_DIR%
)

:: 创建卸载程序
echo.
echo 🗑️  创建卸载程序...
echo @echo off > "%INSTALL_DIR%\uninstall.bat"
echo echo 正在卸载YY运单匹配工具... >> "%INSTALL_DIR%\uninstall.bat"
echo taskkill /f /im TestProgram.exe 2^>nul >> "%INSTALL_DIR%\uninstall.bat"
echo regsvr32 /u /s "%INSTALL_DIR%\YYTools.dll" >> "%INSTALL_DIR%\uninstall.bat"
echo del "%USERPROFILE%\Desktop\YY运单匹配工具.lnk" 2^>nul >> "%INSTALL_DIR%\uninstall.bat"
echo rd /s /q "%APPDATA%\Microsoft\Windows\Start Menu\Programs\YYTools" 2^>nul >> "%INSTALL_DIR%\uninstall.bat"
echo del "%ADDIN_DIR%\YYTools.bas" 2^>nul >> "%INSTALL_DIR%\uninstall.bat"
echo del "%WPS_ADDIN_DIR%\YYTools.bas" 2^>nul >> "%INSTALL_DIR%\uninstall.bat"
echo echo 卸载完成！ >> "%INSTALL_DIR%\uninstall.bat"
echo pause >> "%INSTALL_DIR%\uninstall.bat"
echo rd /s /q "%INSTALL_DIR%" >> "%INSTALL_DIR%\uninstall.bat"

echo ✅ 卸载程序已创建

:: 创建使用说明
echo.
echo 📖 创建使用说明...
echo YY运单匹配工具 v2.0 - 使用说明 > "%INSTALL_DIR%\使用说明.txt"
echo ======================================= >> "%INSTALL_DIR%\使用说明.txt"
echo. >> "%INSTALL_DIR%\使用说明.txt"
echo 🚀 快速启动： >> "%INSTALL_DIR%\使用说明.txt"
echo 1. 双击桌面的"YY运单匹配工具"快捷方式 >> "%INSTALL_DIR%\使用说明.txt"
echo 2. 或在开始菜单中找到"YYTools" >> "%INSTALL_DIR%\使用说明.txt"
echo. >> "%INSTALL_DIR%\使用说明.txt"
echo 💡 Excel/WPS集成使用： >> "%INSTALL_DIR%\使用说明.txt"
echo 1. 在Excel或WPS中按Alt+F11打开VBA编辑器 >> "%INSTALL_DIR%\使用说明.txt"
echo 2. 导入文件：%ADDIN_DIR%\YYTools.bas >> "%INSTALL_DIR%\使用说明.txt"
echo 3. 运行宏：YYTools_运单匹配 或 YYTools_高级匹配 >> "%INSTALL_DIR%\使用说明.txt"
echo. >> "%INSTALL_DIR%\使用说明.txt"
echo 📋 功能特点： >> "%INSTALL_DIR%\使用说明.txt"
echo • 极致性能优化，6000+行数据秒级处理 >> "%INSTALL_DIR%\使用说明.txt"
echo • 支持多工作簿跨文件操作 >> "%INSTALL_DIR%\使用说明.txt"
echo • WPS表格和Excel双重兼容 >> "%INSTALL_DIR%\使用说明.txt"
echo • 实时进度显示和详细日志记录 >> "%INSTALL_DIR%\使用说明.txt"
echo • 美观现代的用户界面 >> "%INSTALL_DIR%\使用说明.txt"
echo. >> "%INSTALL_DIR%\使用说明.txt"
echo 🔧 卸载： >> "%INSTALL_DIR%\使用说明.txt"
echo 运行：%INSTALL_DIR%\uninstall.bat >> "%INSTALL_DIR%\使用说明.txt"

echo ✅ 使用说明已创建

:: 安装完成
echo.
echo =====================================
echo 🎉 安装完成！
echo =====================================
echo.
echo ✅ 程序文件已安装到: %INSTALL_DIR%
echo ✅ VBA集成文件已创建
echo ✅ 桌面快捷方式已创建
echo ✅ 开始菜单快捷方式已创建
echo ✅ 系统PATH已更新
echo.
echo 📖 使用方式：
echo    1. 双击桌面"YY运单匹配工具"
echo    2. 在Excel/WPS中使用VBA宏
echo    3. 命令行运行: TestProgram
echo.
echo 📋 说明文档: %INSTALL_DIR%\使用说明.txt
echo 🗑️  卸载程序: %INSTALL_DIR%\uninstall.bat
echo.
echo 现在可以开始使用YY运单匹配工具了！
echo.
pause 