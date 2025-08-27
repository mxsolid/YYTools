@echo off
chcp 65001 >nul
echo ========================================
echo YY工具 v2.6 一键打包脚本
echo ========================================
echo.

:: 检查MSBuild是否可用
where msbuild >nul 2>nul
if %errorlevel% neq 0 (
    echo 错误：未找到MSBuild，请确保已安装Visual Studio或Build Tools
    echo 请安装Visual Studio 2019/2022或Microsoft Build Tools
    pause
    exit /b 1
)

:: 设置版本号
set VERSION=2.6.0.0
set RELEASE_DIR=Release_v%VERSION%
set PROJECT_DIR=YYTools

echo 正在清理旧的构建文件...
if exist "%PROJECT_DIR%\bin" rmdir /s /q "%PROJECT_DIR%\bin"
if exist "%PROJECT_DIR%\obj" rmdir /s /q "%PROJECT_DIR%\obj"
if exist "%RELEASE_DIR%" rmdir /s /q "%RELEASE_DIR%"

echo.
echo 正在构建项目...
cd %PROJECT_DIR%
msbuild YYTools.csproj /p:Configuration=Release /p:Platform="Any CPU" /p:OutputPath=bin\Release\ /verbosity:minimal
if %errorlevel% neq 0 (
    echo.
    echo 构建失败！请检查错误信息
    cd ..
    pause
    exit /b 1
)

echo.
echo 构建成功！正在创建发布包...
cd ..

:: 创建发布目录
mkdir "%RELEASE_DIR%"
mkdir "%RELEASE_DIR%\YYTools"

:: 复制可执行文件
echo 复制可执行文件...
copy "%PROJECT_DIR%\bin\Release\YYTools.exe" "%RELEASE_DIR%\YYTools\"
copy "%PROJECT_DIR%\bin\Release\YYTools.dll" "%RELEASE_DIR%\YYTools\"
copy "%PROJECT_DIR%\bin\Release\*.dll" "%RELEASE_DIR%\YYTools\"

:: 复制配置文件
echo 复制配置文件...
copy "%PROJECT_DIR%\app.config" "%RELEASE_DIR%\YYTools\"
copy "%PROJECT_DIR%\YYProgram.ico" "%RELEASE_DIR%\YYTools\"

:: 复制文档
echo 复制文档...
copy "README.md" "%RELEASE_DIR%\"
copy "使用指南.md" "%RELEASE_DIR%\"
copy "项目完成总结.md" "%RELEASE_DIR%\"

:: 创建安装脚本
echo 创建安装脚本...
(
echo @echo off
echo chcp 65001 ^>nul
echo echo ========================================
echo echo YY工具 v2.6 安装脚本
echo echo ========================================
echo echo.
echo echo 正在安装YY工具...
echo.
echo :: 检查管理员权限
echo net session ^>nul 2^>^&1
echo if %%errorlevel%% neq 0 ^(
echo     echo 需要管理员权限来安装COM组件
echo     echo 请右键选择"以管理员身份运行"
echo     pause
echo     exit /b 1
echo ^)
echo.
echo :: 注册COM组件
echo echo 正在注册COM组件...
echo regsvr32 /s "%~dp0YYTools\YYTools.dll"
echo if %%errorlevel%% neq 0 ^(
echo     echo COM组件注册失败，但程序仍可运行
echo     echo.
echo ^)
echo.
echo :: 创建桌面快捷方式
echo echo 正在创建桌面快捷方式...
echo set DESKTOP=%%USERPROFILE%%\Desktop
echo if not exist "%%DESKTOP%%" set DESKTOP=%%USERPROFILE%%\桌面
echo.
echo echo @echo off ^> "%%DESKTOP%%\YY工具.bat"
echo echo cd /d "%%~dp0YYTools" ^>^> "%%DESKTOP%%\YY工具.bat"
echo echo start YYTools.exe ^>^> "%%DESKTOP%%\YY工具.bat"
echo.
echo echo 安装完成！
echo echo 桌面快捷方式已创建
echo echo.
echo pause
) > "%RELEASE_DIR%\安装.bat"

:: 创建卸载脚本
echo 创建卸载脚本...
(
echo @echo off
echo chcp 65001 ^>nul
echo echo ========================================
echo echo YY工具 v2.6 卸载脚本
echo echo ========================================
echo echo.
echo echo 正在卸载YY工具...
echo.
echo :: 检查管理员权限
echo net session ^>nul 2^>^&1
echo if %%errorlevel%% neq 0 ^(
echo     echo 需要管理员权限来卸载COM组件
echo     echo 请右键选择"以管理员身份运行"
echo     pause
echo     exit /b 1
echo ^)
echo.
echo :: 注销COM组件
echo echo 正在注销COM组件...
echo regsvr32 /u /s "%~dp0YYTools\YYTools.dll"
echo.
echo :: 删除桌面快捷方式
echo echo 正在删除桌面快捷方式...
echo set DESKTOP=%%USERPROFILE%%\Desktop
echo if not exist "%%DESKTOP%%" set DESKTOP=%%USERPROFILE%%\桌面
echo if exist "%%DESKTOP%%\YY工具.bat" del "%%DESKTOP%%\YY工具.bat"
echo.
echo echo 卸载完成！
echo echo.
echo pause
) > "%RELEASE_DIR%\卸载.bat"

:: 创建使用说明
echo 创建使用说明...
(
echo YY工具 v2.6 使用说明
echo ========================================
echo.
echo 安装方法：
echo 1. 以管理员身份运行"安装.bat"
echo 2. 等待安装完成
echo 3. 桌面会创建快捷方式
echo.
echo 使用方法：
echo 1. 双击桌面快捷方式或YYTools.exe
echo 2. 在Excel/WPS中加载插件
echo 3. 配置运单匹配参数
echo 4. 开始匹配任务
echo.
echo 卸载方法：
echo 1. 以管理员身份运行"卸载.bat"
echo 2. 等待卸载完成
echo.
echo 注意事项：
echo - 首次使用需要以管理员身份安装
echo - 支持Excel 2010及以上版本
echo - 支持WPS Office
echo - 建议在匹配前备份数据
echo.
echo 技术支持：
echo 作者：皮皮熊
echo 邮箱：oyxo@qq.com
echo.
echo 版本：v2.6.0.0
echo 更新日期：2025年
) > "%RELEASE_DIR%\使用说明.txt"

:: 创建版本信息文件
echo 创建版本信息文件...
(
echo YY工具版本信息
echo ========================================
echo 版本号：%VERSION%
echo 构建日期：%date% %time%
echo 目标框架：.NET Framework 4.8
echo 平台：Windows x86/x64
echo.
echo 更新内容：
echo - 智能列选择功能
echo - 列搜索和预览功能
echo - 性能优化，支持大数据量
echo - 完善的错误处理和日志记录
echo - 界面布局优化
echo - 版本号规范化
echo.
echo 系统要求：
echo - Windows 7 SP1 及以上
echo - .NET Framework 4.8
echo - Excel 2010 或 WPS Office
echo.
echo 文件清单：
echo - YYTools.exe (主程序)
echo - YYTools.dll (核心库)
echo - app.config (配置文件)
echo - YYProgram.ico (程序图标)
echo - 安装.bat (安装脚本)
echo - 卸载.bat (卸载脚本)
echo - 使用说明.txt (使用说明)
) > "%RELEASE_DIR%\版本信息.txt"

:: 创建ZIP压缩包
echo 创建压缩包...
powershell -command "Compress-Archive -Path '%RELEASE_DIR%' -DestinationPath '%RELEASE_DIR%.zip' -Force"

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo.
echo 发布目录：%RELEASE_DIR%
echo 压缩包：%RELEASE_DIR%.zip
echo.
echo 文件清单：
dir /b "%RELEASE_DIR%"
echo.
echo 压缩包大小：
for %%A in ("%RELEASE_DIR%.zip") do echo %%~zA 字节
echo.
echo 按任意键打开发布目录...
pause >nul
explorer "%RELEASE_DIR%"