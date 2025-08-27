@echo off
chcp 65001 >nul
echo ========================================
echo YY工具 v3.0 一键打包脚本
echo ========================================
echo.

:: 检查.NET Framework版本
echo [1/6] 检查运行环境...
dotnet --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未检测到.NET Framework，请先安装.NET Framework 4.8或更高版本
    echo 下载地址: https://dotnet.microsoft.com/download
    pause
    exit /b 1
)

:: 检查MSBuild
echo [2/6] 检查编译工具...
where msbuild >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未检测到MSBuild，请安装Visual Studio或Build Tools
    echo 下载地址: https://visualstudio.microsoft.com/downloads/
    pause
    exit /b 1
)

:: 清理旧文件
echo [3/6] 清理旧文件...
if exist "bin\Release" rmdir /s /q "bin\Release"
if exist "obj\Release" rmdir /s /q "obj\Release"
if exist "YYTools_v3.0" rmdir /s /q "YYTools_v3.0"

:: 编译项目
echo [4/6] 编译项目...
msbuild YYTools\YYTools.csproj /p:Configuration=Release /p:Platform="Any CPU" /p:OutputPath=bin\Release\ /verbosity:minimal
if %errorlevel% neq 0 (
    echo 错误: 编译失败，请检查代码错误
    pause
    exit /b 1
)

:: 创建发布目录
echo [5/6] 创建发布包...
mkdir "YYTools_v3.0"
mkdir "YYTools_v3.0\YYTools"
mkdir "YYTools_v3.0\YYTools\Logs"

:: 复制文件
copy "YYTools\bin\Release\YYTools.exe" "YYTools_v3.0\YYTools\"
copy "YYTools\bin\Release\YYTools.dll" "YYTools_v3.0\YYTools\"
copy "YYTools\YYProgram.ico" "YYTools_v3.0\YYTools\"
copy "README.md" "YYTools_v3.0\"
copy "使用指南.md" "YYTools_v3.0\"
copy "INSTALL.md" "YYTools_v3.0\"

:: 创建安装脚本
echo @echo off > "YYTools_v3.0\安装.bat"
echo chcp 65001 ^>nul >> "YYTools_v3.0\安装.bat"
echo echo 正在安装YY工具... >> "YYTools_v3.0\安装.bat"
echo echo. >> "YYTools_v3.0\安装.bat"
echo echo 请以管理员身份运行此脚本 >> "YYTools_v3.0\安装.bat"
echo echo. >> "YYTools_v3.0\安装.bat"
echo pause >> "YYTools_v3.0\安装.bat"

:: 创建使用说明
echo YY工具 v3.0 使用说明 > "YYTools_v3.0\使用说明.txt"
echo ======================== >> "YYTools_v3.0\使用说明.txt"
echo. >> "YYTools_v3.0\使用说明.txt"
echo 1. 双击 YYTools.exe 运行程序 >> "YYTools_v3.0\使用说明.txt"
echo 2. 首次运行可能需要安装.NET Framework 4.8 >> "YYTools_v3.0\使用说明.txt"
echo 3. 程序会自动检测Excel/WPS文件 >> "YYTools_v3.0\使用说明.txt"
echo 4. 详细使用说明请查看 使用指南.md >> "YYTools_v3.0\使用说明.txt"
echo. >> "YYTools_v3.0\使用说明.txt"
echo 技术支持: oyxo@qq.com >> "YYTools_v3.0\使用说明.txt"

:: 创建自解压包
echo [6/6] 创建自解压包...
powershell -Command "Compress-Archive -Path 'YYTools_v3.0' -DestinationPath 'YYTools_v3.0.zip' -Force"

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo.
echo 生成的文件:
echo - YYTools_v3.0\  (程序目录)
echo - YYTools_v3.0.zip (压缩包)
echo.
echo 使用说明:
echo 1. 将 YYTools_v3.0 文件夹复制到目标机器
echo 2. 确保目标机器安装了.NET Framework 4.8
echo 3. 双击 YYTools.exe 运行程序
echo.
echo 如果没有.NET Framework，请先安装:
echo https://dotnet.microsoft.com/download/dotnet-framework/net48
echo.
pause