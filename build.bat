@echo off
chcp 65001 >nul
echo =====================================
echo YY运单匹配工具 编译脚本 v2.1
echo =====================================
echo.

REM 设置MSBuild路径（可能需要根据实际安装路径调整）
set MSBUILD_PATH="C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
if not exist %MSBUILD_PATH% (
    set MSBUILD_PATH="C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
)
if not exist %MSBUILD_PATH% (
    set MSBUILD_PATH="C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
)
if not exist %MSBUILD_PATH% (
    set MSBUILD_PATH="C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe"
)
if not exist %MSBUILD_PATH% (
    set MSBUILD_PATH="C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
)

REM 检查MSBuild是否存在
if not exist %MSBUILD_PATH% (
    echo 错误：找不到MSBuild.exe
    echo 请确保已安装Visual Studio 2019/2022 或 .NET Framework SDK
    echo 或者手动设置MSBUILD_PATH变量
    pause
    exit /b 1
)

echo 找到MSBuild：%MSBUILD_PATH%
echo.

REM 清理之前的编译结果
echo 正在清理之前的编译结果...
if exist "YYTools\bin" rmdir /s /q "YYTools\bin"
if exist "YYTools\obj" rmdir /s /q "YYTools\obj"
echo 清理完成
echo.

REM 编译Debug版本
echo 正在编译Debug版本...
%MSBUILD_PATH% YYTools\YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
if %errorlevel% neq 0 (
    echo 编译Debug版本失败！
    echo 尝试使用解决方案文件...
    if exist "YYTools.sln" (
        %MSBUILD_PATH% YYTools.sln /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
        if %errorlevel% neq 0 (
            echo 解决方案编译也失败！请检查项目配置。
            pause
            exit /b 1
        )
    ) else (
        echo 没有找到解决方案文件，编译失败！
        pause
        exit /b 1
    )
)

echo.
echo =====================================
echo 编译成功！
echo =====================================

REM 检查输出文件
if exist "YYTools\bin\Debug\YYTools.dll" (
    echo Debug版本编译完成
    echo 输出路径: YYTools\bin\Debug\YYTools.dll
    dir "YYTools\bin\Debug\YYTools.dll" | find "YYTools.dll"
) else (
    echo 警告：找不到输出的DLL文件
)

echo.
echo 编译后续步骤：
echo 1. 注册COM组件: 以管理员身份运行 install_admin.bat
echo 2. 测试功能: 运行 bin\Debug\YYToolsTest.exe
echo 3. 在WPS/Excel中调用: CreateObject("YYTools.ExcelAddin").InstallMenu()
echo.

pause 