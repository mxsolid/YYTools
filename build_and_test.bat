@echo off
chcp 65001 >nul
setlocal ENABLEDELAYEDEXPANSION
echo ==^> 开始编译 (UTF-8)

set CONFIG=Release
set PLATFORM="Any CPU"

set MSBUILD=msbuild
where %MSBUILD% >nul 2>nul
if errorlevel 1 (
  echo 未找到 msbuild。请在“开发者命令提示符 for VS”中运行本脚本。
  exit /b 1
)

for %%P in ("YYTools\\YYTools.csproj" "TestApp\\TestApp.csproj") do (
  if exist %%~fP (
    echo -- 编译: %%~fP
    %MSBUILD% %%~fP -t:Rebuild -p:Configuration=%CONFIG% -p:Platform=%PLATFORM% -verbosity:minimal
    if errorlevel 1 (
      echo ✗ 编译失败: %%~fP
      exit /b 2
    )
  ) else (
    echo -- 跳过: %%~fP (未找到)
  )
)

echo ==^> 编译完成
exit /b 0


@echo off
chcp 65001 >nul
echo ========================================
echo YY工具编译测试脚本
echo 版本: 3.2 (性能优化版)
echo 时间: %date% %time%
echo ========================================
echo.

echo [1/4] 清理旧的编译文件...
if exist "YYTools\bin" rmdir /s /q "YYTools\bin"
if exist "YYTools\obj" rmdir /s /q "YYTools\obj"
echo 清理完成
echo.

echo [2/4] 检查.NET Framework版本...
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Version
if %errorlevel% neq 0 (
    echo 错误: 未检测到.NET Framework 4.0或更高版本
    echo 请安装.NET Framework 4.8或更高版本
    pause
    exit /b 1
)
echo .NET Framework检查通过
echo.

echo [3/4] 编译项目...
cd YYTools
msbuild YYTools.csproj /p:Configuration=Release /p:Platform="Any CPU" /verbosity:minimal
if %errorlevel% neq 0 (
    echo.
    echo 错误: 编译失败！
    echo 请检查代码中的错误
    cd ..
    pause
    exit /b 1
)
cd ..
echo 编译成功！
echo.

echo [4/4] 运行编译测试...
if exist "YYTools\bin\Release\YYTools.exe" (
    echo 编译测试通过！
    echo 可执行文件位置: YYTools\bin\Release\YYTools.exe
    echo.
    echo 是否要运行程序进行测试？(Y/N)
    set /p choice=
    if /i "%choice%"=="Y" (
        echo 启动程序...
        start "" "YYTools\bin\Release\YYTools.exe"
    )
) else (
    echo 错误: 未找到编译后的可执行文件
    pause
    exit /b 1
)

echo.
echo ========================================
echo 编译测试完成！
echo ========================================
pause