@echo off
chcp 65001 >nul
echo ========================================
echo YY工具 v2.6 测试构建脚本
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

:: 设置项目目录
set PROJECT_DIR=YYTools

echo 正在清理旧的构建文件...
if exist "%PROJECT_DIR%\bin" rmdir /s /q "%PROJECT_DIR%\bin"
if exist "%PROJECT_DIR%\obj" rmdir /s /q "%PROJECT_DIR%\obj"

echo.
echo 正在检查源代码文件...
if not exist "%PROJECT_DIR%\MatchForm.cs" (
    echo 错误：找不到MatchForm.cs文件
    pause
    exit /b 1
)

if not exist "%PROJECT_DIR%\SmartColumnService.cs" (
    echo 错误：找不到SmartColumnService.cs文件
    pause
    exit /b 1
)

if not exist "%PROJECT_DIR%\DataModels.cs" (
    echo 错误：找不到DataModels.cs文件
    pause
    exit /b 1
)

if not exist "%PROJECT_DIR%\AppSettings.cs" (
    echo 错误：找不到AppSettings.cs文件
    pause
    exit /b 1
)

if not exist "%PROJECT_DIR%\MatchService.cs" (
    echo 错误：找不到MatchService.cs文件
    pause
    exit /b 1
)

if not exist "%PROJECT_DIR%\ExcelHelper.cs" (
    echo 错误：找不到ExcelHelper.cs文件
    pause
    exit /b 1
)

echo 源代码文件检查完成

echo.
echo 正在构建项目（Debug模式）...
cd %PROJECT_DIR%
msbuild YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /p:OutputPath=bin\Debug\ /verbosity:minimal
if %errorlevel% neq 0 (
    echo.
    echo Debug构建失败！请检查错误信息
    cd ..
    pause
    exit /b 1
)

echo.
echo Debug构建成功！正在构建Release版本...
msbuild YYTools.csproj /p:Configuration=Release /p:Platform="Any CPU" /p:OutputPath=bin\Release\ /verbosity:minimal
if %errorlevel% neq 0 (
    echo.
    echo Release构建失败！请检查错误信息
    cd ..
    pause
    exit /b 1
)

echo.
echo Release构建成功！正在检查输出文件...
cd ..

if not exist "%PROJECT_DIR%\bin\Release\YYTools.exe" (
    echo 错误：Release构建后找不到YYTools.exe
    pause
    exit /b 1
)

if not exist "%PROJECT_DIR%\bin\Release\YYTools.dll" (
    echo 错误：Release构建后找不到YYTools.dll
    pause
    exit /b 1
)

echo.
echo ========================================
echo 测试构建成功！
echo ========================================
echo.
echo 输出文件：
echo - %PROJECT_DIR%\bin\Debug\YYTools.exe
echo - %PROJECT_DIR%\bin\Debug\YYTools.dll
echo - %PROJECT_DIR%\bin\Release\YYTools.exe
echo - %PROJECT_DIR%\bin\Release\YYTools.dll
echo.
echo 文件大小：
for %%A in ("%PROJECT_DIR%\bin\Release\YYTools.exe") do echo YYTools.exe: %%~zA 字节
for %%A in ("%PROJECT_DIR%\bin\Release\YYTools.dll") do echo YYTools.dll: %%~zA 字节
echo.
echo 现在可以运行build_release.bat来创建完整的发布包
echo.
pause