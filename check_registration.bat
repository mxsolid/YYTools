@echo off
chcp 65001 >nul

echo ========================================
echo YYTools COM注册检查
echo ========================================

echo 1. 检查DLL文件...
if exist "bin\Debug\YYTools.dll" (
    echo ✓ DLL文件存在: bin\Debug\YYTools.dll
    dir "bin\Debug\YYTools.dll" | find "YYTools.dll"
) else (
    echo ✗ DLL文件不存在
    echo 请先运行 install_admin.bat 编译生成DLL
    pause
    exit /b 1
)

echo.
echo 2. 检查注册表项...
echo 查找 YYTools.ExcelAddin...
reg query "HKEY_CLASSES_ROOT\YYTools.ExcelAddin" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo ✓ 注册表项存在: HKEY_CLASSES_ROOT\YYTools.ExcelAddin
    reg query "HKEY_CLASSES_ROOT\YYTools.ExcelAddin\CLSID" 2>nul
) else (
    echo ✗ 注册表项不存在
)

echo.
echo 查找 CLSID...
reg query "HKEY_CLASSES_ROOT\CLSID\{12345678-1234-5678-9ABC-123456789ABC}" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo ✓ CLSID注册表项存在
    reg query "HKEY_CLASSES_ROOT\CLSID\{12345678-1234-5678-9ABC-123456789ABC}\InprocServer32" 2>nul
) else (
    echo ✗ CLSID注册表项不存在
)

echo.
echo 3. 尝试创建COM对象测试...
powershell -Command "try { $obj = New-Object -ComObject 'YYTools.ExcelAddin'; if ($obj) { Write-Host '✓ COM对象创建成功'; $obj = $null } else { Write-Host '✗ COM对象创建失败' } } catch { Write-Host '✗ COM对象创建异常:' $_.Exception.Message }"

echo.
echo 4. 检查.NET Framework注册...
set FRAMEWORK_PATH=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
if not exist "%FRAMEWORK_PATH%\RegAsm.exe" (
    set FRAMEWORK_PATH=%WINDIR%\Microsoft.NET\Framework\v4.0.30319
)

echo 使用RegAsm路径: %FRAMEWORK_PATH%\RegAsm.exe
if exist "%FRAMEWORK_PATH%\RegAsm.exe" (
    echo ✓ RegAsm.exe 存在
) else (
    echo ✗ RegAsm.exe 不存在
)

echo.
echo ========================================
echo 检查完成
echo ========================================

pause 