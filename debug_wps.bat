@echo off
chcp 65001 >nul

echo ========================================
echo WPS 调试检查脚本
echo ========================================

echo 1. 检查正在运行的Office/WPS进程...
echo.
echo 查找WPS进程:
tasklist | findstr /i "wps"
echo.
echo 查找Excel进程:
tasklist | findstr /i "excel"
echo.

echo 2. 检查WPS ProgID注册情况...
echo.
echo 检查 Ket.Application:
reg query "HKEY_CLASSES_ROOT\Ket.Application" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo ✓ Ket.Application 已注册
    reg query "HKEY_CLASSES_ROOT\Ket.Application\CLSID" 2>nul
) else (
    echo ✗ Ket.Application 未注册
)

echo.
echo 检查 WPS.Application:
reg query "HKEY_CLASSES_ROOT\WPS.Application" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo ✓ WPS.Application 已注册
    reg query "HKEY_CLASSES_ROOT\WPS.Application\CLSID" 2>nul
) else (
    echo ✗ WPS.Application 未注册
)

echo.
echo 检查 ET.Application:
reg query "HKEY_CLASSES_ROOT\ET.Application" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo ✓ ET.Application 已注册
    reg query "HKEY_CLASSES_ROOT\ET.Application\CLSID" 2>nul
) else (
    echo ✗ ET.Application 未注册
)

echo.
echo 3. 尝试通过PowerShell连接WPS...
powershell -Command "try { $wps = New-Object -ComObject 'Ket.Application'; if ($wps) { Write-Host '✓ 成功连接到Ket.Application'; $name = $wps.Name; Write-Host '应用程序名称:' $name; $wps = $null } else { Write-Host '✗ 无法连接到Ket.Application' } } catch { Write-Host '✗ Ket.Application连接异常:' $_.Exception.Message }"

powershell -Command "try { $wps = New-Object -ComObject 'WPS.Application'; if ($wps) { Write-Host '✓ 成功连接到WPS.Application'; $name = $wps.Name; Write-Host '应用程序名称:' $name; $wps = $null } else { Write-Host '✗ 无法连接到WPS.Application' } } catch { Write-Host '✗ WPS.Application连接异常:' $_.Exception.Message }"

powershell -Command "try { $et = New-Object -ComObject 'ET.Application'; if ($et) { Write-Host '✓ 成功连接到ET.Application'; $name = $et.Name; Write-Host '应用程序名称:' $name; $et = $null } else { Write-Host '✗ 无法连接到ET.Application' } } catch { Write-Host '✗ ET.Application连接异常:' $_.Exception.Message }"

echo.
echo 4. 测试YYTools在不同应用程序中的表现...
powershell -Command "try { $yytools = New-Object -ComObject 'YYTools.ExcelAddin'; $info = $yytools.GetDetailedApplicationInfo(); Write-Host '详细信息:'; Write-Host $info; $result = $yytools.InstallMenu(); Write-Host '菜单安装结果:' $result; $yytools = $null } catch { Write-Host '✗ YYTools测试异常:' $_.Exception.Message }"

echo.
echo 5. 检查WPS版本信息...
if exist "C:\Program Files (x86)\Kingsoft\WPS Office\*\office6\et.exe" (
    echo ✓ 发现WPS Office安装 (Program Files x86)
    dir "C:\Program Files (x86)\Kingsoft\WPS Office\" | find "<DIR>"
)

if exist "C:\Program Files\Kingsoft\WPS Office\*\office6\et.exe" (
    echo ✓ 发现WPS Office安装 (Program Files)
    dir "C:\Program Files\Kingsoft\WPS Office\" | find "<DIR>"
)

echo.
echo ========================================
echo WPS调试完成
echo ========================================

pause 