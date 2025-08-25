@echo off
chcp 65001 >nul
echo =====================================
echo YY运单匹配工具 v1.5 最终发布脚本
echo =====================================
echo.

set MSBUILD_PATH="D:\Develop\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
set CSC_PATH="C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

echo 1. 清理并编译项目...
%MSBUILD_PATH% YYTools\YYTools.csproj /p:Configuration=Release /p:Platform=AnyCPU /verbosity:minimal
if errorlevel 1 goto :build_fail

echo 2. 编译测试程序...
%CSC_PATH% /target:winexe /r:YYTools\bin\Release\YYTools.dll /r:System.Windows.Forms.dll TestProgram.cs
if errorlevel 1 goto :build_fail

echo 3. 创建发布目录...
if exist "YYTools_Final" rmdir /s /q "YYTools_Final"
mkdir "YYTools_Final"
copy "YYTools\bin\Release\YYTools.dll" "YYTools_Final\" >nul
copy "YYTools\bin\Release\YYTools.pdb" "YYTools_Final\" >nul
copy "TestProgram.exe" "YYTools_Final\" >nul

rem 4. 生成安装脚本
(
echo @echo off

echo chcp 65001 ^>nul

echo echo =====================================
echo echo YY运单匹配工具 v1.5 安装程序
echo echo =====================================
echo echo.

echo setlocal enableextensions
echo set INSTALL_DIR=%%USERPROFILE%%\Documents\YYTools
echo if not exist "%%INSTALL_DIR%%" mkdir "%%INSTALL_DIR%%"

echo echo 复制文件到Documents目录...
echo copy /y "YYTools.dll" "%%INSTALL_DIR%%\" ^>nul

echo copy /y "TestProgram.exe" "%%INSTALL_DIR%%\" ^>nul

echo echo.

echo echo 注册COM组件...

echo set REGASM64=%%SystemRoot%%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

echo set REGASM32=%%SystemRoot%%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe

echo echo 尝试64位注册...

echo "%%REGASM64%%" "%%INSTALL_DIR%%\YYTools.dll" /codebase ^>nul

echo if errorlevel 1 (
	echo echo 64位注册可能部分失败，尝试32位...
	echo "%%REGASM32%%" "%%INSTALL_DIR%%\YYTools.dll" /codebase ^>nul
)

echo rem 写入WPS/Excel加载项注册表（CurrentUser）

echo reg add "HKCU\Software\Kingsoft\Office\6.0\ET\Addins\YYTools.ExcelAddin" /v FriendlyName /t REG_SZ /d "YY运单匹配工具" /f ^>nul

echo reg add "HKCU\Software\Kingsoft\Office\6.0\ET\Addins\YYTools.ExcelAddin" /v Description /t REG_SZ /d "运单匹配工具加载项" /f ^>nul

echo reg add "HKCU\Software\Kingsoft\Office\6.0\ET\Addins\YYTools.ExcelAddin" /v LoadBehavior /t REG_DWORD /d 3 /f ^>nul

echo reg add "HKCU\Software\Kingsoft\Office\6.0\ET\Addins\YYTools.ExcelAddin" /v CommandLineSafe /t REG_DWORD /d 1 /f ^>nul

echo reg add "HKCU\Software\Microsoft\Office\Excel\Addins\YYTools.ExcelAddin" /v FriendlyName /t REG_SZ /d "YY运单匹配工具" /f ^>nul

echo reg add "HKCU\Software\Microsoft\Office\Excel\Addins\YYTools.ExcelAddin" /v Description /t REG_SZ /d "运单匹配工具加载项" /f ^>nul

echo reg add "HKCU\Software\Microsoft\Office\Excel\Addins\YYTools.ExcelAddin" /v LoadBehavior /t REG_DWORD /d 3 /f ^>nul

echo reg add "HKCU\Software\Microsoft\Office\Excel\Addins\YYTools.ExcelAddin" /v CommandLineSafe /t REG_DWORD /d 1 /f ^>nul

echo echo 创建桌面快捷方式...

echo powershell -NoProfile -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\\YY运单匹配工具.lnk'); $s.TargetPath='%%INSTALL_DIR%%\\TestProgram.exe'; $s.WorkingDirectory='%%INSTALL_DIR%%'; $s.Description='YY运单匹配工具 v1.5'; $s.Save()" ^>nul

echo echo.

echo echo 安装完成！

echo echo 使用方法：

echo echo 1. 重启WPS表格（ET），在顶部“YY工具”菜单中点击“运单匹配工具”

echo echo 2. 或双击桌面“YY运单匹配工具”快捷方式

echo endlocal
) > "YYTools_Final\安装.bat"

rem 5. 生成卸载脚本
(
echo @echo off

echo chcp 65001 ^>nul

echo setlocal enableextensions

echo set INSTALL_DIR=%%USERPROFILE%%\Documents\YYTools

echo echo 注销COM组件...

echo set REGASM64=%%SystemRoot%%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

echo set REGASM32=%%SystemRoot%%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe

echo "%%REGASM64%%" /u "%%INSTALL_DIR%%\YYTools.dll" ^>nul

echo "%%REGASM32%%" /u "%%INSTALL_DIR%%\YYTools.dll" ^>nul

echo echo 删除注册表加载项...

echo reg delete "HKCU\Software\Kingsoft\Office\6.0\ET\Addins\YYTools.ExcelAddin" /f ^>nul

echo reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\YYTools.ExcelAddin" /f ^>nul

echo echo 删除文件...

echo del /q "%%INSTALL_DIR%%\YYTools.dll" ^>nul

echo del /q "%%INSTALL_DIR%%\TestProgram.exe" ^>nul

echo echo 卸载完成！

echo endlocal
) > "YYTools_Final\卸载.bat"

echo.
echo =====================================
echo 发布完成！目录：YYTools_Final\
echo =====================================
echo.

goto :eof

:build_fail
echo ❌ 编译失败，请检查错误信息。
pause 