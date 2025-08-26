@echo off
chcp 65001 >nul
echo =====================================
echo YY运单匹配工具 - 简化编译脚本
echo =====================================
echo.

REM 清理之前的编译结果
if exist "YYTools\bin" rmdir /s /q "YYTools\bin"
if exist "YYTools\obj" rmdir /s /q "YYTools\obj"

REM 创建输出目录
mkdir "YYTools\bin\Release" 2>nul

echo 正在编译YYTools.dll...

REM 使用CSC直接编译DLL（不包含Microsoft.Office.Interop.Excel引用来避免依赖问题）
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /target:library /out:YYTools\bin\Release\YYTools.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Data.dll YYTools\ExcelAddin.cs YYTools\MatchService.cs YYTools\MatchForm.cs YYTools\MatchForm.Designer.cs YYTools\ExcelHelper.cs YYTools\ColumnSelectionForm.cs YYTools\Properties\AssemblyInfo.cs

if %errorlevel% neq 0 (
    echo 编译DLL失败！
    pause
    exit /b 1
)

echo ✓ YYTools.dll 编译成功

echo.
echo 正在编译测试程序...

REM 编译测试程序
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /target:winexe /r:YYTools\bin\Release\YYTools.dll /r:System.Windows.Forms.dll TestProgram.cs

if %errorlevel% neq 0 (
    echo 编译测试程序失败！
    pause
    exit /b 1
)

echo ✓ TestProgram.exe 编译成功

echo.
echo =====================================
echo 编译完成！
echo =====================================

echo 输出文件：
dir YYTools\bin\Release\YYTools.dll | find "YYTools.dll"
dir TestProgram.exe | find "TestProgram.exe"

echo.
echo 现在可以运行 TestProgram.exe 进行测试！

pause