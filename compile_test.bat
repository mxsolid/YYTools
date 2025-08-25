@echo off
echo 编译测试程序...

"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" ^
  /target:winexe ^
  /reference:"YYTools\bin\Release\YYTools.dll" ^
  /reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Windows.Forms.dll" ^
  /reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.dll" ^
  "TestProgram.cs" ^
  "/out:TestProgram.exe"

if errorlevel 1 (
    echo 编译失败！
    pause
    exit /b 1
) else (
    echo 编译成功！
    dir TestProgram.exe
)

pause 