#!/bin/bash

echo "========================================"
echo "YY工具编译测试脚本 (Linux版本)"
echo "版本: 3.2 (性能优化版)"
echo "时间: $(date)"
echo "========================================"
echo

echo "[1/4] 清理旧的编译文件..."
if [ -d "YYTools/bin" ]; then
    rm -rf "YYTools/bin"
fi
if [ -d "YYTools/obj" ]; then
    rm -rf "YYTools/obj"
fi
echo "清理完成"
echo

echo "[2/4] 检查 .NET SDK 版本..."
if command -v dotnet &> /dev/null; then
    dotnet --version
    echo ".NET检查通过"
else
    echo "警告: 未检测到dotnet命令，可能需要在Windows环境中编译"
    echo "继续尝试使用mono编译..."
fi
echo

echo "[3/4] 编译项目 (.NET 8)..."
cd YYTools

# 尝试使用dotnet编译
if command -v dotnet &> /dev/null; then
    export DOTNET_CLI_UI_LANGUAGE=zh-Hans
    export DOTNET_CLI_TELEMETRY_OPTOUT=1
    dotnet restore --nologo --verbosity minimal
    dotnet build YYTools.csproj -c Release -p:ContinuousIntegrationBuild=true --nologo --verbosity minimal
    BUILD_SUCCESS=$?
else
    # 尝试使用mono编译
    if command -v mcs &> /dev/null; then
        echo "使用Mono编译器..."
        mcs -target:winexe -out:bin/Release/YYTools.exe -r:System.Windows.Forms.dll -r:System.Drawing.dll -r:System.Data.dll -r:System.Core.dll -r:System.Xml.dll -r:Microsoft.Office.Interop.Excel.dll *.cs
        BUILD_SUCCESS=$?
    else
        echo "错误: 未找到可用的编译器"
        echo "请在Windows环境中使用MSBuild编译，或安装Mono"
        cd ..
        exit 1
    fi
fi

if [ $BUILD_SUCCESS -ne 0 ]; then
    echo
    echo "错误: 编译失败！"
    echo "请检查代码中的错误"
    cd ..
    exit 1
fi

cd ..
echo "编译成功！"
echo

echo "[4/4] 检查编译结果..."
if [ -f "YYTools/bin/Release/YYTools.exe" ]; then
    echo "编译测试通过！"
    echo "可执行文件位置: YYTools/bin/Release/YYTools.exe"
    echo
    echo "注意: 这是一个Windows应用程序，需要在Windows环境中运行"
    echo "建议在Windows 11 + .NET Framework 4.8环境中运行"
else
    echo "错误: 未找到编译后的可执行文件"
    exit 1
fi

echo
echo "========================================"
echo "编译测试完成！"
echo "========================================"
echo
echo "下一步操作建议:"
echo "1. 将项目复制到Windows环境"
echo "2. 使用Visual Studio或MSBuild编译"
echo "3. 在Windows环境中运行测试"
echo