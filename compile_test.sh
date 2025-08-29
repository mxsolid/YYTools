#!/bin/bash

# 设置UTF-8编码
export LANG=zh_CN.UTF-8
export LC_ALL=zh_CN.UTF-8

echo "========================================"
echo "YY工具 DPI优化版本编译测试 (Linux)"
echo "========================================"
echo

echo "正在检查环境..."

# 检查项目文件
if [ ! -f "YYTools/YYTools.csproj" ]; then
    echo "错误: 找不到项目文件 YYTools/YYTools.csproj"
    exit 1
fi

# 检查.NET环境
if command -v dotnet &> /dev/null; then
    echo "✓ 找到 .NET Core: $(dotnet --version)"
    DOTNET_CMD="dotnet"
elif command -v mono &> /dev/null; then
    echo "✓ 找到 Mono: $(mono --version | head -n1)"
    DOTNET_CMD="mono"
else
    echo "错误: 未找到 .NET 运行时环境"
    echo "请安装 .NET Core 或 Mono"
    exit 1
fi

echo "正在清理之前的编译结果..."
if [ -d "YYTools/bin" ]; then
    rm -rf "YYTools/bin"
fi
if [ -d "YYTools/obj" ]; then
    rm -rf "YYTools/obj"
fi

echo
echo "正在编译项目..."
cd YYTools

echo "使用 $DOTNET_CMD 编译..."

if [ "$DOTNET_CMD" = "dotnet" ]; then
    # 使用 .NET Core
    dotnet build YYTools.csproj -c Debug --verbosity minimal
    BUILD_RESULT=$?
else
    # 使用 Mono
    mono --version > /dev/null 2>&1
    if [ $? -eq 0 ]; then
        # 尝试使用 MSBuild
        if command -v msbuild &> /dev/null; then
            msbuild YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /verbosity:minimal
            BUILD_RESULT=$?
        else
            echo "警告: 未找到 MSBuild，尝试使用 xbuild..."
            xbuild YYTools.csproj /p:Configuration=Debug /p:Platform="Any CPU" /verbosity:minimal
            BUILD_RESULT=$?
        fi
    else
        echo "错误: Mono 运行时不可用"
        exit 1
    fi
fi

if [ $BUILD_RESULT -eq 0 ]; then
    echo
    echo "========================================"
    echo "编译成功！"
    echo "========================================"
    echo
    
    # 检查生成的文件
    if [ -f "bin/Debug/YYTools.exe" ]; then
        echo "✓ YYTools.exe 已生成"
        echo "✓ 文件大小: $(du -h bin/Debug/YYTools.exe | cut -f1)"
    else
        echo "✗ YYTools.exe 未找到"
    fi
    
    if [ -f "bin/Debug/YYTools.dll" ]; then
        echo "✓ YYTools.dll 已生成"
        echo "✓ 文件大小: $(du -h bin/Debug/YYTools.dll | cut -f1)"
    else
        echo "✗ YYTools.dll 未找到"
    fi
    
    echo
    echo "编译测试完成！"
    
else
    echo
    echo "========================================"
    echo "编译失败！"
    echo "========================================"
    echo
    echo "请检查错误信息并修复问题。"
    echo
    exit 1
fi

echo
echo "按任意键退出..."
read -n 1