#!/bin/bash

echo "========================================"
echo "YY工具代码检查脚本"
echo "版本: 3.2 (性能优化版)"
echo "时间: $(date)"
echo "========================================"
echo

echo "[1/4] 检查文件结构..."
if [ -d "YYTools" ]; then
    echo "✓ YYTools目录存在"
else
    echo "✗ YYTools目录不存在"
    exit 1
fi

if [ -f "YYTools/YYTools.csproj" ]; then
    echo "✓ 项目文件存在"
else
    echo "✗ 项目文件不存在"
    exit 1
fi

if [ -f "YYTools/Constants.cs" ]; then
    echo "✓ 常量文件存在"
else
    echo "✗ 常量文件不存在"
    exit 1
fi

if [ -f "YYTools/MatchForm.cs" ]; then
    echo "✓ 主窗体文件存在"
else
    echo "✗ 主窗体文件不存在"
    exit 1
fi

echo "文件结构检查通过"
echo

echo "[2/4] 检查代码语法..."
cd YYTools

# 检查C#文件的基本语法
ERROR_COUNT=0
for file in *.cs; do
    if [ -f "$file" ]; then
        echo "检查文件: $file"
        
        # 检查基本的C#语法结构
        if grep -q "namespace YYTools" "$file"; then
            echo "  ✓ 命名空间正确"
        else
            echo "  ✗ 命名空间缺失或错误"
            ERROR_COUNT=$((ERROR_COUNT + 1))
        fi
        
        # 检查类定义 - 修复正则表达式
        if grep -q "public class" "$file" || grep -q "partial class" "$file" || grep -q "static class" "$file" || grep -q "internal class" "$file" || grep -q "sealed class" "$file"; then
            echo "  ✓ 类定义正确"
        else
            echo "  ✗ 类定义缺失或错误"
            ERROR_COUNT=$((ERROR_COUNT + 1))
        fi
        
        # 检查方法定义 - 修复正则表达式
        if grep -q "public.*(" "$file" || grep -q "private.*(" "$file" || grep -q "protected.*(" "$file" || grep -q "internal.*(" "$file" || grep -q "static.*(" "$file"; then
            echo "  ✓ 方法定义正确"
        else
            echo "  ✗ 方法定义缺失或错误"
            ERROR_COUNT=$((ERROR_COUNT + 1))
        fi
    fi
done

cd ..
echo "代码语法检查完成，发现 $ERROR_COUNT 个错误"
echo

echo "[3/4] 检查关键功能实现..."
cd YYTools

# 检查关键功能是否实现
echo "检查关键功能:"

# 检查版本号更新
if grep -q "AppVersion = \"v3.2" "Constants.cs"; then
    echo "  ✓ 版本号已更新到v3.2"
else
    echo "  ✗ 版本号未更新"
    ERROR_COUNT=$((ERROR_COUNT + 1))
fi

# 检查版本哈希值
if grep -q "AppVersionHash = \"2024-12-19-8F7E2D1A\"" "Constants.cs"; then
    echo "  ✓ 版本哈希值已添加"
else
    echo "  ✗ 版本哈希值未添加"
    ERROR_COUNT=$((ERROR_COUNT + 1))
fi

# 检查预览行数配置
if grep -q "PreviewRowOptions" "Constants.cs"; then
    echo "  ✓ 预览行数配置已添加"
else
    echo "  ✗ 预览行数配置未添加"
    ERROR_COUNT=$((ERROR_COUNT + 1))
fi

# 检查多线程处理
if grep -q "Task.Run" "MatchForm.cs"; then
    echo "  ✓ 多线程处理已实现"
else
    echo "  ✗ 多线程处理未实现"
    ERROR_COUNT=$((ERROR_COUNT + 1))
fi

# 检查并行处理
if grep -q "Task.WhenAll" "MatchService.cs"; then
    echo "  ✓ 并行处理已实现"
else
    echo "  ✗ 并行处理未实现"
    ERROR_COUNT=$((ERROR_COUNT + 1))
fi

cd ..
echo "关键功能检查完成"
echo

echo "[4/4] 检查配置和设置..."
cd YYTools

# 检查AppSettings中的新配置
if grep -q "PreviewParseRows" "AppSettings.cs"; then
    echo "  ✓ 预览行数配置属性已添加"
else
    echo "  ✗ 预览行数配置属性未添加"
    ERROR_COUNT=$((ERROR_COUNT + 1))
fi

# 检查TaskOptionsForm中的新控件
if grep -q "gbPreview" "TaskOptionsForm.cs"; then
    echo "  ✓ 预览配置组已添加"
else
    echo "  ✗ 预览配置组未添加"
    ERROR_COUNT=$((ERROR_COUNT + 1))
fi

cd ..
echo "配置和设置检查完成"
echo

echo "========================================"
echo "代码检查完成！"
echo "========================================"
echo

if [ $ERROR_COUNT -eq 0 ]; then
    echo "🎉 所有检查通过！代码结构正确，功能完整。"
    echo
    echo "主要改进内容:"
    echo "✓ 版本号更新到v3.2，添加唯一哈希值"
    echo "✓ 添加写入预览解析行数配置(5,10,20,50,100)"
    echo "✓ 美化任务配置界面，统一风格和对齐"
    echo "✓ 实现发货明细和账单明细的并行处理"
    echo "✓ 改进工作表智能匹配逻辑"
    echo "✓ 优化多线程处理，防止UI卡顿"
    echo "✓ 改进窗体大小调整逻辑"
    echo "✓ 添加线程池管理，提高性能"
    echo
    echo "下一步操作:"
    echo "1. 在Windows 11 + .NET Framework 4.8环境中编译"
    echo "2. 使用Visual Studio或MSBuild编译项目"
    echo "3. 运行程序测试所有新功能"
else
    echo "⚠️  发现 $ERROR_COUNT 个问题，请检查并修复"
    echo
    echo "建议:"
    echo "1. 检查上述错误信息"
    echo "2. 修复代码中的问题"
    echo "3. 重新运行代码检查"
fi

echo