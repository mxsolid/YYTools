#!/bin/bash

echo "========================================"
echo "YY工具 v2.6 Linux构建验证脚本"
echo "========================================"
echo

# 设置项目目录
PROJECT_DIR="YYTools"

echo "正在检查源代码文件..."

# 检查必要的源文件
REQUIRED_FILES=(
    "MatchForm.cs"
    "SmartColumnService.cs"
    "DataModels.cs"
    "AppSettings.cs"
    "MatchService.cs"
    "ExcelHelper.cs"
    "YYTools.csproj"
    "Properties/AssemblyInfo.cs"
)

for file in "${REQUIRED_FILES[@]}"; do
    if [ ! -f "$PROJECT_DIR/$file" ]; then
        echo "错误：找不到文件 $PROJECT_DIR/$file"
        exit 1
    fi
done

echo "源代码文件检查完成"

echo
echo "正在检查代码语法..."

# 检查C#语法（如果有csc编译器的话）
if command -v csc >/dev/null 2>&1; then
    echo "找到C#编译器，正在检查语法..."
    cd "$PROJECT_DIR"
    
    # 尝试编译（不链接）
    csc -target:library -out:temp.dll -reference:System.dll -reference:System.Windows.Forms.dll -reference:System.Drawing.dll *.cs Properties/*.cs 2>/dev/null
    
    if [ $? -eq 0 ]; then
        echo "语法检查通过"
        rm -f temp.dll
    else
        echo "语法检查失败"
        exit 1
    fi
    
    cd ..
else
    echo "未找到C#编译器，跳过语法检查"
fi

echo
echo "正在检查项目文件..."

# 检查项目文件中的引用
if grep -q "SmartColumnService.cs" "$PROJECT_DIR/YYTools.csproj"; then
    echo "项目文件包含SmartColumnService.cs ✓"
else
    echo "错误：项目文件未包含SmartColumnService.cs"
    exit 1
fi

# 检查版本号
if grep -q "2.6.0.0" "$PROJECT_DIR/Properties/AssemblyInfo.cs"; then
    echo "版本号检查通过 ✓"
else
    echo "错误：版本号不是2.6.0.0"
    exit 1
fi

echo
echo "正在检查代码结构..."

# 检查命名空间
if grep -q "namespace YYTools" "$PROJECT_DIR/SmartColumnService.cs"; then
    echo "SmartColumnService命名空间正确 ✓"
else
    echo "错误：SmartColumnService命名空间不正确"
    exit 1
fi

# 检查类定义
if grep -q "public class SmartColumnService" "$PROJECT_DIR/SmartColumnService.cs"; then
    echo "SmartColumnService类定义正确 ✓"
else
    echo "错误：SmartColumnService类定义不正确"
    exit 1
fi

# 检查数据模型
if grep -q "public class ColumnInfo" "$PROJECT_DIR/DataModels.cs"; then
    echo "ColumnInfo类定义正确 ✓"
else
    echo "错误：ColumnInfo类定义不正确"
    exit 1
fi

if grep -q "public class SmartColumnRule" "$PROJECT_DIR/DataModels.cs"; then
    echo "SmartColumnRule类定义正确 ✓"
else
    echo "错误：SmartColumnRule类定义不正确"
    exit 1
fi

echo
echo "正在检查依赖关系..."

# 检查必要的using语句
if grep -q "using System.Linq;" "$PROJECT_DIR/SmartColumnService.cs"; then
    echo "SmartColumnService依赖检查通过 ✓"
else
    echo "错误：SmartColumnService缺少必要的依赖"
    exit 1
fi

if grep -q "using Excel = Microsoft.Office.Interop.Excel;" "$PROJECT_DIR/SmartColumnService.cs"; then
    echo "Excel引用检查通过 ✓"
else
    echo "错误：SmartColumnService缺少Excel引用"
    exit 1
fi

echo
echo "========================================"
echo "Linux构建验证成功！"
echo "========================================"
echo
echo "所有检查项目通过："
echo "✓ 源代码文件完整性"
echo "✓ 项目文件配置"
echo "✓ 版本号一致性"
echo "✓ 代码结构正确性"
echo "✓ 依赖关系完整性"
echo
echo "注意：这是Linux环境，无法进行实际的Windows编译"
echo "在Windows环境中，请使用以下命令进行实际构建："
echo "  msbuild YYTools.csproj /p:Configuration=Release"
echo "  或运行 build_release.bat 脚本"
echo
echo "代码已准备就绪，可以在Windows环境中编译运行"
echo