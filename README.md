# YYTools - WPS/Excel COM 加载项

## 项目概述

YYTools 是一个专为 WPS 表格设计的 COM 加载项，同时兼容 Microsoft Excel。该工具提供运单匹配功能和系统设置，能够自动集成到 WPS 菜单工具栏中。

## 核心特性

### 🚀 WPS 优先策略
- **智能应用程序检测**：优先连接 WPS 表格，其次 Excel
- **多种 ProgID 支持**：`Ket.Application` > `WPS.Application` > `Kingsoft.Application` > `ET.Application` > `Excel.Application`
- **ROT 兜底机制**：当常规方式失败时，通过 Running Object Table 检测

### 📊 可靠的工作簿检测
- **三重检测逻辑**：
  1. 直接获取 `Workbooks.Count`
  2. 尝试访问 `ActiveWorkbook`
  3. 遍历工作簿集合验证
- **容错机制**：即使某种检测方式失败，其他方式仍能正常工作

### 🎯 默认选中激活文件
- **智能选择**：自动识别并标记当前激活的工作簿
- **备用逻辑**：如无激活工作簿，自动选择第一个打开的文件
- **多种获取方式**：索引访问和 foreach 遍历双重保障

### 🔧 增强的菜单集成
- **自动菜单创建**：启动时自动在 WPS 工具栏添加"YY工具"菜单
- **双重添加策略**：优先菜单栏，失败则添加为工具栏
- **工具提示支持**：每个按钮都有详细的工具提示
- **位置优化**：尝试将菜单放置在顶部显著位置

### 🛡️ 向下兼容设计
- **C# 5.0 兼容**：避免使用 C# 6.0+ 语法特性
- **传统字符串操作**：使用 `string.Format()` 替代字符串插值
- **稳定的语法结构**：确保在旧版本 .NET Framework 上正常编译
- **无Office依赖**：运行时通过动态COM调用，无需安装Office开发工具

## 主要修复内容

### 1. 工作簿检测逻辑重写 ✅
**问题**：原 `GetWorkbookCountInternal` 方法检测不准确
**解决**：新的 `HasOpenWorkbooks` 方法，三种检测方式并行

### 2. 激活文件默认选中 ✅
**问题**：多文件时无法识别当前激活文件
**解决**：`GetOpenWorkbooks` 方法正确标记激活工作簿

### 3. WPS 菜单工具栏集成 ✅
**问题**：仅做安装注册，未真正集成到界面
**解决**：增强的 `CreateWPSMenu` 方法，支持菜单栏和工具栏

### 4. 语法兼容性修复 ✅
**问题**：使用了 C# 6.0 语法导致编译错误
**解决**：改用 C# 5.0 兼容语法，通过编译测试

### 5. WPS 优先级调整 ✅
**问题**：WPS 和 Excel 检测优先级不明确
**解决**：明确的优先级序列，WPS 各版本优先于 Excel

### 6. Office Interop 依赖移除 ✅
**问题**：依赖 Microsoft.Office.Interop.Excel.dll 导致编译失败
**解决**：改用动态COM调用，运行时绑定，无需Office开发工具

## 文件结构

```
YYTools/
├── YYTools/
│   ├── ExcelAddin.cs           # 主要的 COM 加载项类
│   └── ColumnSelectionForm.cs  # 列选择对话框（含ExcelHelper）
├── build_and_test.bat          # 普通编译脚本
├── install_admin.bat           # 管理员安装脚本（推荐）
├── WPS_Test_Macros.vba         # WPS 测试宏文件
└── README.md                   # 项目文档
```

## 安装和使用

### 环境要求
- Windows 操作系统
- .NET Framework 4.0 或更高版本
- WPS 表格 或 Microsoft Excel

### 编译安装

#### 方式一：管理员安装（推荐）
1. **右键点击 `install_admin.bat`**
2. **选择"以管理员身份运行"**
3. 等待编译和注册完成

#### 方式二：普通编译
1. **以管理员身份运行** 命令提示符
2. 导航到项目目录
3. 执行编译脚本：
   ```cmd
   build_and_test.bat
   ```

### 验证安装
1. 打开 WPS 表格
2. 查看工具栏是否出现"YY工具"菜单
3. 或运行测试宏：
   ```vba
   ' 在 VBA 编辑器中运行
   Sub 测试YYTools()
       Dim result As String
       result = CreateObject("YYTools.ExcelAddin").GetApplicationInfo()
       MsgBox result
   End Sub
   ```

## API 方法

### 主要功能方法
- `ShowMatchForm()` - 显示运单匹配工具
- `ShowSettings()` - 显示工具设置
- `GetApplicationInfo()` - 获取应用程序信息
- `RefreshMenu()` - 刷新菜单

### 应用程序检测
- `GetExcelApplication()` - 获取 WPS/Excel 应用程序实例
- `GetOpenWorkbooks()` - 获取打开的工作簿列表
- `GetWorksheetNames()` - 获取工作表名称列表

### 菜单管理
- `CreateWPSMenu()` - 创建 WPS 菜单
- `RegisterFunction()` - COM 注册时调用
- `UnregisterFunction()` - COM 反注册时调用

## VBA 调用示例

```vba
' 运单匹配工具
Sub YYTools_运单匹配()
    Dim objYYTools As Object
    Set objYYTools = CreateObject("YYTools.ExcelAddin")
    objYYTools.ShowMatchForm
    Set objYYTools = Nothing
End Sub

' 工具设置
Sub YYTools_设置()
    Dim objYYTools As Object
    Set objYYTools = CreateObject("YYTools.ExcelAddin")
    objYYTools.ShowSettings
    Set objYYTools = Nothing
End Sub

' 测试连接
Sub YYTools_测试连接()
    Dim objYYTools As Object
    Set objYYTools = CreateObject("YYTools.ExcelAddin")
    Dim info As String
    info = objYYTools.GetApplicationInfo()
    MsgBox "应用程序信息: " & vbCrLf & info
    Set objYYTools = Nothing
End Sub
```

## 技术特点

### 应用程序检测策略
1. **WPS 优先检测**：按优先级尝试各种 WPS ProgID
2. **Excel 备用检测**：WPS 失败后尝试 Excel
3. **延迟重试机制**：短暂等待后重新尝试 WPS
4. **ROT 兜底策略**：最后通过 Running Object Table 搜索

### 工作簿检测机制
1. **计数检测**：直接获取 `Workbooks.Count`
2. **激活检测**：尝试访问 `ActiveWorkbook`
3. **遍历检测**：尝试获取第一个工作簿
4. **容错处理**：任一方式成功即返回 true

### 菜单集成策略
1. **菜单栏优先**：首先尝试添加到菜单栏
2. **工具栏备用**：菜单栏失败则添加为工具栏
3. **位置优化**：设置菜单显示在顶部
4. **清理机制**：注册前先删除已存在的菜单

### COM调用设计
1. **动态绑定**：使用反射和 `InvokeMember` 调用COM接口
2. **类型安全**：运行时类型检查和转换
3. **异常处理**：完整的try-catch保护
4. **资源管理**：及时释放COM对象引用

## 兼容性说明

### 支持的应用程序
- **WPS 表格**（推荐）
  - Kingsoft Office
  - WPS Office 2016/2019/365
  - ET 表格
- **Microsoft Excel**
  - Excel 2010/2013/2016/2019/365

### 支持的 .NET 版本
- .NET Framework 4.0+
- .NET Framework 4.5+
- .NET Framework 4.6+
- .NET Framework 4.7+
- .NET Framework 4.8

## 故障排除

### 常见问题

**Q: 编译时提示找不到 Microsoft.Office.Interop.Excel.dll**
A: ✅ **已解决**！新版本不再依赖Office Interop，使用动态COM调用。

**Q: WPS 中看不到 YY工具 菜单**
A: 
1. 确保以管理员权限运行了安装脚本
2. 重启 WPS 表格
3. 手动运行 `YYTools_刷新菜单` 宏

**Q: 提示"无法连接到WPS表格或Excel"**
A:
1. 确保 WPS 或 Excel 已启动
2. 确保至少打开一个工作簿文件
3. 检查文件是否处于保护模式

**Q: COM 注册失败**
A:
1. 使用 `install_admin.bat` 以管理员身份安装
2. 确保 .NET Framework 已正确安装
3. 检查防病毒软件是否阻止了注册

**Q: 运行时出现"未能找到程序集"错误**
A:
1. 确认DLL文件在正确位置（bin\Debug\YYTools.dll）
2. 重新运行管理员安装脚本
3. 检查注册表中的COM组件信息

## 开发团队

本项目专为 WPS 用户优化，确保与 WPS 表格的最佳兼容性，同时保持对 Excel 的支持。

## 版本历史

### v2.1 (当前版本)
- ✅ 移除Office Interop依赖，使用动态COM调用
- ✅ 重写工作簿检测逻辑，解决检测不准确问题
- ✅ 实现默认选中激活文件功能
- ✅ 增强 WPS 菜单工具栏集成
- ✅ 修复 C# 6.0 语法兼容性问题
- ✅ 优化 WPS 优先检测策略
- ✅ 添加管理员安装脚本
- ✅ 编译测试通过，生成17KB DLL文件

### v2.0 (上一版本)
- ✅ 重写工作簿检测逻辑，解决检测不准确问题
- ✅ 实现默认选中激活文件功能
- ✅ 增强 WPS 菜单工具栏集成
- ✅ 修复 C# 6.0 语法兼容性问题
- ✅ 优化 WPS 优先检测策略
- ✅ 添加完整的测试和文档

### v1.0 (原始版本)
- 基础的 COM 加载项功能
- 简单的应用程序检测
- 基础菜单创建

---

**注意**：请确保在安装和使用过程中具有足够的系统权限，建议使用 `install_admin.bat` 脚本进行安装。 