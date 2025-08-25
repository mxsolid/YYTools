# 📦 YY运单匹配工具 - 安装指南

## 🎯 系统要求

### 必需组件
- ✅ **操作系统**：Windows 10/11 (64位推荐)
- ✅ **Office软件**：Microsoft Excel 2016+ 或 WPS表格
- ✅ **.NET Framework**：4.8 或更高版本
- ✅ **VSTO运行时**：Microsoft Visual Studio Tools for Office Runtime

### 开发环境（仅开发者需要）
- ✅ **Visual Studio**：2019/2022（包含Office开发工具）
- ✅ **MSBuild**：随Visual Studio安装
- ✅ **Office PIA**：Primary Interop Assemblies

## 🚀 用户安装（推荐）

### 方法一：使用安装包（最简单）

1. **下载安装包**
   ```
   从发布页面下载：YYTools_Setup.zip
   解压到任意文件夹
   ```

2. **运行安装程序**
   ```
   双击 setup.exe
   按照向导完成安装
   ```

3. **验证安装**
   ```
   打开Excel → 查看功能区是否有"YY工具"选项卡
   ```

### 方法二：便携版安装

1. **下载便携版**
   ```
   下载：YYTools_Portable.zip
   解压到任意文件夹（如：C:\YYTools）
   ```

2. **手动注册**
   ```powershell
   # 以管理员身份运行PowerShell
   cd "C:\YYTools"
   regsvr32 YYTools.dll
   ```

3. **Excel中启用**
   ```
   Excel → 文件 → 选项 → 加载项 → 管理(COM加载项) → 添加YYTools
   ```

## 🛠️ 开发者安装

### 环境配置

1. **安装Visual Studio 2022**
   ```
   下载地址：https://visualstudio.microsoft.com/downloads/
   工作负载：选择"Office/SharePoint开发"
   组件：包含VSTO开发工具
   ```

2. **验证环境**
   ```bash
   # 检查.NET Framework
   reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release
   
   # 检查MSBuild
   "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" -version
   ```

### 编译安装

1. **克隆项目**
   ```bash
   git clone <项目地址>
   cd YYTools
   ```

2. **编译项目**
   ```bash
   # 方法1：使用批处理（推荐）
   build.bat
   
   # 方法2：手动编译
   msbuild YYTools.sln /p:Configuration=Release
   ```

3. **生成安装包**
   ```bash
   # 生成完整安装包
   publish.bat
   
   # 输出位置：YYTools\bin\Release\publish\
   ```

## 🔧 故障排除

### 常见问题及解决方案

#### 1. 编译错误

**错误**：`找不到Microsoft.Office.Interop.Excel`
```bash
# 解决方案：
# 1. 确保安装了Office
# 2. 安装Office Developer Tools
# 3. 修改项目引用路径
```

**错误**：`找不到MSBuild.exe`
```bash
# 解决方案：
# 1. 安装Visual Studio Build Tools
# 2. 手动设置MSBuild路径
set MSBUILD_PATH="你的MSBuild路径"
```

#### 2. 运行时错误

**错误**：`插件未加载`
```bash
# 解决方案：
# 1. 检查Excel信任中心设置
Excel → 文件 → 选项 → 信任中心 → 信任中心设置 → 加载项
勾选："启用应用程序加载项"

# 2. 检查VSTO运行时
下载并安装：Microsoft Visual Studio Tools for Office Runtime
```

**错误**：`权限不足`
```bash
# 解决方案：
# 1. 以管理员身份运行Excel
# 2. 或者降低安全设置（不推荐）
```

#### 3. 功能问题

**问题**：`工具栏不显示`
```bash
# 解决方案：
# 1. 重启Excel
# 2. 重新安装插件
# 3. 检查注册表项是否正确
```

**问题**：`匹配速度慢`
```bash
# 解决方案：
# 1. 关闭Excel自动计算
# 2. 增加系统内存
# 3. 分批处理大文件
```

## 📋 卸载说明

### 完全卸载

1. **使用控制面板**
   ```
   控制面板 → 程序和功能 → 找到"YY运单匹配工具" → 卸载
   ```

2. **手动清理（可选）**
   ```bash
   # 删除用户设置
   删除文件夹：%APPDATA%\YYTools
   
   # 清理注册表（谨慎操作）
   删除注册表项：HKEY_CURRENT_USER\Software\YYTools
   ```

3. **重启Excel**
   ```
   确保插件完全卸载
   ```

## 🔐 安全说明

### 数字签名
- 📝 插件使用自签名证书
- ⚠️ 首次安装可能提示"未知发布者"
- ✅ 这是正常现象，可以安全安装

### 权限要求
- 📊 读取Excel工作簿和工作表
- ✏️ 写入指定单元格数据
- 🚫 不会访问网络或其他文件

### 隐私保护
- 🛡️ 所有数据处理在本地完成
- 🔒 不会上传任何数据到服务器
- 📈 不会收集用户使用情况

## 📞 技术支持

### 获取帮助
- 📖 [用户手册](README.md)
- 🚀 [快速入门](QUICKSTART.md)
- 🐛 [问题反馈](https://github.com/your-repo/issues)

### 联系方式
- 📧 邮件：support@yytools.com
- 💬 QQ群：123456789
- 🌐 官网：https://yytools.com

---

**安装遇到问题？随时联系我们！** 💪 