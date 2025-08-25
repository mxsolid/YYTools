# YYTools 快速解决方案指南

## 🔧 问题：工具栏没有出现"YY工具"菜单

### ✅ 解决步骤

#### 1. **确认COM注册状态**
```cmd
# 运行检查脚本
.\check_registration.bat
```
如果显示"✓ COM对象创建成功"，说明注册正常。

#### 2. **重新注册COM组件（以管理员身份）**
```cmd
# 右键点击以下文件，选择"以管理员身份运行"
install_admin.bat
```

#### 3. **手动安装菜单**
在WPS/Excel的VBA编辑器中运行：
```vba
Sub 安装菜单()
    Dim result As String
    result = CreateObject("YYTools.ExcelAddin").InstallMenu()
    MsgBox result
End Sub
```

#### 4. **验证安装**
运行测试程序：
```cmd
.\bin\Debug\YYToolsTest.exe
```

---

## 🔧 问题：VBA代码运行提示找不到对象

### ✅ 解决步骤

#### 1. **检查COM注册**
```cmd
.\check_registration.bat
```

#### 2. **使用正确的VBA调用方式**
```vba
' 正确的调用方式
Sub 测试YYTools()
    On Error Resume Next
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If Err.Number <> 0 Then
        MsgBox "错误: " & Err.Description
        Exit Sub
    End If
    
    ' 调用方法
    Dim info As String
    info = obj.GetDetailedApplicationInfo()
    MsgBox info
    
    Set obj = Nothing
End Sub
```

#### 3. **导入完整的测试宏**
将 `WPS_Test_Macros.vba` 文件内容复制到VBA编辑器中，然后运行：
```vba
Call YYTools_综合测试
```

---

## 🚀 可用的VBA方法

### 基础方法
- `GetApplicationInfo()` - 获取基本应用程序信息
- `GetDetailedApplicationInfo()` - 获取详细应用程序信息
- `ShowMatchForm()` - 显示匹配窗体
- `ShowSettings()` - 显示设置窗体

### 菜单管理
- `InstallMenu()` - 手动安装菜单到工具栏
- `RefreshMenu()` - 刷新菜单
- `CreateWPSMenu()` - 创建WPS菜单（内部调用）

### 数据获取
- `GetExcelApplication()` - 获取应用程序实例
- `GetOpenWorkbooks()` - 获取打开的工作簿列表
- `GetWorksheetNames(workbook)` - 获取工作表名称

---

## 🧪 测试用VBA代码

### 1. 快速测试
```vba
Sub 快速测试()
    MsgBox CreateObject("YYTools.ExcelAddin").GetApplicationInfo()
End Sub
```

### 2. 安装菜单
```vba
Sub 安装菜单()
    MsgBox CreateObject("YYTools.ExcelAddin").InstallMenu()
End Sub
```

### 3. 详细信息
```vba
Sub 详细信息()
    MsgBox CreateObject("YYTools.ExcelAddin").GetDetailedApplicationInfo()
End Sub
```

### 4. 综合测试
```vba
Sub 综合测试()
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    ' 测试1: 获取信息
    MsgBox "基本信息:" & vbCrLf & obj.GetApplicationInfo()
    
    ' 测试2: 安装菜单
    MsgBox "菜单安装:" & vbCrLf & obj.InstallMenu()
    
    ' 测试3: 显示匹配工具
    obj.ShowMatchForm
    
    Set obj = Nothing
End Sub
```

---

## 🛠️ 故障排除

### 如果COM创建失败
1. **以管理员身份重新运行** `install_admin.bat`
2. **检查DLL文件**是否存在：`bin\Debug\YYTools.dll`
3. **查看Windows事件查看器**的应用程序日志
4. **重启WPS/Excel**后再试

### 如果菜单不显示
1. **手动调用** `InstallMenu()` 方法
2. **检查WPS/Excel权限**，确保允许COM加载项
3. **尝试刷新菜单**：调用 `RefreshMenu()` 方法

### 如果应用程序检测失败
1. **确保WPS/Excel已启动**
2. **打开至少一个工作簿文件**
3. **检查文件是否处于保护模式**

---

## 📋 文件列表

- `YYTools.dll` - 主程序文件（在bin\Debug目录）
- `YYToolsTest.exe` - 测试程序
- `check_registration.bat` - 注册状态检查
- `install_admin.bat` - 管理员安装脚本
- `WPS_Test_Macros.vba` - VBA测试宏

---

## ✅ 成功标志

当一切正常时，您应该看到：
1. ✓ 测试程序运行无错误
2. ✓ VBA调用返回正确信息
3. ✓ WPS/Excel工具栏出现"YY工具"菜单
4. ✓ 菜单按钮可以正常点击

---

**需要帮助？** 请运行 `.\bin\Debug\YYToolsTest.exe` 查看详细的测试结果。 