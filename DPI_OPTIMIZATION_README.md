# YY工具 DPI优化版本说明

## 概述

本版本针对Windows高DPI显示器进行了全面优化，解决了以下问题：
- 2K分辨率下字体不清晰
- 3200×2000屏幕下界面正常但2K屏幕有问题
- 启用Per-Monitor V2 DPI感知后字体过大、显示不全
- 多显示器间移动时界面比例不协调
- 布局混乱和组件不协调

## 主要改进

### 1. 全新的DPI管理器 (DPIManager.cs)

- **多显示器支持**: 自动检测所有显示器的DPI设置
- **Per-Monitor V2感知**: 支持Windows 10的每显示器DPI感知
- **动态DPI响应**: 实时响应显示器DPI变化
- **智能缩放算法**: 限制最大缩放比例，避免界面过大

### 2. 优化的UI增强器 (UIEnhancer.cs)

- **控件级DPI优化**: 针对不同控件类型进行专门优化
- **字体大小控制**: 限制字体缩放范围(0.8x - 1.5x)
- **布局自适应**: 自动调整控件大小和位置
- **防止溢出**: 限制控件最大尺寸，避免超出屏幕边界

### 3. 改进的程序启动 (Program.cs)

- **智能DPI感知设置**: 自动选择最佳的DPI感知API
- **错误处理**: 完善的错误处理和回退机制
- **版本检测**: 根据Windows版本选择合适的DPI设置

### 4. 配置文件优化 (app.config)

- **DPI感知配置**: 启用Per-Monitor V2支持
- **高DPI自动调整**: 启用Windows Forms高DPI自动调整
- **运行时配置**: 添加必要的运行时配置

## 技术特性

### DPI感知级别
1. **Per-Monitor V2** (Windows 10 1703+): 最高级别，每显示器独立DPI
2. **Per-Monitor** (Windows 8.1+): 每显示器DPI感知
3. **System DPI** (Windows Vista+): 系统级DPI感知

### 智能缩放算法
- **字体缩放**: 限制在0.8x - 1.5x范围内
- **控件缩放**: 限制最大宽度800px，最大高度600px
- **窗体缩放**: 确保不超出屏幕边界
- **位置调整**: 防止控件位置为负数

### 多显示器支持
- **自动检测**: 启动时枚举所有显示器
- **DPI缓存**: 缓存每个显示器的DPI信息
- **动态更新**: 支持运行时显示器热插拔

## 使用方法

### 1. 编译项目

#### Windows环境
```batch
# 使用编译测试脚本
compile_test.bat

# 或手动编译
cd YYTools
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" YYTools.csproj /p:Configuration=Debug
```

#### Linux环境
```bash
# 使用编译测试脚本
chmod +x compile_test.sh
./compile_test.sh

# 或手动编译
cd YYTools
dotnet build YYTools.csproj -c Debug
```

### 2. 运行程序

编译成功后，运行 `YYTools\bin\Debug\YYTools.exe`

### 3. 验证DPI优化

1. **检查日志**: 程序启动时会输出DPI相关信息
2. **界面显示**: 字体应该清晰且大小适中
3. **多显示器**: 在不同DPI显示器间移动时界面应保持协调

## 配置选项

### app.config 配置项

```xml
<!-- DPI感知配置 -->
<add key="DpiAwareness" value="PerMonitorV2"/>
<add key="EnableWindowsFormsHighDpiAutoResizing" value="true"/>

<!-- 应用程序设置 -->
<add key="EnableHighDpiSupport" value="true"/>
<add key="MaxFontScale" value="1.5"/>
<add key="MinFontScale" value="0.8"/>
```

### 代码中的常量

```csharp
// UIEnhancer.cs 中的常量
private const float MAX_FONT_SCALE = 1.5f;      // 最大字体缩放
private const float MIN_FONT_SCALE = 0.8f;      // 最小字体缩放
private const int MAX_CONTROL_WIDTH = 800;      // 最大控件宽度
private const int MAX_CONTROL_HEIGHT = 600;     // 最大控件高度
```

## 故障排除

### 常见问题

1. **字体仍然过大**
   - 检查是否启用了Per-Monitor V2
   - 查看日志中的DPI信息
   - 调整MAX_FONT_SCALE常量

2. **界面布局混乱**
   - 确保在InitializeComponent之后调用DPI优化
   - 检查控件的大小限制设置
   - 查看是否有控件位置为负数

3. **多显示器问题**
   - 刷新显示器DPI信息: `DPIManager.RefreshMonitorDpiInfo()`
   - 检查每个显示器的DPI设置
   - 确保Windows显示设置正确

### 调试信息

程序会输出详细的DPI信息到日志：
```
DPI管理器初始化完成 - 系统DPI: 1.25, Per-Monitor V2: True
窗体DPI感知已启用: MatchForm, DPI缩放: 1.25
窗体DPI设置完成: 系统DPI: 1.25, 主显示器DPI: 1.25, 高DPI: True, 超高DPI: False, Per-Monitor V2: True
```

## 性能优化

### 内存管理
- **DPI缓存**: 避免重复计算DPI信息
- **控件池**: 重用DPI调整后的控件
- **延迟加载**: DPI优化在窗体显示后执行

### 响应性
- **异步处理**: DPI变化事件异步处理
- **防抖机制**: 避免频繁的DPI调整
- **增量更新**: 只更新变化的控件

## 兼容性

### 操作系统支持
- **Windows 10 1703+**: 完整支持Per-Monitor V2
- **Windows 8.1**: 支持Per-Monitor DPI感知
- **Windows 7/8**: 支持System DPI感知
- **Windows Vista**: 基础DPI感知支持

### .NET Framework版本
- **.NET Framework 4.8**: 主要目标版本
- **.NET Framework 4.7.2+**: 支持高DPI功能
- **.NET Framework 4.6.1+**: 基础DPI支持

## 更新日志

### v3.3.0 (当前版本)
- ✨ 全新的DPI管理器，支持多显示器
- ✨ 智能UI增强器，自动优化所有控件
- ✨ Per-Monitor V2 DPI感知支持
- ✨ 动态DPI变化响应
- ✨ 字体大小智能控制
- ✨ 界面布局自适应
- 🐛 修复字体过大问题
- 🐛 修复布局混乱问题
- 🐛 修复多显示器比例不协调问题

### 已知问题
- 无

## 技术支持

如果遇到问题，请：
1. 查看程序日志输出
2. 检查Windows显示设置
3. 确认显示器DPI设置
4. 提供详细的错误信息

## 许可证

本项目遵循原有许可证条款。