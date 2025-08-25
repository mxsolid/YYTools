Attribute VB_Name = "YYTools_Global_Macros"

' ========================================
' YYTools 全局宏模块
' 用于响应菜单按钮的OnAction事件
' ========================================

' 显示运单匹配工具
Sub YYToolsShowMatchForm()
    On Error GoTo ErrorHandler
    
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If obj Is Nothing Then
        MsgBox "无法创建YYTools对象，请确保COM组件已正确注册", vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 直接调用ShowMatchForm方法
    obj.ShowMatchForm
    
    Set obj = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "运行运单匹配工具时出错: " & Err.Number & " - " & Err.Description, vbCritical, "YYTools错误"
    If Not obj Is Nothing Then Set obj = Nothing
End Sub

' 显示工具设置
Sub YYToolsShowSettings()
    On Error GoTo ErrorHandler
    
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If obj Is Nothing Then
        MsgBox "无法创建YYTools对象，请确保COM组件已正确注册", vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 直接调用ShowSettings方法
    obj.ShowSettings
    
    Set obj = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "显示工具设置时出错: " & Err.Number & " - " & Err.Description, vbCritical, "YYTools错误"
    If Not obj Is Nothing Then Set obj = Nothing
End Sub

' 显示关于信息
Sub YYToolsShowAbout()
    On Error GoTo ErrorHandler
    
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If obj Is Nothing Then
        MsgBox "无法创建YYTools对象，请确保COM组件已正确注册", vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 直接调用ShowAbout方法
    obj.ShowAbout
    
    Set obj = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "显示关于信息时出错: " & Err.Number & " - " & Err.Description, vbCritical, "YYTools错误"
    If Not obj Is Nothing Then Set obj = Nothing
End Sub

' 安装YY工具菜单
Sub InstallYYToolsMenu()
    On Error GoTo ErrorHandler
    
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If obj Is Nothing Then
        MsgBox "无法创建YYTools对象，请确保COM组件已正确注册", vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 安装菜单
    Dim result As String
    result = obj.InstallMenu()
    
    MsgBox result, vbInformation, "YY工具菜单安装"
    
    Set obj = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "安装菜单时出错: " & Err.Number & " - " & Err.Description, vbCritical, "YYTools错误"
    If Not obj Is Nothing Then Set obj = Nothing
End Sub

' 刷新YY工具菜单
Sub RefreshYYToolsMenu()
    On Error GoTo ErrorHandler
    
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If obj Is Nothing Then
        MsgBox "无法创建YYTools对象，请确保COM组件已正确注册", vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 刷新菜单
    obj.RefreshMenu
    
    MsgBox "YY工具菜单已刷新！", vbInformation, "YYTools"
    
    Set obj = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "刷新菜单时出错: " & Err.Number & " - " & Err.Description, vbCritical, "YYTools错误"
    If Not obj Is Nothing Then Set obj = Nothing
End Sub 