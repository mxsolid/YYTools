Attribute VB_Name = "YYTools_TestModule"

' ========================================
' YYTools 测试宏模块 v2.1
' 用于WPS表格中测试YYTools功能
' ========================================

' 运单匹配工具调用宏
Sub YYTools_运单匹配()
    On Error Resume Next
    Dim objYYTools As Object
    Set objYYTools = CreateObject("YYTools.ExcelAddin")
    
    If Err.Number <> 0 Then
        MsgBox "无法创建YYTools对象，请确保已正确注册COM组件。" & vbCrLf & _
               "错误信息: " & Err.Description, vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 调用显示匹配窗体方法
    objYYTools.ShowMatchForm
    
    Set objYYTools = Nothing
End Sub

' 工具设置调用宏
Sub YYTools_设置()
    On Error Resume Next
    Dim objYYTools As Object
    Set objYYTools = CreateObject("YYTools.ExcelAddin")
    
    If Err.Number <> 0 Then
        MsgBox "无法创建YYTools对象，请确保已正确注册COM组件。" & vbCrLf & _
               "错误信息: " & Err.Description, vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 调用显示设置窗体方法
    objYYTools.ShowSettings
    
    Set objYYTools = Nothing
End Sub

' 测试应用程序连接
Sub YYTools_测试连接()
    On Error Resume Next
    Dim objYYTools As Object
    Set objYYTools = CreateObject("YYTools.ExcelAddin")
    
    If Err.Number <> 0 Then
        MsgBox "无法创建YYTools对象，请确保已正确注册COM组件。" & vbCrLf & _
               "错误信息: " & Err.Description, vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 获取应用程序信息
    Dim info As String
    info = objYYTools.GetApplicationInfo()
    
    MsgBox "应用程序信息: " & vbCrLf & info, vbInformation, "YYTools连接测试"
    
    Set objYYTools = Nothing
End Sub

' 获取详细应用程序信息
Sub YYTools_详细信息()
    On Error Resume Next
    Dim objYYTools As Object
    Set objYYTools = CreateObject("YYTools.ExcelAddin")
    
    If Err.Number <> 0 Then
        MsgBox "无法创建YYTools对象，请确保已正确注册COM组件。" & vbCrLf & _
               "错误信息: " & Err.Description, vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 获取详细应用程序信息
    Dim info As String
    info = objYYTools.GetDetailedApplicationInfo()
    
    MsgBox info, vbInformation, "YYTools详细信息"
    
    Set objYYTools = Nothing
End Sub

' 手动安装菜单
Sub YYTools_安装菜单()
    On Error Resume Next
    Dim objYYTools As Object
    Set objYYTools = CreateObject("YYTools.ExcelAddin")
    
    If Err.Number <> 0 Then
        MsgBox "无法创建YYTools对象，请确保已正确注册COM组件。" & vbCrLf & _
               "错误信息: " & Err.Description, vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 安装菜单
    Dim result As String
    result = objYYTools.InstallMenu()
    
    MsgBox result, vbInformation, "YYTools菜单安装"
    
    Set objYYTools = Nothing
End Sub

' 刷新菜单
Sub YYTools_刷新菜单()
    On Error Resume Next
    Dim objYYTools As Object
    Set objYYTools = CreateObject("YYTools.ExcelAddin")
    
    If Err.Number <> 0 Then
        MsgBox "无法创建YYTools对象，请确保已正确注册COM组件。" & vbCrLf & _
               "错误信息: " & Err.Description, vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 刷新菜单
    objYYTools.RefreshMenu
    
    MsgBox "菜单已刷新！", vbInformation, "YYTools"
    
    Set objYYTools = Nothing
End Sub

' 综合测试程序
Sub YYTools_综合测试()
    On Error Resume Next
    Dim objYYTools As Object
    Set objYYTools = CreateObject("YYTools.ExcelAddin")
    
    If Err.Number <> 0 Then
        MsgBox "无法创建YYTools对象，请确保已正确注册COM组件。" & vbCrLf & _
               "错误信息: " & Err.Description, vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    Dim testResult As String
    testResult = "=== YYTools 综合测试结果 ===" & vbCrLf & vbCrLf
    
    ' 测试1：获取详细信息
    testResult = testResult & "1. 应用程序详细信息：" & vbCrLf
    Dim detailInfo As String
    detailInfo = objYYTools.GetDetailedApplicationInfo()
    testResult = testResult & detailInfo & vbCrLf & vbCrLf
    
    ' 测试2：安装菜单
    testResult = testResult & "2. 菜单安装测试：" & vbCrLf
    Dim menuResult As String
    menuResult = objYYTools.InstallMenu()
    testResult = testResult & menuResult & vbCrLf & vbCrLf
    
    ' 测试3：基本信息
    testResult = testResult & "3. 基本应用程序信息：" & vbCrLf
    Dim basicInfo As String
    basicInfo = objYYTools.GetApplicationInfo()
    testResult = testResult & basicInfo & vbCrLf & vbCrLf
    
    testResult = testResult & "=== 测试完成 ==="
    
    MsgBox testResult, vbInformation, "YYTools综合测试"
    
    Set objYYTools = Nothing
End Sub

' 安装YYTools菜单（兼容旧版本）
Sub 安装YYTools菜单()
    Call YYTools_安装菜单
End Sub

' 测试工作簿信息获取
Sub YYTools_测试工作簿()
    On Error Resume Next
    Dim objYYTools As Object
    Set objYYTools = CreateObject("YYTools.ExcelAddin")
    
    If Err.Number <> 0 Then
        MsgBox "无法创建YYTools对象，请确保已正确注册COM组件。" & vbCrLf & _
               "错误信息: " & Err.Description, vbCritical, "YYTools错误"
        Exit Sub
    End If
    
    ' 测试获取打开的工作簿
    Dim workbooks As Object
    Set workbooks = objYYTools.GetOpenWorkbooks()
    
    If workbooks Is Nothing Then
        MsgBox "无法获取工作簿信息", vbWarning, "YYTools测试"
    Else
        MsgBox "成功获取工作簿信息！工作簿数量: " & workbooks.Count, vbInformation, "YYTools测试"
    End If
    
    Set objYYTools = Nothing
End Sub 