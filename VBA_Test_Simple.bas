Attribute VB_Name = "YYTools_Simple_Test"

' ========================================
' YYTools 简化测试模块
' 用于解决VBA调用问题
' ========================================

' 最简单的测试
Sub Test1_CreateObject()
    On Error GoTo ErrorHandler
    
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If obj Is Nothing Then
        MsgBox "创建COM对象失败", vbCritical
    Else
        MsgBox "创建COM对象成功！", vbInformation
        Set obj = Nothing
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "错误: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' 测试基本方法调用
Sub Test2_BasicMethod()
    On Error GoTo ErrorHandler
    
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If obj Is Nothing Then
        MsgBox "创建COM对象失败", vbCritical
        Exit Sub
    End If
    
    ' 尝试调用GetApplicationInfo方法
    Dim info As String
    info = obj.GetApplicationInfo()
    
    MsgBox "应用程序信息: " & vbCrLf & info, vbInformation
    
    Set obj = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "调用方法错误: " & Err.Number & " - " & Err.Description, vbCritical
    If Not obj Is Nothing Then Set obj = Nothing
End Sub

' 测试详细信息方法
Sub Test3_DetailedInfo()
    On Error GoTo ErrorHandler
    
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If obj Is Nothing Then
        MsgBox "创建COM对象失败", vbCritical
        Exit Sub
    End If
    
    ' 尝试调用GetDetailedApplicationInfo方法
    Dim info As String
    info = obj.GetDetailedApplicationInfo()
    
    MsgBox info, vbInformation, "详细应用程序信息"
    
    Set obj = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "调用详细信息方法错误: " & Err.Number & " - " & Err.Description, vbCritical
    If Not obj Is Nothing Then Set obj = Nothing
End Sub

' 测试菜单安装
Sub Test4_InstallMenu()
    On Error GoTo ErrorHandler
    
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If obj Is Nothing Then
        MsgBox "创建COM对象失败", vbCritical
        Exit Sub
    End If
    
    ' 尝试安装菜单
    Dim result As String
    result = obj.InstallMenu()
    
    MsgBox "菜单安装结果: " & vbCrLf & result, vbInformation
    
    Set obj = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "安装菜单错误: " & Err.Number & " - " & Err.Description, vbCritical
    If Not obj Is Nothing Then Set obj = Nothing
End Sub

' 测试匹配窗体显示
Sub Test5_ShowMatchForm()
    On Error GoTo ErrorHandler
    
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If obj Is Nothing Then
        MsgBox "创建COM对象失败", vbCritical
        Exit Sub
    End If
    
    ' 显示匹配窗体
    obj.ShowMatchForm
    
    MsgBox "匹配窗体调用完成", vbInformation
    
    Set obj = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "显示匹配窗体错误: " & Err.Number & " - " & Err.Description, vbCritical
    If Not obj Is Nothing Then Set obj = Nothing
End Sub

' 完整测试流程
Sub TestAll()
    MsgBox "开始完整测试流程", vbInformation
    
    Call Test1_CreateObject
    Call Test2_BasicMethod
    Call Test3_DetailedInfo
    Call Test4_InstallMenu
    Call Test5_ShowMatchForm
    
    MsgBox "完整测试流程结束", vbInformation
End Sub

' 诊断函数
Sub Diagnose()
    On Error Resume Next
    
    Dim result As String
    result = "YYTools 诊断结果:" & vbCrLf & vbCrLf
    
    ' 测试COM对象创建
    Dim obj As Object
    Set obj = CreateObject("YYTools.ExcelAddin")
    
    If Err.Number <> 0 Then
        result = result & "1. COM对象创建: 失败" & vbCrLf
        result = result & "   错误: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf
        result = result & "可能原因:" & vbCrLf
        result = result & "- YYTools.dll未注册" & vbCrLf
        result = result & "- 需要管理员权限重新注册" & vbCrLf
        result = result & "- 运行 install_admin.bat" & vbCrLf
    Else
        result = result & "1. COM对象创建: 成功" & vbCrLf & vbCrLf
        
        ' 测试基本方法
        Err.Clear
        Dim info As String
        info = obj.GetApplicationInfo()
        
        If Err.Number <> 0 Then
            result = result & "2. 基本方法调用: 失败" & vbCrLf
            result = result & "   错误: " & Err.Number & " - " & Err.Description & vbCrLf
        Else
            result = result & "2. 基本方法调用: 成功" & vbCrLf
            result = result & "   信息: " & info & vbCrLf
        End If
        
        Set obj = Nothing
    End If
    
    MsgBox result, vbInformation, "YYTools 诊断"
End Sub 