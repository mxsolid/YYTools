echo Set oWS = WScript.CreateObject("WScript.Shell") 
echo sLinkFile = "%SHORTCUT_PATH%" 
echo Set oLink = oWS.CreateShortcut(sLinkFile) 
echo oLink.TargetPath = "%TARGET_DIR%\TestProgram.exe" 
echo oLink.Description = "YY运单匹配工具 v1.5" 
echo oLink.Save 
