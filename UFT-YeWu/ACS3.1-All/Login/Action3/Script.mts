'初始化关闭
'SystemUtil.CloseProcessByName("iexplore.exe")

'启动浏览器方式1
SystemUtil.Run "iexplore.exe","http://192.168.10.185"

'启动浏览器方式2
'Set ie=CreateObject("InternetExplorer.Application")
'ie.Visible=true
'Wait 3
'ie.navigate("http://192.168.10.185")
'Set ie=Nothing

'启动浏览器方式3 WScript脚本宿主模式
'Dim oWsh
'Set  oWsh=WScript.CreateObject(WScript.Shell)
'oWsh.Run("http://www.baidu.com")
'Set oWsh=Nothing
Wait 3
Browser("登录").Page("登录").WebEdit("WebEdit").Set "admin"
Browser("登录").Page("登录").WebEdit("WebEdit_2").Set "admin"
Wait 1
Browser("登录").Page("登录").Link("登录").Click

Wait 1
Browser("登录").Page("登录").WebCheckBox("WebCheckBox").Click
Wait 1
Browser("登录").Page("登录").Link("Link").Click
Msgbox "进场成功"

