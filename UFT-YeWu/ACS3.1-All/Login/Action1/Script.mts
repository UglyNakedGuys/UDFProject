WriteLogs("登录测试模块！")
'初始判断操作 
If SwfWindow("登录界面").Exist(1) Then
	SwfWindow("登录界面").Close()
	WriteLogs("关闭登录界面！")
End If
If SwfWindow("视频识别出入口管理系统").Exist(1) Then
'	SwfWindow("视频识别出入口管理系统").Close()
	SystemUtil.CloseProcessByName("PakingVideo_Login.exe")
	WriteLogs("关闭主界面！")
End If
wait 5
'程序启动操作
SystemUtil.Run("C:\Users\Administrator\Desktop\兔巢acs3.1\软件201705181100-Debug\Debug\PakingVideo_Login.exe")
WriteLogs("启动登录程序！")

'迭代数据表
SwfWindow("登录界面").SwfEdit("SwfEdit").Set Datatable("Name","Action1")
SwfWindow("登录界面").SwfEdit("SwfEdit_2").Set Datatable("Pwd","Action1")

SwfWindow("登录界面").SwfObject("登录").Click

'日志写入
WriteLogs("-------------------------------------------------------")
If SwfWindow("视频识别出入口管理系统").Exist(1) Then
	WriteLogs("用户"&Datatable("Name","Action1")&"登录成功！")
ElseIf SwfWindow("登录界面").SwfWindow("错误信息").Exist(1) Then	
	SwfWindow("登录界面").SwfWindow("错误信息").Close()
	WriteLogs("用户"&Datatable("Name","Action1")&"登录失败，错误信息：账户与密码不匹配！")
Else
	WriteLogs("用户"&Datatable("Name","Action1")&"登录失败，错误信息：未知（用户被屏蔽或者拦截等）！")
End If
WriteLogs("-------------------------------------------------------")