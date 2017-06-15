WriteLogs("====================操作员模块开始=====================")
WriteLogs("前置初始化操作！")
Do While True
If SwfWindow("视频识别出入口管理系统").Exist(1) Then
    Exit Do
End If
Loop 
SwfWindow("视频识别出入口管理系统").SwfObject("系统设置").Click
Wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu2").Click @@ hightlight id_;_591216_;_script infofile_;_ZIP::ssf2.xml_;_
Do	While True
If  SwfWindow("操作员管理").Exist(1) Then
	Exit Do
End If
Loop
Wait 1
SwfWindow("操作员管理").SwfObject("添加(A)").Click
'======================================================================================================
'操作员信息添加
WriteLogs("==================操作员信息添加开始===================")
Wait 1
Dim Name
Name=Datatable.GetSheet("操作员管理").GetParameter("姓名").ValueByRow(1) @@ hightlight id_;_2689152_;_script infofile_;_ZIP::ssf24.xml_;_
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit").Set Name @@ hightlight id_;_3147492_;_script infofile_;_ZIP::ssf25.xml_;_
Wait 1
 @@ hightlight id_;_5377084_;_script infofile_;_ZIP::ssf39.xml_;_
Dim Role
Role=Datatable.GetSheet("操作员管理").GetParameter("所属角色").ValueByRow(1)
i=17
basenum=8 @@ hightlight id_;_395522_;_script infofile_;_ZIP::ssf58.xml_;_
'==============================================================================================================
'所属角色信息--循环读取Combox内容逻辑判断
Do While True
	SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("cmbRole").Click
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>110 Then @@ hightlight id_;_8917212_;_script infofile_;_ZIP::ssf45.xml_;_
		SwfWindow("区域管理").SwfWindow("区域设置").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 10,110
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 60,basenum-i
	Else
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 60,basenum
		If basenum<=110 Then
			basenum=basenum+i
		End If
	End If
'检测是否Text相等
Wait 1
	If SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("cmbRole").GetROProperty("Text")=Role Then
		Exit Do
	End If	
Loop 
'================================================================================================================
Wait 1 @@ hightlight id_;_6160832_;_script infofile_;_ZIP::ssf34.xml_;_
Dim loginName
loginName=Datatable.GetSheet("操作员管理").GetParameter("登录名").ValueByRow(1)
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit_2").Set loginName @@ hightlight id_;_1510704_;_script infofile_;_ZIP::ssf40.xml_;_
Wait 1
Dim loginPwd
loginPwd=Datatable.GetSheet("操作员管理").GetParameter("登录密码").ValueByRow(1)
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit_3").Set loginPwd @@ hightlight id_;_658632_;_script infofile_;_ZIP::ssf41.xml_;_
Wait 1
Dim web,flag
web=Datatable.GetSheet("操作员管理").GetParameter("是否Web减免").ValueByRow(1)
tempStr=GetLocalPos(web)
web=tempStr(0)
flag=tempStr(1)
Select Case web
Case "是"
	SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("chkQueryWeb").Click @@ hightlight id_;_1968546_;_script infofile_;_ZIP::ssf45.xml_;_
	SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("MerchantType").Click
	If flag="时间" Then
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 71,9
	Else
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 66,23
	End If @@ hightlight id_;_2883854_;_script infofile_;_ZIP::ssf48.xml_;_
Case Else
End Select
Wait 1 @@ hightlight id_;_1968546_;_script infofile_;_ZIP::ssf42.xml_;_
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("保存(S)").Click @@ hightlight id_;_854858_;_script infofile_;_ZIP::ssf43.xml_;_
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("提示信息").SwfObject("OK").Click
passFlag=false
If Not SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("提示信息").Exist(1) Then
	passFlag=true
	WriteLogs("添加操作员返回===成功！")
Else
	WriteLogs("添加操作员返回===失败！")
End If
If passFlag Then
	reporter.ReportEvent micPass,"Add","添加成功！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("操作员管理").SetCurrentRow(1)
	datatable.Value("添加结果","操作员管理")="成功"
	WriteLogs("数据表修改成功")
Else
	reporter.ReportEvent micPass,"Add","添加失败！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("操作员管理").SetCurrentRow(1)
	datatable.Value("添加结果","操作员管理")="失败"
	WriteLogs("数据表修改成功")
End If
'====================================================================================================
 '操作员信息修改
WriteLogs("==================操作员信息修改开始===================")
Wait 1
Dim tempName
tempName=Datatable.GetSheet("操作员管理").GetParameter("姓名").ValueByRow(1)
For Iterator = 0 To SwfWindow("操作员管理").SwfTable("gridControlOperator").RowCount-1
	If tempName=SwfWindow("操作员管理").SwfTable("gridControlOperator").GetCellData(Iterator,1) Then
		SwfWindow("操作员管理").SwfTable("gridControlOperator").ActivateCell Iterator,1
	End If
Next
Do While True
	If SwfWindow("操作员管理").SwfWindow("操作员编辑").Exist(1) Then
		Exit Do
	End If
Loop
Wait 1
Name=Datatable.GetSheet("操作员管理").GetParameter("修改姓名").ValueByRow(1) @@ hightlight id_;_2689152_;_script infofile_;_ZIP::ssf24.xml_;_
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit").Set Name @@ hightlight id_;_3147492_;_script infofile_;_ZIP::ssf25.xml_;_
Wait 1
Role=Datatable.GetSheet("操作员管理").GetParameter("修改所属角色").ValueByRow(1)
i=17
basenum=8 @@ hightlight id_;_395522_;_script infofile_;_ZIP::ssf58.xml_;_
'==============================================================================================================
'所属角色信息--循环读取Combox内容逻辑判断
Do While True
	SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("cmbRole").Click
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>110 Then @@ hightlight id_;_8917212_;_script infofile_;_ZIP::ssf45.xml_;_
		SwfWindow("区域管理").SwfWindow("区域设置").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 10,110
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 60,basenum-i
	Else
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 60,basenum
		If basenum<=110 Then
			basenum=basenum+i
		End If
	End If
'检测是否Text相等
Wait 1
	If SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("cmbRole").GetROProperty("Text")=Role Then
		Exit Do
	End If	
Loop 
'================================================================================================================ @@ hightlight id_;_1116034_;_script infofile_;_ZIP::ssf26.xml_;_
Wait 1
loginName=Datatable.GetSheet("操作员管理").GetParameter("修改登录名").ValueByRow(1)
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit_2").Set loginName @@ hightlight id_;_1510704_;_script infofile_;_ZIP::ssf40.xml_;_
Wait 1
loginPwd=Datatable.GetSheet("操作员管理").GetParameter("修改登录密码").ValueByRow(1)
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit_3").Set loginPwd @@ hightlight id_;_658632_;_script infofile_;_ZIP::ssf41.xml_;_
Wait 1
web=Datatable.GetSheet("操作员管理").GetParameter("修改是否Web减免").ValueByRow(1)
tempStr=GetLocalPos(web)
web=tempStr(0)
flag=tempStr(1)
Select Case web
Case "是"
	SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("chkQueryWeb").Click @@ hightlight id_;_1968546_;_script infofile_;_ZIP::ssf45.xml_;_
	SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("MerchantType").Click
	If flag="时间" Then
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 71,9
	Else
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 66,23
	End If
Case Else
End Select
Wait 1 @@ hightlight id_;_1968546_;_script infofile_;_ZIP::ssf42.xml_;_
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("保存(S)").Click @@ hightlight id_;_854858_;_script infofile_;_ZIP::ssf43.xml_;_
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("提示信息").SwfObject("OK").Click
passFlag=false
If Not SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("提示信息").Exist(1) Then
	passFlag=true
	WriteLogs("修改操作员返回===成功！")
Else
	WriteLogs("修改操作员返回===失败！")
End If
If passFlag Then
	reporter.ReportEvent micPass,"Add","修改成功！"
	datatable.LocalSheet.AddParameter "修改结果"," "
	datatable.GetSheet("操作员管理").SetCurrentRow(1)
	datatable.Value("修改结果","操作员管理")="成功"
	WriteLogs("数据表修改成功")
Else
	reporter.ReportEvent micPass,"Add","修改失败！"
	datatable.LocalSheet.AddParameter "修改结果"," "
	datatable.GetSheet("操作员管理").SetCurrentRow(1)
	datatable.Value("修改结果","操作员管理")="失败"
	WriteLogs("数据表修改成功")
End If
Wait 2
'=========================================================================================
WriteLogs("==================操作员信息删除开始===================")
'操作员信息删除
Dim deleteName
deleteName=Datatable.GetSheet("操作员管理").GetParameter("修改姓名").ValueByRow(1)

For Iterator = 0 To SwfWindow("操作员管理").SwfTable("gridControlOperator").RowCount-1
If deleteName=SwfWindow("操作员管理").SwfTable("gridControlOperator").GetCellData(Iterator,1) Then
	SwfWindow("操作员管理").SwfTable("gridControlOperator").SelectCell Iterator,1
	SwfWindow("操作员管理").SwfObject("删除(D)").Click
End If
Next
Do While True
	If SwfWindow("操作员管理").SwfWindow("确认信息").Exist(1) Then
		Exit Do
	End If
Loop
Wait 1
SwfWindow("操作员管理").SwfWindow("确认信息").SwfObject("Yes").Click

If SwfWindow("操作员管理").SwfWindow("提示信息").Exist(1) Then
	SwfWindow("操作员管理").SwfWindow("提示信息").SwfObject("OK").Click
	If Not SwfWindow("操作员管理").SwfWindow("提示信息").Exist(1) Then
		WriteLogs("删除操作员返回===成功！")	
	Else
		WriteLogs("删除操作员返回====失败！")
	End If
Else
	WriteLogs("删除操作员返回====失败！")
End If

reporter.ReportEvent micPass,"Delete","修改成功！"
datatable.LocalSheet.AddParameter "删除结果"," "
datatable.GetSheet("操作员管理").SetCurrentRow(1)
datatable.Value("删除结果","操作员管理")="成功"
WriteLogs("数据表修改成功")
'===================================================================================================
Wait 2
SwfWindow("操作员管理").Close()

WriteLogs("==================操作员管理模块结束===================")
Wait 2