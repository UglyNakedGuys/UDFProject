WriteLogs("===================通道管理模块开始====================")
WriteLogs("前置初始化操作！")
Do While True
If SwfWindow("视频识别出入口管理系统").Exist(1) Then
    Exit Do
End If
Loop 
SwfWindow("视频识别出入口管理系统").SwfObject("设备管理").Click @@ hightlight id_;_198172_;_script infofile_;_ZIP::ssf1.xml_;_
Wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu2").Click @@ hightlight id_;_591216_;_script infofile_;_ZIP::ssf2.xml_;_
Do	While True
If  SwfWindow("停车场通道管理").Exist(1) Then
	Exit Do
End If
Loop
Wait 1
SwfWindow("停车场通道管理").SwfObject("添加(A)").Click
'=======================================================================================================
'通道信息添加
WriteLogs("===================通道信息添加开始====================")
Wait 1 @@ hightlight id_;_656740_;_script infofile_;_ZIP::ssf3.xml_;_
Dim areaInfo
areaInfo=Datatable.GetSheet("通道管理").GetParameter("区域信息").ValueByRow(1)
'Dim Name,x,y
'tempStr=GetLocalPos(areaInfo)
'Name=tempStr(0)
'x=tempStr(1)
'y=tempStr(2)
i=17
basenum=8 @@ hightlight id_;_264064_;_script infofile_;_ZIP::ssf56.xml_;_
'==============================================================================================================
'区域信息--循环读取Combox内容逻辑判断
Do While True
	SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbParkingLotName").Click
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>=103 Then @@ hightlight id_;_2492618_;_script infofile_;_ZIP::ssf80.xml_;_
		SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject_3").Click 10,110 @@ hightlight id_;_2361728_;_script infofile_;_ZIP::ssf81.xml_;_
		SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
	Else
		SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
		basenum=basenum+i
	End If
'检测是否Text相等
Wait 1
'Msgbox SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").GetROProperty("Text")
	If SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbParkingLotName").GetROProperty("Text")=areaInfo Then
		Exit Do
	End If	
Loop 
'================================================================================================================ @@ hightlight id_;_1114330_;_script infofile_;_ZIP::ssf49.xml_;_
Wait 1
Dim ChannelName
ChannelName=Datatable.GetSheet("通道管理").GetParameter("通道名称").ValueByRow(1)
SwfWindow("停车场通道管理").Activate

SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfEdit("SwfEdit").Set ChannelName @@ hightlight id_;_1050654_;_script infofile_;_ZIP::ssf6.xml_;_
Wait 1
Dim InCount,OutCount
InCount=Datatable.GetSheet("通道管理").GetParameter("进场通道数").ValueByRow(1)
OutCount=Datatable.GetSheet("通道管理").GetParameter("出场通道数").ValueByRow(1)

SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("spinInCount").DblClick 5,5,micLeftBtn
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfEdit("SwfEdit_2").Type InCount
Wait 1
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("spinOutCount").DblClick 5,5,micLeftBtn
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfEdit("SwfEdit_3").Type OutCount
Wait 1 @@ hightlight id_;_920696_;_script infofile_;_ZIP::ssf7.xml_;_
Dim computer
computer=Datatable.GetSheet("通道管理").GetParameter("管理电脑").ValueByRow(1)
'tempStr=GetLocalPos(computer)
'x=tempStr(1)
'y=tempStr(2) @@ hightlight id_;_1378156_;_script infofile_;_ZIP::ssf50.xml_;_
'SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject").Click x,y

i=17
basenum=8 @@ hightlight id_;_264064_;_script infofile_;_ZIP::ssf56.xml_;_
'==============================================================================================================
'管理电脑--循环读取Combox内容逻辑判断
Do While True
	SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").Click
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>=103 Then @@ hightlight id_;_2492618_;_script infofile_;_ZIP::ssf80.xml_;_
		SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject_3").Click 10,110 @@ hightlight id_;_2361728_;_script infofile_;_ZIP::ssf81.xml_;_
		SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
	Else
		SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
		basenum=basenum+i
	End If
'检测是否Text相等
Wait 1
'Msgbox SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").GetROProperty("Text")
	If SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").GetROProperty("Text")=computer Then
		Exit Do
	End If	
Loop 
'==============================================================================================================
Wait 1
Dim chargingRule
chargingRule=Datatable.GetSheet("通道管理").GetParameter("收费规则").ValueByRow(1)
i=17
basenum=8
'tempStr=GetLocalPos(chargingRule)
'x=tempStr(1)
'y=tempStr(2) @@ hightlight id_;_3475048_;_script infofile_;_ZIP::ssf52.xml_;_
'SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject").Click x,y @@ hightlight id_;_2296696_;_script infofile_;_ZIP::ssf53.xml_;_

'=======================================================================================================
'收费规则--循环读取Combox内容逻辑判断
Do While True @@ hightlight id_;_985780_;_script infofile_;_ZIP::ssf90.xml_;_
	SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbChargeRule").Click 67,6
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>=102 Then @@ hightlight id_;_1706128_;_script infofile_;_ZIP::ssf83.xml_;_
		SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject_3").Click 10,110 @@ hightlight id_;_1706128_;_script infofile_;_ZIP::ssf85.xml_;_
		SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
	Else @@ hightlight id_;_1904188_;_script infofile_;_ZIP::ssf112.xml_;_
		SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
		basenum=basenum+i
	End If
'检测是否Text相等
Wait 1
'Msgbox SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbChargeRule").GetROProperty("Text")
	If SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbChargeRule").GetROProperty("Text")=chargingRule Then
		Exit Do
	End If	
Loop 
'========================================================================================================
Wait 1
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("保存").Click @@ hightlight id_;_591048_;_script infofile_;_ZIP::ssf11.xml_;_

passFlag=False
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("提示信息").SwfObject("OK").Click
If SwfWindow("停车场通道管理").Exist(1) Then
	passFlag=True
	WriteLogs("通道添加返回====成功！")
End If

'数据表信息写入
If passFlag Then
	reporter.ReportEvent micPass,"Add","添加成功！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("通道管理").SetCurrentRow(1)
	datatable.Value("添加结果","通道管理")="成功"
	WriteLogs("数据表修改成功")
else
	reporter.ReportEvent  micFail ,"Add","添加失败！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("通道管理").SetCurrentRow(1)
	datatable.Value("添加结果","通道管理")="失败"
End If
Wait 2
'======================================================================================================
'通道信息修改
WriteLogs("===================通道信息修改开始====================")
Dim tempChannelName
tempChannelName=Datatable.GetSheet("通道管理").GetParameter("通道名称").ValueByRow(1)
For Iterator = 0 To SwfWindow("停车场通道管理").SwfTable("gridControl1").RowCount-1
	If tempChannelName=SwfWindow("停车场通道管理").SwfTable("gridControl1").GetCellData(Iterator,0) Then
		SwfWindow("停车场通道管理").SwfTable("gridControl1").ActivateCell Iterator,0
	End If
Next
Do While True
	If SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").Exist(1) Then
		Exit Do
	End If
Loop
Wait 1
Dim ChannelNameEdit
ChannelNameEdit=Datatable.GetSheet("通道管理").GetParameter("修改通道名称").ValueByRow(1)

SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfEdit("SwfEdit").Set ChannelNameEdit
Wait 1
Dim ChargingRuleEdit
ChargingRuleEdit=Datatable.GetSheet("通道管理").GetParameter("修改收费规则").ValueByRow(1)
i=17
basenum=8 @@ hightlight id_;_2296696_;_script infofile_;_ZIP::ssf53.xml_;_

'=======================================================================================================
'收费规则--循环读取Combox内容逻辑判断
Do While True @@ hightlight id_;_985780_;_script infofile_;_ZIP::ssf90.xml_;_
	SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbChargeRule").Click 67,6
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>=102 Then @@ hightlight id_;_1706128_;_script infofile_;_ZIP::ssf83.xml_;_
		SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject_3").Click 10,110 @@ hightlight id_;_1706128_;_script infofile_;_ZIP::ssf85.xml_;_
		SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
	Else @@ hightlight id_;_1904188_;_script infofile_;_ZIP::ssf112.xml_;_
		SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
		basenum=basenum+i
	End If
'检测是否Text相等
Wait 1
	If SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbChargeRule").GetROProperty("Text")=ChargingRuleEdit Then
		Exit Do
	End If	
Loop 
'======================================================================================================== @@ hightlight id_;_2296040_;_script infofile_;_ZIP::ssf45.xml_;_
Wait 1
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("保存").Click

If SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("提示信息").Exist(1) Then
	SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("提示信息").SwfObject("OK").Click
End If
passFlag=False
If SwfWindow("停车场通道管理").Exist(1) Then
	passFlag=True
	WriteLogs("通道修改返回====成功！")
End If

'数据表信息写入
If passFlag Then
	reporter.ReportEvent micPass,"Edit","修改成功！"
	datatable.LocalSheet.AddParameter "修改结果"," "
	datatable.GetSheet("通道管理").SetCurrentRow(1)
	datatable.Value("修改结果","通道管理")="成功"
	WriteLogs("数据表修改成功")
else
	reporter.ReportEvent  micFail ,"Edit","修改失败！"
	datatable.LocalSheet.AddParameter "修改结果"," "
	datatable.GetSheet("通道管理").SetCurrentRow(1)
	datatable.Value("修改结果","通道管理")="失败"
End If
Wait 2
'=====================================================================================================
'通道信息删除
WriteLogs("===================通道信息删除开始====================")

Dim deleteChannelName
deleteChannelName=datatable.GetSheet("通道管理").GetParameter("修改通道名称").ValueByRow(1)

For Iterator = 0 To SwfWindow("停车场通道管理").SwfTable("gridControl1").RowCount-1
If deleteChannelName=SwfWindow("停车场通道管理").SwfTable("gridControl1").GetCellData(Iterator,0) Then
	SwfWindow("停车场通道管理").SwfTable("gridControl1").SelectCell Iterator,0
	SwfWindow("停车场通道管理").SwfObject("删除(D)").Click
End If
Next
Do While True
	If SwfWindow("停车场通道管理").SwfWindow("确认信息").Exist(1) Then
		Exit Do
	End If
Loop
Wait 1
SwfWindow("停车场通道管理").SwfWindow("确认信息").SwfObject("Yes").Click

If SwfWindow("停车场通道管理").SwfWindow("提示信息").Exist(1) Then
	WriteLogs("删除通道返回====成功！")	
	Wait 1
	SwfWindow("停车场通道管理").SwfWindow("提示信息").SwfObject("OK").Click
End If

reporter.ReportEvent micPass,"Delete","修改成功！"
datatable.LocalSheet.AddParameter "删除结果"," "
datatable.GetSheet("通道管理").SetCurrentRow(1)
datatable.Value("删除结果","通道管理")="成功"
WriteLogs("数据表修改成功")
Wait 2
SwfWindow("停车场通道管理").Close()
WriteLogs("===================通道管理模块结束====================")

Wait 1