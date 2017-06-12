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
Dim Name,x,y
tempStr=GetLocalPos(areaInfo)
Name=tempStr(0)
x=tempStr(1)
y=tempStr(2)
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbParkingLotName").Click @@ hightlight id_;_2098790_;_script infofile_;_ZIP::ssf4.xml_;_
SwfWindow("SwfWindow").SwfObject("SwfObject").Click x,y @@ hightlight id_;_985206_;_script infofile_;_ZIP::ssf5.xml_;_


SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbParkingLotName").Click 115,14 @@ hightlight id_;_3605860_;_script infofile_;_ZIP::ssf48.xml_;_
SwfWindow("SwfWindow").SwfObject("SwfObject").Click 62,25 @@ hightlight id_;_1114330_;_script infofile_;_ZIP::ssf49.xml_;_
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
Wait 1
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").Click @@ hightlight id_;_920696_;_script infofile_;_ZIP::ssf7.xml_;_
Dim computer
computer=Datatable.GetSheet("通道管理").GetParameter("管理电脑").ValueByRow(1)
tempStr=GetLocalPos(computer)
x=tempStr(1)
y=tempStr(2) @@ hightlight id_;_1378156_;_script infofile_;_ZIP::ssf50.xml_;_
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject").Click x,y @@ hightlight id_;_3606964_;_script infofile_;_ZIP::ssf51.xml_;_
Wait 1 @@ hightlight id_;_1574544_;_script infofile_;_ZIP::ssf8.xml_;_
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbChargeRule").Click
Dim chargingRule
chargingRule=Datatable.GetSheet("通道管理").GetParameter("收费规则").ValueByRow(1)
tempStr=GetLocalPos(chargingRule)
x=tempStr(1)
y=tempStr(2) @@ hightlight id_;_3475048_;_script infofile_;_ZIP::ssf52.xml_;_
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject").Click x,y @@ hightlight id_;_2296696_;_script infofile_;_ZIP::ssf53.xml_;_
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
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbChargeRule").Click 72,7 @@ hightlight id_;_4524666_;_script infofile_;_ZIP::ssf44.xml_;_
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 45,44 @@ hightlight id_;_2296040_;_script infofile_;_ZIP::ssf45.xml_;_
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
