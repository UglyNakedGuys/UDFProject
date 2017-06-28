WriteLogs("==================进出口管理模块开始===================")
WriteLogs("前置初始化操作！")
Do While True
If SwfWindow("视频识别出入口管理系统").Exist(1) Then
    Exit Do
End If
Loop 
SwfWindow("视频识别出入口管理系统").SwfObject("设备管理").Click @@ hightlight id_;_198172_;_script infofile_;_ZIP::ssf1.xml_;_
Wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu3").Click @@ hightlight id_;_3999384_;_script infofile_;_ZIP::ssf96.xml_;_
Do	While True
If  SwfWindow("进出口管理").Exist(1) Then
	Exit Do
End If
Loop
Wait 1
SwfWindow("进出口管理").SwfObject("添加(A)").Click
'=======================================================================================================
'进出口信息添加
WriteLogs("==================进出口信息添加开始===================")
Wait 1 @@ hightlight id_;_2492028_;_script infofile_;_ZIP::ssf15.xml_;_
Dim channelName
channelName=Datatable.GetSheet("进出口管理").GetParameter("通道名称").ValueByRow(1)
'==============================================================================================================
'区域信息--循环读取Combox内容逻辑判断
i=17
basenum=8
Do While True
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbParkingChannelName").Click
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>110 Then @@ hightlight id_;_2492618_;_script infofile_;_ZIP::ssf80.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_3").Click 10,110 @@ hightlight id_;_2361728_;_script infofile_;_ZIP::ssf81.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum-17
	Else @@ hightlight id_;_854240_;_script infofile_;_ZIP::ssf81.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
		If basenum<=110 Then
			basenum=basenum+i
		End If
	End If
'检测是否Text相等
Wait 1
'Msgbox SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").GetROProperty("Text")
	If SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbParkingChannelName").GetROProperty("Text")=channelName Then
		Exit Do
	End If	
Loop 
'================================================================================================================

Wait 1
Dim InOutName
InOutName=Datatable.GetSheet("进出口管理").GetParameter("进出口名称").ValueByRow(1)
SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfEdit("SwfEdit").Set InOutName @@ hightlight id_;_2229446_;_script infofile_;_ZIP::ssf17.xml_;_
Wait 1
Dim InOutType @@ hightlight id_;_3736788_;_script infofile_;_ZIP::ssf18.xml_;_
InOutType=Datatable.GetSheet("进出口管理").GetParameter("进出类型").ValueByRow(1)
Select Case InOutType
Case "进"
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("rdoInOut").Click 13,7
Case Else 
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("rdoInOut").Click 169,12
End Select @@ hightlight id_;_28837524_;_script infofile_;_ZIP::ssf34.xml_;_
Wait 1 @@ hightlight id_;_3736788_;_script infofile_;_ZIP::ssf19.xml_;_
Dim computer
computer=Datatable.GetSheet("进出口管理").GetParameter("管理电脑").ValueByRow(1) @@ hightlight id_;_4655558_;_script infofile_;_ZIP::ssf37.xml_;_
'==============================================================================================================
'管理电脑信息--循环读取Combox内容逻辑判断
i=17
basenum=8
Do While True
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbMStation").Click
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>110 Then @@ hightlight id_;_2492618_;_script infofile_;_ZIP::ssf80.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_3").Click 10,110 @@ hightlight id_;_2361728_;_script infofile_;_ZIP::ssf81.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum-17
	Else
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
		If basenum<=110 Then
			basenum=basenum+i
		End If
	End If
'检测是否Text相等
Wait 1
'Msgbox SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").GetROProperty("Text")
	If SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbMStation").GetROProperty("Text")=computer Then
		Exit Do
	End If	
Loop 
'================================================================================================================
Wait 1 @@ hightlight id_;_1376484_;_script infofile_;_ZIP::ssf22.xml_;_
Dim chargingRule
chargingRule=Datatable.GetSheet("进出口管理").GetParameter("收费规则").ValueByRow(1)
'==============================================================================================================
'收费规则信息--循环读取Combox内容逻辑判断
i=17
basenum=8
Do While True
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbChargeRule").Click
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>110 Then @@ hightlight id_;_2492618_;_script infofile_;_ZIP::ssf80.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_3").Click 10,110 @@ hightlight id_;_2361728_;_script infofile_;_ZIP::ssf81.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum-17
	Else
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
		If basenum<=110 Then
			basenum=basenum+i
		End If
	End If
'检测是否Text相等
Wait 1
'Msgbox SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").GetROProperty("Text")
	If SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbChargeRule").GetROProperty("Text")=chargingRule Then
		Exit Do
	End If	
Loop 
'================================================================================================================
Wait 1 @@ hightlight id_;_1573076_;_script infofile_;_ZIP::ssf24.xml_;_
Dim camera
camera=Datatable.GetSheet("进出口管理").GetParameter("主相机").ValueByRow(1)
'==============================================================================================================
'相机信息--循环读取Combox内容逻辑判断
i=17
basenum=8
Do While True
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbCameraList").Click
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>110 Then @@ hightlight id_;_2492618_;_script infofile_;_ZIP::ssf80.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_3").Click 10,110 @@ hightlight id_;_2361728_;_script infofile_;_ZIP::ssf81.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum-17
	Else
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
		If basenum<=110 Then
			basenum=basenum+i
		End If
	End If
'检测是否Text相等
Wait 1
'Msgbox SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").GetROProperty("Text")
	If SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbCameraList").GetROProperty("Text")=camera Then
		Exit Do
	End If	
Loop 
'================================================================================================================ @@ hightlight id_;_131258_;_script infofile_;_ZIP::ssf29.xml_;_
Wait 1
SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("保存").Click @@ hightlight id_;_3409802_;_script infofile_;_ZIP::ssf38.xml_;_
SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("提示信息").SwfObject("OK").Click 36,12 @@ hightlight id_;_1900944_;_script infofile_;_ZIP::ssf39.xml_;_
passFlag=False
If Not SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("提示信息").Exist(1) Then
	WriteLogs("添加进出口返回===成功！")
	passFlag=True
Else
	WriteLogs("添加进出口返回===失败！")	
End If
If passFlag Then
	reporter.ReportEvent micPass,"Add","添加成功！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("进出口管理").SetCurrentRow(1)
	datatable.Value("添加结果","进出口管理")="成功"
	WriteLogs("数据表修改成功")
Else
	reporter.ReportEvent micPass,"Add","添加失败！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("进出口管理").SetCurrentRow(1)
	datatable.Value("添加结果","进出口管理")="失败"
	WriteLogs("数据表修改成功")
End If
Wait 2
'====================================================================================================
'进出口信息修改
WriteLogs("==================进出口信息修改开始===================")
Wait 1
Dim tempInOutName
tempInOutName=Datatable.GetSheet("进出口管理").GetParameter("进出口名称").ValueByRow(1)
For Iterator = 0 To SwfWindow("进出口管理").SwfTable("gridControl1").RowCount-1
	If tempInOutName=SwfWindow("进出口管理").SwfTable("gridControl1").GetCellData(Iterator,0) Then
		SwfWindow("进出口管理").SwfTable("gridControl1").ActivateCell Iterator,0
		Exit For
	End If
Next
Do While True
	If SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").Exist(1) Then
		Exit Do
	End If
Loop
Wait 1
channelName=Datatable.GetSheet("进出口管理").GetParameter("修改通道名称").ValueByRow(1)
'==============================================================================================================
'区域信息--循环读取Combox内容逻辑判断
i=17
basenum=8
Do While True
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbParkingChannelName").Click
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>110 Then @@ hightlight id_;_2492618_;_script infofile_;_ZIP::ssf80.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 10,110 @@ hightlight id_;_2361728_;_script infofile_;_ZIP::ssf81.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 60,basenum-17
	Else
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 60,basenum
		If basenum<=110 Then
			basenum=basenum+i
		End If
	End If
'检测是否Text相等
Wait 1
'Msgbox SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").GetROProperty("Text")
	If SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbParkingChannelName").GetROProperty("Text")=channelName Then
		Exit Do
	End If	
Loop 
'================================================================================================================
Wait 1
InOutName=Datatable.GetSheet("进出口管理").GetParameter("修改进出口名称").ValueByRow(1)
SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfEdit("SwfEdit").Set InOutName @@ hightlight id_;_2229446_;_script infofile_;_ZIP::ssf17.xml_;_
Wait 1 @@ hightlight id_;_3736788_;_script infofile_;_ZIP::ssf18.xml_;_
InOutType=Datatable.GetSheet("进出口管理").GetParameter("修改进出类型").ValueByRow(1)
Select Case InOutType
Case "进"
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("rdoInOut").Click 13,7
Case Else 
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("rdoInOut").Click 169,12
End Select @@ hightlight id_;_28837524_;_script infofile_;_ZIP::ssf34.xml_;_
Wait 1
computer=Datatable.GetSheet("进出口管理").GetParameter("修改管理电脑").ValueByRow(1)
'==============================================================================================================
'管理电脑信息--循环读取Combox内容逻辑判断
i=17
basenum=8
Do While True
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbMStation").Click
	Wait 1
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>110 Then @@ hightlight id_;_2492618_;_script infofile_;_ZIP::ssf80.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_3").Click 10,110 @@ hightlight id_;_2361728_;_script infofile_;_ZIP::ssf81.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum-17
		Wait 1
	Else
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
		Wait 1
		If basenum<=110 Then
			basenum=basenum+i
		End If
	End If
'检测是否Text相等
Wait 1
'Msgbox SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").GetROProperty("Text")
	If SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbMStation").GetROProperty("Text")=computer Then
		Exit Do
	End If
Loop 
'================================================================================================================ @@ hightlight id_;_4655558_;_script infofile_;_ZIP::ssf37.xml_;_
 @@ hightlight id_;_1114316_;_script infofile_;_ZIP::ssf21.xml_;_
Wait 1
chargingRule=Datatable.GetSheet("进出口管理").GetParameter("修改收费规则").ValueByRow(1)
'==============================================================================================================
'收费规则信息--循环读取Combox内容逻辑判断
i=17
basenum=8
Do While True
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbChargeRule").Click
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>110 Then @@ hightlight id_;_2492618_;_script infofile_;_ZIP::ssf80.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_3").Click 10,110 @@ hightlight id_;_2361728_;_script infofile_;_ZIP::ssf81.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum-17
	Else
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
		If basenum<=110 Then
			basenum=basenum+i
		End If
	End If
'检测是否Text相等
Wait 1
'Msgbox SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").GetROProperty("Text")
	If SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbChargeRule").GetROProperty("Text")=chargingRule Then
		Exit Do
	End If	
Loop 
'================================================================================================================
Wait 1
camera=Datatable.GetSheet("进出口管理").GetParameter("修改主相机").ValueByRow(1)
'==============================================================================================================
'相机信息--循环读取Combox内容逻辑判断
i=17
basenum=8
Do While True
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbCameraList").Click
'判断是否处于边界处 @@ hightlight id_;_1509658_;_script infofile_;_ZIP::ssf79.xml_;_
	If basenum>110 Then @@ hightlight id_;_2492618_;_script infofile_;_ZIP::ssf80.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_3").Click 10,110 @@ hightlight id_;_2361728_;_script infofile_;_ZIP::ssf81.xml_;_
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum-17
	Else
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject_2").Click 60,basenum
		If basenum<=110 Then
			basenum=basenum+i
		End If
	End If
'检测是否Text相等
Wait 1
'Msgbox SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").GetROProperty("Text")
	If SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbCameraList").GetROProperty("Text")=camera Then
		Exit Do
	End If	
Loop 
'================================================================================================================

SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("保存").Click @@ hightlight id_;_4981990_;_script infofile_;_ZIP::ssf42.xml_;_
SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("提示信息").SwfObject("OK").Click @@ hightlight id_;_3410808_;_script infofile_;_ZIP::ssf43.xml_;_
passFlag=false
If Not SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("提示信息").Exist(1) Then
	passFlag=True
	WriteLogs("进出口修改返回====成功！")
Else
	WriteLogs("进出口修改返回====失败！")	
End If
'数据表信息写入
If passFlag Then
	reporter.ReportEvent micPass,"Edit","修改成功！"
	datatable.LocalSheet.AddParameter "修改结果"," "
	datatable.GetSheet("进出口管理").SetCurrentRow(1)
	datatable.Value("修改结果","进出口管理")="成功"
	WriteLogs("数据表修改成功")
else
	reporter.ReportEvent  micFail ,"Edit","修改失败！"
	datatable.LocalSheet.AddParameter "修改结果"," "
	datatable.GetSheet("进出口管理").SetCurrentRow(1)
	datatable.Value("修改结果","进出口管理")="失败"
End If
Wait 2
'========================================================================================
WriteLogs("==================进出口信息删除开始===================")
'进出口信息删除
Dim deleteInOutName
deleteInOutName=datatable.GetSheet("进出口管理").GetParameter("修改进出口名称").ValueByRow(1)

For Iterator = 0 To SwfWindow("进出口管理").SwfTable("gridControl1").RowCount-1
If deleteInOutName=SwfWindow("进出口管理").SwfTable("gridControl1").GetCellData(Iterator,0) Then
	SwfWindow("进出口管理").SwfTable("gridControl1").SelectCell Iterator,0
	SwfWindow("进出口管理").SwfObject("删除(D)").Click
	Exit For
End If
Next
Do While True
	If SwfWindow("进出口管理").SwfWindow("确认信息").Exist(1) Then
		Exit Do
	End If
Loop
Wait 1
SwfWindow("进出口管理").SwfWindow("确认信息").SwfObject("Yes").Click

If SwfWindow("进出口管理").SwfWindow("提示信息").Exist(1) Then
	SwfWindow("进出口管理").SwfWindow("提示信息").SwfObject("OK").Click
	If Not SwfWindow("进出口管理").SwfWindow("提示信息").Exist(1) Then
		WriteLogs("删除进出口返回====成功！")	
	Else
		WriteLogs("删除进出口返回====失败！")
	End If
Else
	WriteLogs("删除进出口返回====失败！")
End If

reporter.ReportEvent micPass,"Delete","修改成功！"
datatable.LocalSheet.AddParameter "删除结果"," "
datatable.GetSheet("进出口管理").SetCurrentRow(1)
datatable.Value("删除结果","进出口管理")="成功"
WriteLogs("数据表修改成功")
Wait 2
SwfWindow("进出口管理").Close()

WriteLogs("==================进出口管理模块结束===================")
wait 1	