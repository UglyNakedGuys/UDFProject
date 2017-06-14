WriteLogs("初始化登录操作！")
SystemUtil.CloseProcessByName("ParkingVideo_Login.exe")
Wait 1
Dim Address
Address=Datatable.GetSheet("Global").GetParameter("Address").ValueByRow(1)
Address=Address&"\PakingVideo_Login.exe"
If  SwfWindow("登录界面").Exist(2) Then
 SwfWindow("登录界面").Close
End If
wait 1
systemutil.Run(Address)

'==================================================================================
'电脑终端注册
Dim CNo
CNo=Datatable.GetSheet("Global").GetParameter("CNo").ValueByRow(1)

If  SwfWindow("电脑终端注册").Exist(2) Then
	WriteLogs("电脑终端注册！")
	SwfWindow("电脑终端注册").Activate
'注册参数控件单独处理
	SwfWindow("电脑终端注册").SwfObject("txtStaionNo").DblClick 5,5,micLeftBtn
	SwfWindow("电脑终端注册").SwfObject("txtStaionNo").Type micDel
	wait 1
	SwfWindow("电脑终端注册").SwfEdit("SwfEdit").Type CNo
	wait 1
	SwfWindow("电脑终端注册").SwfObject("注册").Click @@ hightlight id_;_197996_;_script infofile_;_ZIP::ssf12.xml_;_
	wait 1
	SwfWindow("电脑终端注册").SwfWindow("提示信息").SwfObject("OK").Click
End If
'===================================================================================
'系统管理员权限登录 @@ hightlight id_;_460716_;_script infofile_;_ZIP::ssf13.xml_;_
wait 2
SwfWindow("登录界面").SwfEdit("SwfEdit").Set "admin"
SwfWindow("登录界面").SwfEdit("SwfEdit_2").SetSecure "5924e4b99935029c317c8fdbcdda0b6b"
wait 1

SwfWindow("登录界面").SwfObject("登录").Click

'====================================================================================================================
do While true
	if(SwfWindow("视频识别出入口管理系统").Exist(1)) then
		WriteLogs("管理员登录成功！")
		Exit do
	end if
loop
WriteLogs("===================区域管理模块开始====================")
WriteLogs("前置初始化操作！")
SwfWindow("视频识别出入口管理系统").SwfObject("设备管理").Click
wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu1").Click	
wait 1

do While true
	if(SwfWindow("区域管理").Exist(1)) then
		Exit do
	end if
loop
SwfWindow("区域管理").SwfObject("添加(A)").Click
'==================================================================================================================
'区域信息添加
WriteLogs("===================区域信息添加开始====================")

do While true
	if(SwfWindow("区域管理").SwfWindow("区域设置").Exist(1)) then
		Exit do
	end if
loop
WriteLogs("区域添加参数赋值开始！")
'区域名称赋值
Dim parkingLotName
parkingLotName=datatable.GetSheet("区域管理").GetParameter("区域名称").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit").Set parkingLotName
wait 1
'区域总车位数赋值
Dim totalNum
SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("spinEditTotalStopNumber").DblClick 5,5,micLeftBtn
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_2").Type  micDel   
totalNum=datatable.GetSheet("区域管理").GetParameter("总车位数").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_2").Type  totalNum  
wait 1
'区域已停车位数赋值
Dim stopNum
SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("spinEditStopCarNumber").DblClick 5,5,micLeftBtn
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_3").Type  micDel   
stopNum=datatable.GetSheet("区域管理").GetParameter("已停车位数").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_3").Type  stopNum  
wait 1
'区域是否内场赋值
Dim isInternl
isInternl=datatable.GetSheet("区域管理").GetParameter("是否内场").ValueByRow(1)
Select Case isInternl
			Case "是"  
			SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("radGpIsInternl").Click 111,14
'关联外场停车场名称
'			Dim ParkingLot
'			ParkingLot=Datatable.GetSheet("区域管理").GetParameter("关联停车场").ValueByRow(1)

			SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("radGpIsInternl").Click 105,17 @@ hightlight id_;_2295678_;_script infofile_;_ZIP::ssf22.xml_;_
			SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("cmbParkingLot").Click 131,11 @@ hightlight id_;_4655292_;_script infofile_;_ZIP::ssf23.xml_;_
			SwfWindow("SwfWindow").SwfObject("SwfObject").Click 175,9
			wait 1
			Case "否"   SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("radGpIsInternl").Click 13,12
End Select
WriteLogs("区域添加参数赋值结束！")

SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("保存").Click

passFlag=false
tableRowCount=SwfWindow("区域管理").SwfTable("gridControl1").RowCount
For i=0 to tableRowCount -1
	tempParkLotName=SwfWindow("区域管理").SwfTable("gridControl1").GetCellData(i,0)
	If  tempParkLotName=parkingLotName Then
		tempTotalNum=SwfWindow("区域管理").SwfTable("gridControl1").GetCellData(i,1)
		tempStopNum = SwfWindow("区域管理").SwfTable("gridControl1").GetCellData(i,2)
		If  tempTotalNum=  totalNum and  tempStopNum=stopNum Then
			WriteLogs("区域添加返回====成功！")
			passFlag=true
'			SwfWindow("区域管理").Close
			Exit for
		End If
	End If
Next
'数据表信息写入
If passFlag Then
	reporter.ReportEvent micPass,"Add","添加成功！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("区域管理").SetCurrentRow(1)
	datatable.Value("添加结果","区域管理")="成功"
	WriteLogs("数据表修改成功")
else
	reporter.ReportEvent  micFail ,"Add","添加失败！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("区域管理").SetCurrentRow(1)
	datatable.Value("添加结果","区域管理")="失败"
End If
Wait 2
'=========================================================================================================================
'区域信息修改
WriteLogs("===================区域信息修改开始====================")

Dim tempLotName
tempLotName=datatable.GetSheet("区域管理").GetParameter("区域名称").ValueByRow(1)
For Iterator = 0 To SwfWindow("区域管理").SwfTable("gridControl1").RowCount-1
If tempLotName=SwfWindow("区域管理").SwfTable("gridControl1").GetCellData(Iterator,0) Then
	SwfWindow("区域管理").SwfTable("gridControl1").ActivateCell Iterator,0
End If
Next
do While true
	if(SwfWindow("区域管理").SwfWindow("区域设置").Exist(1)) then
		Exit do
	end if
loop
WriteLogs("区域修改参数赋值开始！")
'区域名称赋值
parkingLotName=datatable.GetSheet("区域管理").GetParameter("修改区域名称").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit").Set parkingLotName
wait 1
'区域总车位数赋值
SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("spinEditTotalStopNumber").DblClick 5,5,micLeftBtn
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_2").Type  micDel   
totalNum=datatable.GetSheet("区域管理").GetParameter("修改总车位数").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_2").Type  totalNum  
wait 1
'区域已停车位数赋值
SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("spinEditStopCarNumber").DblClick 5,5,micLeftBtn
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_3").Type  micDel   
stopNum=datatable.GetSheet("区域管理").GetParameter("修改已停车位数").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_3").Type  stopNum  
wait 1
'区域是否内场赋值
isInternl=datatable.GetSheet("区域管理").GetParameter("修改是否内场").ValueByRow(1)
Select Case isInternl
			Case "是"  
			SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("radGpIsInternl").Click 111,14
'关联外场停车场名称
'			Dim ParkingLot
'			ParkingLot=Datatable.GetSheet("区域管理").GetParameter("关联停车场").ValueByRow(1)

			SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("radGpIsInternl").Click 105,17 @@ hightlight id_;_2295678_;_script infofile_;_ZIP::ssf22.xml_;_
			SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("cmbParkingLot").Click 131,11 @@ hightlight id_;_4655292_;_script infofile_;_ZIP::ssf23.xml_;_
			SwfWindow("SwfWindow").SwfObject("SwfObject").Click 175,9
			wait 1
			Case "否"   SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("radGpIsInternl").Click 13,12
End Select
WriteLogs("区域修改参数赋值结束！")

SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("保存").Click

passFlag=false
tableRowCount=SwfWindow("区域管理").SwfTable("gridControl1").RowCount
For i=0 to tableRowCount -1
	tempParkLotName=SwfWindow("区域管理").SwfTable("gridControl1").GetCellData(i,0)
	If  tempParkLotName=parkingLotName Then
		tempTotalNum=SwfWindow("区域管理").SwfTable("gridControl1").GetCellData(i,1)
		tempStopNum = SwfWindow("区域管理").SwfTable("gridControl1").GetCellData(i,2)
		If  tempTotalNum=  totalNum and  tempStopNum=stopNum Then
			WriteLogs("区域修改返回====成功！")
			passFlag=true
'			SwfWindow("区域管理").Close
			Exit for
		End If
	End If
Next
'数据表信息写入
If passFlag Then
	reporter.ReportEvent micPass,"Edit","修改成功！"
	datatable.LocalSheet.AddParameter "修改结果"," "
	datatable.GetSheet("区域管理").SetCurrentRow(1)
	datatable.Value("修改结果","区域管理")="成功"
	WriteLogs("数据表修改成功")
else
	reporter.ReportEvent  micFail ,"Edit","修改失败！"
	datatable.LocalSheet.AddParameter "修改结果"," "
	datatable.GetSheet("区域管理").SetCurrentRow(1)
	datatable.Value("修改结果","区域管理")="失败"
End If
Wait 2
'==============================================================================================================
WriteLogs("===================区域信息删除开始====================")
'区域信息删除
Dim deleteLotName
deleteLotName=datatable.GetSheet("区域管理").GetParameter("修改区域名称").ValueByRow(1)

For Iterator = 0 To SwfWindow("区域管理").SwfTable("gridControl1").RowCount-1
If deleteLotName=SwfWindow("区域管理").SwfTable("gridControl1").GetCellData(Iterator,0) Then
	SwfWindow("区域管理").SwfTable("gridControl1").SelectCell Iterator,0
	SwfWindow("区域管理").SwfObject("删除(D)").Click
End If
Next
Do While True
	If SwfWindow("区域管理").SwfWindow("确认信息").Exist(1) Then
		Exit Do
	End If
Loop
Wait 1
SwfWindow("区域管理").SwfWindow("确认信息").SwfObject("Yes").Click

If SwfWindow("区域管理").SwfWindow("提示信息").Exist(1) Then
	WriteLogs("删除区域返回====成功！")	
	Wait 1
	SwfWindow("区域管理").SwfWindow("提示信息").SwfObject("OK").Click
Else
	WriteLogs("删除区域返回====失败！")
End If

reporter.ReportEvent micPass,"Delete","修改成功！"
datatable.LocalSheet.AddParameter "删除结果"," "
datatable.GetSheet("区域管理").SetCurrentRow(1)
datatable.Value("删除结果","区域管理")="成功"
WriteLogs("数据表修改成功")
Wait 2
SwfWindow("区域管理").Close()

WriteLogs("===================区域管理模块结束====================")
wait 1	
 @@ hightlight id_;_1903680_;_script infofile_;_ZIP::ssf35.xml_;_