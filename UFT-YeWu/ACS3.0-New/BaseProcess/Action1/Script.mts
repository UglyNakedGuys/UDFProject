WriteLogs("初始化登录操作！")
SystemUtil.CloseProcessByName("ParkingVideo_Login.exe")
Wait 1
If  SwfWindow("登录界面").Exist(2) Then
 SwfWindow("登录界面").Close
End If
wait 1
systemutil.Run("C:\Users\Administrator\Desktop\兔巢acs3.1\软件201705181100-Debug\Debug\PakingVideo_Login.exe")

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
WriteLogs("=================区域信息添加开始==================")
do While true
	if(SwfWindow("区域管理").SwfWindow("区域设置").Exist(1)) then
		Exit do
	end if
loop
WriteLogs("区域添加参数赋值开始！")
'区域名称赋值
Dim parkingLotName
parkingLotName=datatable.GetSheet("Action1").GetParameter("区域名称").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit").Set parkingLotName
wait 1
'区域总车位数赋值
Dim totalNum
SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("spinEditTotalStopNumber").DblClick 5,5,micLeftBtn
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_2").Type  micDel   
totalNum=datatable.GetSheet("Action1").GetParameter("总车位数").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_2").Type  totalNum  
wait 1
'区域已停车位数赋值
Dim stopNum
SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("spinEditStopCarNumber").DblClick 5,5,micLeftBtn
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_3").Type  micDel   
stopNum=datatable.GetSheet("Action1").GetParameter("已停车位数").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_3").Type  stopNum  
wait 1
'区域是否内场赋值
Dim isInternl
isInternl=datatable.GetSheet("Action1").GetParameter("是否内场").ValueByRow(1)
Select Case isInternl
			Case "是"  
			SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("radGpIsInternl").Click 111,14
'关联外场停车场名称
'			Dim ParkingLot
'			ParkingLot=Datatable.GetSheet("Action1").GetParameter("关联停车场").ValueByRow(1)

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
			WriteLogs("区域添加返回===========成功！")
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
	datatable.GetSheet("Action1").SetCurrentRow(1)
	datatable.Value("添加结果","Action1")="成功"
	WriteLogs("数据表导出成功")
else
	reporter.ReportEvent  micFail ,"Add","添加失败！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("Action1").SetCurrentRow(1)
	datatable.Value("添加结果","Action1")="失败"
End If
'区域信息修改
WriteLogs("=================区域信息修改开始==================")
Dim tempLotName
tempLotName=datatable.GetSheet("Action1").GetParameter("区域名称").ValueByRow(1)
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
parkingLotName=datatable.GetSheet("Action1").GetParameter("修改区域名称").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit").Set parkingLotName
wait 1
'区域总车位数赋值
SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("spinEditTotalStopNumber").DblClick 5,5,micLeftBtn
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_2").Type  micDel   
totalNum=datatable.GetSheet("Action1").GetParameter("修改总车位数").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_2").Type  totalNum  
wait 1
'区域已停车位数赋值
SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("spinEditStopCarNumber").DblClick 5,5,micLeftBtn
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_3").Type  micDel   
stopNum=datatable.GetSheet("Action1").GetParameter("修改已停车位数").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_3").Type  stopNum  
wait 1
'区域是否内场赋值
isInternl=datatable.GetSheet("Action1").GetParameter("修改是否内场").ValueByRow(1)
Select Case isInternl
			Case "是"  
			SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("radGpIsInternl").Click 111,14
'关联外场停车场名称
'			Dim ParkingLot
'			ParkingLot=Datatable.GetSheet("Action1").GetParameter("关联停车场").ValueByRow(1)

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
			WriteLogs("区域修改返回===========成功！")
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
	datatable.GetSheet("Action1").SetCurrentRow(1)
	datatable.Value("修改结果","Action1")="成功"
	WriteLogs("数据表导出成功")
else
	reporter.ReportEvent  micFail ,"Edit","修改失败！"
	datatable.LocalSheet.AddParameter "修改结果"," "
	datatable.GetSheet("Action1").SetCurrentRow(1)
	datatable.Value("修改结果","Action1")="失败"
End If

datatable.Export("E:\Jangboer201705\UFT-YeWu\ACS3.0-New\BaseProcess\Excel\区域管理添加.xls")


'wait 1
'RunAction "临时时长收费", oneIteration
'
'wait 1
'RunAction "长期按期收费", oneIteration
'
'wait 1
'
'RunAction "通道管理", oneIteration
'wait 1
'
'RunAction "摄像机管理", oneIteration
'
'wait 1
'RunAction "进出口管理", oneIteration

wait 1	
 @@ hightlight id_;_1903680_;_script infofile_;_ZIP::ssf35.xml_;_