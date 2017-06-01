If  SwfWindow("登录界面").Exist(2) Then
 SwfWindow("登录界面").Close
 WriteLogs("关闭程序初始化！")
End If
wait 1
systemutil.Run("C:\Users\Administrator\Desktop\兔巢acs3.1\软件201705181100-Debug\Debug\PakingVideo_Login.exe")
 WriteLogs("启动程序！")

'==================================================================================
'电脑终端注册
If  SwfWindow("电脑终端注册").Exist(5) Then
	SwfWindow("电脑终端注册").SwfEdit("SwfEdit").Set "176"
	wait 2
	SwfWindow("电脑终端注册").SwfObject("注册").Click 44,16 @@ hightlight id_;_197996_;_script infofile_;_ZIP::ssf12.xml_;_
	wait 2
	SwfWindow("电脑终端注册").SwfWindow("提示信息").SwfObject("OK").Click 55,8 @@ hightlight id_;_460716_;_script infofile_;_ZIP::ssf13.xml_;_
	 WriteLogs("电脑注册成功！")
End If
'===================================================================================
 @@ hightlight id_;_460716_;_script infofile_;_ZIP::ssf13.xml_;_
wait	2
SwfWindow("登录界面").SwfEdit("SwfEdit").Set "admin"
SwfWindow("登录界面").SwfEdit("SwfEdit_2").SetSecure "5924e4b99935029c317c8fdbcdda0b6b"
wait 1

SwfWindow("登录界面").SwfObject("登录").Click
 WriteLogs("用户admin登录成功！")

'=========================================================================================
'首次登录配置画面处理
If  SwfWindow("配置向导").Exist(15)Then
	SwfWindow("配置向导").Close
	 WriteLogs("关闭配置画面！")
	wait 1
End If
'====================================================================================================================

do While true
	if(SwfWindow("视频识别出入口管理系统").Exist(1)) then
		Exit do
	end if
loop

SwfWindow("视频识别出入口管理系统").SwfObject("设备管理").Click
 WriteLogs("设备管理模块入口！")
 WriteLogs("-------------------------------------------------------")
wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu1").Click	
wait 1

do While true
	if(SwfWindow("区域管理").Exist(1)) then
		Exit do
	end if
loop
SwfWindow("区域管理").SwfObject("添加(A)").Click

do While true
	if(SwfWindow("区域管理").SwfWindow("区域设置").Exist(1)) then
		Exit do
	end if
loop

Dim parkingLotName
parkingLotName=datatable.GetSheet("Action1").GetParameter("区域名称").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_3").Set parkingLotName
WriteLogs("区域名称参数设置！")
wait 3
Dim totalNum
For i=0 to 5
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_2").Type  micDel   
Next
totalNum=datatable.GetSheet("Action1").GetParameter("总车位数").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit_2").Type  totalNum  
WriteLogs("区域总车位数参数设置！")
wait 3
Dim stopNum
For i=0 to 5
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit").Type  micDel   
Next
stopNum=datatable.GetSheet("Action1").GetParameter("已停车位数").ValueByRow(1)
SwfWindow("区域管理").SwfWindow("区域设置").SwfEdit("SwfEdit").Type  stopNum  
WriteLogs("区域已停车位数参数设置！")
wait 3

Dim isInternl
isInternl=datatable.GetSheet("Action1").GetParameter("是否内场").ValueByRow(1)
Select Case isInternl
			Case "是"  SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("radGpIsInternl").Click 113,14
			Case "否"   SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("radGpIsInternl").Click 19,17
			WriteLogs("区域是否内场参数设置！")
End Select

SwfWindow("区域管理").SwfWindow("区域设置").SwfObject("保存").Click
WriteLogs("区域配置成功！")

passFlag=false
tableRowCount=SwfWindow("区域管理").SwfTable("gridControl1").RowCount
For i=0 to tableRowCount -1
	tempParkLotName=SwfWindow("区域管理").SwfTable("gridControl1").GetCellData(i,0)
	If  tempParkLotName=parkingLotName Then
		tempTotalNum=SwfWindow("区域管理").SwfTable("gridControl1").GetCellData(i,1)
		tempStopNum = SwfWindow("区域管理").SwfTable("gridControl1").GetCellData(i,2)
		If  tempTotalNum=  totalNum and  tempStopNum=stopNum Then
			passFlag=true
			SwfWindow("区域管理").Close
			WriteLogs("设备管理模块出口！")
			WriteLogs("-------------------------------------------------------")
			Exit for
		End If
	End If
Next

If passFlag Then
	reporter.ReportEvent micPass,"Add","添加成功！"
	datatable.LocalSheet.AddParameter "结果"," "
	datatable.GetSheet("Action1").SetCurrentRow(1)
	datatable.Value("结果","Action1")="成功"
else
	reporter.ReportEvent  micFail ,"Add","添加失败！"
	datatable.LocalSheet.AddParameter "结果"," "
	datatable.GetSheet("Action1").SetCurrentRow(1)
	datatable.Value("结果","Action1")="失败"
End If

datatable.Export("E:\Jangboer201705\UFT-Demo\Excel\区域管理添加.xls")


wait 1
RunAction "临时时长收费", oneIteration

wait 1
RunAction "长期按期收费", oneIteration

wait 1

RunAction "通道管理", oneIteration
wait 1

RunAction "摄像机管理", oneIteration

wait 1
RunAction "进出口管理", oneIteration

wait 1
