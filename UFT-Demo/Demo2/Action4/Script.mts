
SwfWindow("视频识别出入口管理系统").SwfObject("设备管理").Click @@ hightlight id_;_132046_;_script infofile_;_ZIP::ssf7.xml_;_
wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu2").Click @@ hightlight id_;_197578_;_script infofile_;_ZIP::ssf8.xml_;_

Do While true
	if(SwfWindow("停车场通道管理").Exist(1)) then
		Exit do
	end if
loop
wait 1
SwfWindow("停车场通道管理").SwfObject("添加(A)").Click @@ hightlight id_;_4259910_;_script infofile_;_ZIP::ssf9.xml_;_


Do While true
	if(SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").Exist(1)) then
		Exit do
	end if
loop
wait 1

'=========================================================================
'ss=datatable.GetSheet("通道管理").GetParameter("区域信息").ValueByRow(1)
'msgbox ss
'arr=split(ss,"/")
'For i=0 to ubound(arr)
'msgbox arr(i)
'Next
'==========================================================================
Dim x,y,nameAndLocationArr

Dim parkingLotName
nameAndLocation=datatable.GetSheet("通道管理").GetParameter("区域信息").ValueByRow(1)

nameAndLocationArr=GetNameAndLocation(nameAndLocation)

parkingLotName=nameAndLocationArr(0)
x=nameAndLocationArr(1)
y=nameAndLocationArr(2)

SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbParkingLotName").Click
wait 1
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject").Click x,y @@ hightlight id_;_3212334_;_script infofile_;_ZIP::ssf11.xml_;_

wait 1
Dim channelName
channelName=datatable.GetSheet("通道管理").GetParameter("通道名称").ValueByRow(1)
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfEdit("SwfEdit_3").Set channelName
wait 1

incount=datatable.GetSheet("通道管理").GetParameter("进场通道数").ValueByRow(1)
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("spinInCount").DblClick 5,5,micLeftBtn
wait 1
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("spinInCount").Type micDel 
wait 1
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("spinInCount").Type incount


outCount = datatable.GetSheet("通道管理").GetParameter("出场通道数").ValueByRow(1)
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("spinOutCount").DblClick 5,5,micLeftBtn
wait 1
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("spinOutCount").Type micDel 
wait 1
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("spinOutCount").Type outCount


Dim mStation
nameAndLocation=datatable.GetSheet("通道管理").GetParameter("管理电脑").ValueByRow(1)

nameAndLocationArr=GetNameAndLocation(nameAndLocation)

mStation=nameAndLocationArr(0)
x=nameAndLocationArr(1)
y=nameAndLocationArr(2)

SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbMStation").Click @@ hightlight id_;_1639588_;_script infofile_;_ZIP::ssf12.xml_;_
wait 1
SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject").Click x,y @@ hightlight id_;_6882366_;_script infofile_;_ZIP::ssf13.xml_;_
wait 1

Dim chargeRule
nameAndLocation=datatable.GetSheet("通道管理").GetParameter("收费规则").ValueByRow(1)
If  nameAndLocation<>"" Then
	nameAndLocationArr=GetNameAndLocation(nameAndLocation)
	chargeRule=nameAndLocationArr(0)
	x=nameAndLocationArr(1)
	y=nameAndLocationArr(2)
	SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("cmbChargeRule").Click @@ hightlight id_;_1705332_;_script infofile_;_ZIP::ssf14.xml_;_
	wait 1
	SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("SwfWindow").SwfObject("SwfObject").Click x,y @@ hightlight id_;_4589102_;_script infofile_;_ZIP::ssf15.xml_;_
	wait 1
End If

SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfObject("保存").Click
 @@ hightlight id_;_4588598_;_script infofile_;_ZIP::ssf16.xml_;_
do While true
	if(SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("提示信息").Exist(1)) then
		wait 1
		SwfWindow("停车场通道管理").SwfWindow("保存停车场通道").SwfWindow("提示信息").SwfObject("OK").Click @@ hightlight id_;_4589700_;_script infofile_;_ZIP::ssf15.xml_;_
		wait 1
		Exit do
	end if
loop
wait 1

passFlag=false
tableRowCount=SwfWindow("停车场通道管理").SwfTable("gridControl1").RowCount
For i=0 to tableRowCount -1
	tempChannelName=SwfWindow("停车场通道管理").SwfTable("gridControl1").GetCellData(i,0)
	If  tempChannelName=channelName Then
			passFlag=true
			Exit for
	End If
Next

SwfWindow("停车场通道管理").Close
wait 1

If passFlag Then
	reporter.ReportEvent micPass,"Add","添加成功！"
	datatable.LocalSheet.AddParameter "结果"," "
	datatable.GetSheet("通道管理").SetCurrentRow(1)
	datatable.Value("结果","通道管理")="成功"
else
	reporter.ReportEvent  micFail ,"Add","添加失败！"
	datatable.LocalSheet.AddParameter "结果"," "
	datatable.GetSheet("通道管理").SetCurrentRow(1)
	datatable.Value("结果","通道管理")="失败"
End If

datatable.Export("F:\通道管理.xls")

 
'===========================================================================================================
Function GetNameAndLocation(strParamName)
	Dim x,y,paraName
	myArray=split(strParamName,"/")
	paraName=myArray(0)
	locationArr=split(myArray(1),",")
	x=locationArr(0)
	y=locationArr(1)
	NameAndLocationArr=array(paraName,x,y)
    GetNameAndLocation=NameAndLocationArr
End Function
'===========================================================================================================

