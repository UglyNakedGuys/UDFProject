SwfWindow("视频识别出入口管理系统").SwfObject("设备管理").Click @@ hightlight id_;_132046_;_script infofile_;_ZIP::ssf1.xml_;_
wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu3").Click @@ hightlight id_;_1313064_;_script infofile_;_ZIP::ssf2.xml_;_

Do While true
	if(SwfWindow("进出口管理").Exist(1)) then
		Exit do
	end if
loop
wait 1

call AddCamera("进",1)
wait 2
call AddCamera("出",2)
wait 2
SwfWindow("进出口管理").Close 

datatable.Export("F:\进出口管理.xls")

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

'===========================================================================
Sub AddCamera(inType,cellNumber)

Dim parkingChannelName,cellNum,findFlag
findFlag=false
' 处理进口，添加摄像机=====
parkingChannelName=datatable.GetSheet("进出口管理").GetParameter("停车场通道").ValueByRow(cellNumber)

tableRowCount=SwfWindow("进出口管理").SwfTable("gridControl1").RowCount
For i=0 to tableRowCount -1
	'gridControl1实际上有11列，代码隐藏了几列
	tempParkingChannelName=SwfWindow("进出口管理").SwfTable("gridControl1").GetCellData(i,7)
	If  tempParkingChannelName=parkingChannelName Then
			If SwfWindow("进出口管理").SwfTable("gridControl1").GetCellData(i,1)=inType Then
					cellNum=i
					findFlag=true
					Exit for
			End If
	End If
Next

wait 1
If findFlag Then
	SwfWindow("进出口管理").SwfTable("gridControl1").SetView "" @@ hightlight id_;_3409660_;_script infofile_;_ZIP::ssf3.xml_;_
	wait 1
	SwfWindow("进出口管理").SwfTable("gridControl1").SelectCell cellNum,"进出状态" @@ hightlight id_;_3409660_;_script infofile_;_ZIP::ssf4.xml_;_
	wait 1
	SwfWindow("进出口管理").SwfObject("编辑(E)").Click @@ hightlight id_;_1509850_;_script infofile_;_ZIP::ssf5.xml_;_
End If

' 点击编辑后进行添加摄像机
wait 1
Do While true
	if(SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").Exist(1)) then
		Exit do
	end if
loop
wait 1

Dim channelName
channelName=datatable.GetSheet("进出口管理").GetParameter("通道名称").ValueByRow(cellNumber)
If  channelName<>"" Then
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfEdit("SwfEdit").Set channelName
	wait 1
End If


Dim x,y,nameAndLocationArr

' 收费规则
nameAndLocation=datatable.GetSheet("进出口管理").GetParameter("收费规则").ValueByRow(cellNumber)
If  nameAndLocation <>"" Then
	Dim chargeRule
	nameAndLocationArr=GetNameAndLocation(nameAndLocation)
	chargeRule=nameAndLocationArr(0)
	x=nameAndLocationArr(1)
	y=nameAndLocationArr(2)

	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbChargeRule").Click @@ hightlight id_;_529902_;_script infofile_;_ZIP::ssf15.xml_;_
	wait 1
	SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject").Click x,y @@ hightlight id_;_661146_;_script infofile_;_ZIP::ssf16.xml_;_
	wait 1

End If


'选择主相机
Dim cameraName

nameAndLocation=datatable.GetSheet("进出口管理").GetParameter("主相机").ValueByRow(cellNumber)

nameAndLocationArr=GetNameAndLocation(nameAndLocation)
cameraName=nameAndLocationArr(0)
x=nameAndLocationArr(1)
y=nameAndLocationArr(2)

SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("cmbCameraList").Click @@ hightlight id_;_464616_;_script infofile_;_ZIP::ssf6.xml_;_
wait 1
SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("SwfWindow").SwfObject("SwfObject").Click x,y @@ hightlight id_;_1052782_;_script infofile_;_ZIP::ssf7.xml_;_
wait 1

Dim tempMode
tempMode=datatable.GetSheet("进出口管理").GetParameter("临时触发模式").ValueByRow(cellNumber)
Select Case tempMode
	Case "卡片" SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("rdoTempTriggerMode").Click 20,12
	Case "车牌"  SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("rdoTempTriggerMode").Click 170,12
End Select
wait 1

Dim longMode
longMode=datatable.GetSheet("进出口管理").GetParameter("长期触发模式").ValueByRow(cellNumber)
Select Case longMode
		Case "卡片" SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("rdoTriggerMode").Click 22,14
		Case "车牌"  SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("rdoTriggerMode").Click 100,14
		Case "车牌或卡" SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("rdoTriggerMode").Click 180,14
		Case "车牌和卡"  SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("rdoTriggerMode").Click 253,14
End Select
wait 1

SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfObject("保存").Click @@ hightlight id_;_398974_;_script infofile_;_ZIP::ssf8.xml_;_

do While true
	if(SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("提示信息").Exist(1)) then
		wait 1
		SwfWindow("进出口管理").SwfWindow("保存进出口通道信息").SwfWindow("提示信息").SwfObject("OK").Click @@ hightlight id_;_4589700_;_script infofile_;_ZIP::ssf15.xml_;_
		wait 1
		Exit do
	end if
loop
wait 1
 @@ hightlight id_;_398974_;_script infofile_;_ZIP::ssf9.xml_;_

passFlag=false
tableRowCount=SwfWindow("进出口管理").SwfTable("gridControl1").RowCount
For i=0 to tableRowCount -1
	tempCameraName=SwfWindow("进出口管理").SwfTable("gridControl1").GetCellData(i,9)
	tempParkingChannelName = SwfWindow("进出口管理").SwfTable("gridControl1").GetCellData(i,7)  
	If ( tempCameraName=cameraName  and  tempParkingChannelName=parkingChannelName) Then
			passFlag=true
			Exit for
	End If
Next

'避免重复加列
If  cellNumber=1 Then
	datatable.LocalSheet.AddParameter "结果"," "
End If

If passFlag Then
	reporter.ReportEvent micPass,"Add",cameraName&"添加成功！"
'	datatable.LocalSheet.AddParameter "结果"," "
	datatable.GetSheet("进出口管理").SetCurrentRow(cellNumber)
	datatable.Value("结果","进出口管理")="成功"
else
	reporter.ReportEvent  micFail ,"Add",cameraName&"添加失败！"
'	datatable.LocalSheet.AddParameter "结果"," "
	datatable.GetSheet("进出口管理").SetCurrentRow(cellNumber)
	datatable.Value("结果","进出口管理")="失败"
End If

End Sub
'===========================================================================’



