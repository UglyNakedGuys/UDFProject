SwfWindow("视频识别出入口管理系统").SwfObject("设备管理").Click @@ hightlight id_;_132046_;_script infofile_;_ZIP::ssf1.xml_;_
wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu5").Click @@ hightlight id_;_197956_;_script infofile_;_ZIP::ssf2.xml_;_
Do While true
	if(SwfWindow("摄像机管理").Exist(1)) then
		Exit do
	end if
loop
wait 1

Call AddCamere(1)
wait 1
Call AddCamere(2)
wait 1
SwfWindow("摄像机管理").Close
wait 1
datatable.Export("F:\摄像机管理.xls")


Sub AddCamere(cellNumber)

SwfWindow("摄像机管理").SwfObject("添加(A)").Click

Do While true
	if(SwfWindow("摄像机管理").SwfWindow("修改识别器信息").Exist(1)) then
		Exit do
	end if
loop
wait 1

Dim cameraName
cameraName=datatable.GetSheet("摄像机管理").GetParameter("识别器名称").ValueByRow(cellNumber)
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfEdit("SwfEdit_5").Set cameraName
wait 1

camaraType=datatable.GetSheet("摄像机管理").GetParameter("相机类型").ValueByRow(cellNumber)
If  camaraType<>"" Then
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("cmbCameraType").Click @@ hightlight id_;_2557124_;_script infofile_;_ZIP::ssf3.xml_;_
	wait 1
	If camaraType="RF-LPR20D" Then
		SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 44,9
	else
		SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 49,29
	End If
End If

Dim cameraIp
cameraIP=datatable.GetSheet("摄像机管理").GetParameter("识别IP").ValueByRow(cellNumber)
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfEdit("SwfEdit_4").Set cameraIp
wait 1

Dim cameraPort
cameraPort=datatable.GetSheet("摄像机管理").GetParameter("识别端口").ValueByRow(cellNumber)
If  cameraPort<>"" Then
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("txtCameraPort").DblClick 5,5,micLeftBtn
	wait 1
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("txtCameraPort").Type micDel 
	wait 1
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("txtCameraPort").Type cameraPort
	wait 1
End If

Dim userName
userName=datatable.GetSheet("摄像机管理").GetParameter("用户名称").ValueByRow(cellNumber)
If  userName<>"" Then
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfEdit("SwfEdit_2").Set userName
	wait 1
End If

Dim password
password=datatable.GetSheet("摄像机管理").GetParameter("用户密码").ValueByRow(cellNumber)
If password<>"" Then
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("txtPassword").DblClick 5,5,micLeftBtn
	wait 1
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("txtPassword").Type micDel 
	wait 1
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("txtPassword").Type password
	wait 1
End If

Dim workType
workType=datatable.GetSheet("摄像机管理").GetParameter("工作模式").ValueByRow(cellNumber)
If workType<>"" Then
	If workType="有人值守模式" Then
		SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("rbgCameraType").Click 49,10
	else
		SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("rbgCameraType").Click 186,10 @@ hightlight id_;_6554708_;_script infofile_;_ZIP::ssf8.xml_;_
	End If
	wait 1
End If

SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("保存").Click

do While true
	if(SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfWindow("提示信息").Exist(1)) then
		wait 1
		SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfWindow("提示信息").SwfObject("OK").Click @@ hightlight id_;_4589700_;_script infofile_;_ZIP::ssf15.xml_;_
		wait 1
		Exit do
	end if
loop
wait 1

passFlag=false
tableRowCount=SwfWindow("摄像机管理").SwfTable("gridControl1").RowCount
For i=0 to tableRowCount -1
	tempCameraName=SwfWindow("摄像机管理").SwfTable("gridControl1").GetCellData(i,0)
	If  tempCameraName=cameraName Then
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
	datatable.GetSheet("摄像机管理").SetCurrentRow(cellNumber)
	datatable.Value("结果","摄像机管理")="成功"
else
	reporter.ReportEvent  micFail ,"Add",cameraName&"添加失败！"
'	datatable.LocalSheet.AddParameter "结果"," "
	datatable.GetSheet("摄像机管理").SetCurrentRow(cellNumber)
	datatable.Value("结果","摄像机管理")="失败"
End If

End Sub



