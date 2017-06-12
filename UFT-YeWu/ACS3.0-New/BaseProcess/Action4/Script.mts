WriteLogs("===================相机管理模块开始====================")
WriteLogs("前置初始化操作！")
Do While True
If SwfWindow("视频识别出入口管理系统").Exist(1) Then
    Exit Do
End If
Loop 
SwfWindow("视频识别出入口管理系统").SwfObject("设备管理").Click @@ hightlight id_;_198172_;_script infofile_;_ZIP::ssf1.xml_;_
Wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu5").Click @@ hightlight id_;_591216_;_script infofile_;_ZIP::ssf2.xml_;_
Do	While True
If  SwfWindow("摄像机管理").Exist(1) Then
	Exit Do
End If
Loop
Wait 1
SwfWindow("摄像机管理").SwfObject("添加(A)").Click
'====================================================================================================
'摄像机信息添加
WriteLogs("==================摄像机信息添加开始===================")
Wait 1 @@ hightlight id_;_4392580_;_script infofile_;_ZIP::ssf3.xml_;_

Dim cameraName
cameraName=Datatable.GetSheet("相机管理").GetParameter("摄像机名称").ValueByRow(1)
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfEdit("SwfEdit").Set cameraName @@ hightlight id_;_3673316_;_script infofile_;_ZIP::ssf10.xml_;_
Wait 1
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("cmbCameraType").Click 114,15
Dim cameraType,x,y @@ hightlight id_;_4260094_;_script infofile_;_ZIP::ssf21.xml_;_
cameraType=Datatable.GetSheet("相机管理").GetParameter("摄像机类型").ValueByRow(1)
tempStr=GetLocalPos(cameraType)
cameraType=tempStr(0)
x=tempStr(1)
y=tempStr(2)
Select Case cameraType
Case "大华"
	SwfWindow("SwfWindow").SwfObject("SwfObject").Click 49,11
Case Else
	SwfWindow("SwfWindow").SwfObject("SwfObject").Click 49,28
End Select
Wait 1
Dim cameraIp @@ hightlight id_;_7274730_;_script infofile_;_ZIP::ssf12.xml_;_
cameraIp=Datatable.GetSheet("相机管理").GetParameter("摄像机地址").ValueByRow(1)
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfEdit("SwfEdit_2").Set cameraIp @@ hightlight id_;_7473312_;_script infofile_;_ZIP::ssf13.xml_;_
Wait 1
Dim mode
mode=Datatable.GetSheet("相机管理").GetParameter("摄像机工作模式").ValueByRow(1)
Select Case mode
Case "有人值守"
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("rbgCameraType").Click 12,12 @@ hightlight id_;_9110884_;_script infofile_;_ZIP::ssf22.xml_;_
Case Else
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("rbgCameraType").Click 141,9 @@ hightlight id_;_9110884_;_script infofile_;_ZIP::ssf23.xml_;_
End Select
Wait 1 @@ hightlight id_;_1115936_;_script infofile_;_ZIP::ssf16.xml_;_
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("预览").Click @@ hightlight id_;_4129248_;_script infofile_;_ZIP::ssf17.xml_;_
Wait 2
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("保存").Click @@ hightlight id_;_12977452_;_script infofile_;_ZIP::ssf18.xml_;_
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfWindow("提示信息").SwfObject("OK").Click @@ hightlight id_;_7212070_;_script infofile_;_ZIP::ssf19.xml_;_
passFlag=False
If Not SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfWindow("提示信息").Exist(1) Then
	passFlag=True
	WriteLogs("添加摄像机返回===成功！")
Else
	WriteLogs("添加摄像机返回===失败！")
End If
If passFlag Then
	reporter.ReportEvent micPass,"Add","添加成功！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("相机管理").SetCurrentRow(1)
	datatable.Value("添加结果","相机管理")="成功"
	WriteLogs("数据表修改成功")
Else
	reporter.ReportEvent micPass,"Add","添加失败！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("相机管理").SetCurrentRow(1)
	datatable.Value("添加结果","相机管理")="失败"
	WriteLogs("数据表修改成功")
End If
'=======================================================================================
 '摄像机信息修改
WriteLogs("==================摄像机信息修改开始===================")
Wait 1
Dim tempCameraName
tempCameraName=Datatable.GetSheet("相机管理").GetParameter("摄像机名称").ValueByRow(1)
For Iterator = 0 To SwfWindow("摄像机管理").SwfTable("gridControl1").RowCount-1
	If tempCameraName=SwfWindow("摄像机管理").SwfTable("gridControl1").GetCellData(Iterator,0) Then
		SwfWindow("摄像机管理").SwfTable("gridControl1").ActivateCell Iterator,0
	End If
Next
Do While True
	If SwfWindow("摄像机管理").SwfWindow("修改识别器信息").Exist(1) Then
		Exit Do
	End If
Loop
Wait 1

cameraName=Datatable.GetSheet("相机管理").GetParameter("修改摄像机名称").ValueByRow(1)
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfEdit("SwfEdit").Set cameraName @@ hightlight id_;_3673316_;_script infofile_;_ZIP::ssf10.xml_;_
Wait 1
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("cmbCameraType").Click 114,15 @@ hightlight id_;_4260094_;_script infofile_;_ZIP::ssf21.xml_;_
cameraType=Datatable.GetSheet("相机管理").GetParameter("修改摄像机类型").ValueByRow(1)
tempStr=GetLocalPos(cameraType)
cameraType=tempStr(0)
x=tempStr(1)
y=tempStr(2)
Select Case cameraType
Case "大华"
	SwfWindow("SwfWindow").SwfObject("SwfObject").Click 49,11
Case Else
	SwfWindow("SwfWindow").SwfObject("SwfObject").Click 49,28
End Select
Wait 1 @@ hightlight id_;_7274730_;_script infofile_;_ZIP::ssf12.xml_;_
cameraIp=Datatable.GetSheet("相机管理").GetParameter("修改摄像机地址").ValueByRow(1)
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfEdit("SwfEdit_2").Set cameraIp @@ hightlight id_;_7473312_;_script infofile_;_ZIP::ssf13.xml_;_
Wait 1
mode=Datatable.GetSheet("相机管理").GetParameter("修改摄像机工作模式").ValueByRow(1)
Select Case mode
Case "有人值守"
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("rbgCameraType").Click 12,12 @@ hightlight id_;_9110884_;_script infofile_;_ZIP::ssf22.xml_;_
Case Else
	SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("rbgCameraType").Click 141,9 @@ hightlight id_;_9110884_;_script infofile_;_ZIP::ssf23.xml_;_
End Select
Wait 1 @@ hightlight id_;_1115936_;_script infofile_;_ZIP::ssf16.xml_;_
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("预览").Click @@ hightlight id_;_4129248_;_script infofile_;_ZIP::ssf17.xml_;_
Wait 2
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfObject("保存").Click @@ hightlight id_;_12977452_;_script infofile_;_ZIP::ssf18.xml_;_
SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfWindow("提示信息").SwfObject("OK").Click @@ hightlight id_;_7212070_;_script infofile_;_ZIP::ssf19.xml_;_
passFlag=False
If Not SwfWindow("摄像机管理").SwfWindow("修改识别器信息").SwfWindow("提示信息").Exist(1) Then
	passFlag=True
	WriteLogs("摄像机修改返回====成功！")
Else
	WriteLogs("摄像机修改返回====失败！")	
End If
'数据表信息写入
If passFlag Then
	reporter.ReportEvent micPass,"Edit","修改成功！"
	datatable.LocalSheet.AddParameter "修改结果"," "
	datatable.GetSheet("相机管理").SetCurrentRow(1)
	datatable.Value("修改结果","相机管理")="成功"
	WriteLogs("数据表修改成功")
else
	reporter.ReportEvent  micFail ,"Edit","修改失败！"
	datatable.LocalSheet.AddParameter "修改结果"," "
	datatable.GetSheet("相机管理").SetCurrentRow(1)
	datatable.Value("修改结果","相机管理")="失败"
End If
Wait 2 @@ hightlight id_;_1901000_;_script infofile_;_ZIP::ssf30.xml_;_
'============================================================================================
WriteLogs("==================摄像机信息删除开始===================")
'摄像机信息删除
Dim deleteCameraName
deleteCameraName=datatable.GetSheet("相机管理").GetParameter("修改摄像机名称").ValueByRow(1)

For Iterator = 0 To SwfWindow("摄像机管理").SwfTable("gridControl1").RowCount-1
If deleteCameraName=SwfWindow("摄像机管理").SwfTable("gridControl1").GetCellData(Iterator,0) Then
	SwfWindow("摄像机管理").SwfTable("gridControl1").SelectCell Iterator,0
	SwfWindow("摄像机管理").SwfObject("删除(D)").Click
End If
Next
Do While True
	If SwfWindow("摄像机管理").SwfWindow("确认信息").Exist(1) Then
		Exit Do
	End If
Loop
Wait 1
SwfWindow("摄像机管理").SwfWindow("确认信息").SwfObject("Yes").Click

If SwfWindow("摄像机管理").SwfWindow("提示信息").Exist(1) Then
	SwfWindow("摄像机管理").SwfWindow("提示信息").SwfObject("OK").Click
	If Not SwfWindow("摄像机管理").SwfWindow("提示信息").Exist(1) Then
		WriteLogs("删除摄像机返回===成功！")	
	Else
		WriteLogs("删除摄像机返回====失败！")
	End If
Else
	WriteLogs("删除摄像机返回====失败！")
End If

reporter.ReportEvent micPass,"Delete","修改成功！"
datatable.LocalSheet.AddParameter "删除结果"," "
datatable.GetSheet("相机管理").SetCurrentRow(1)
datatable.Value("删除结果","相机管理")="成功"
WriteLogs("数据表修改成功")
Wait 2
SwfWindow("摄像机管理").Close()

WriteLogs("==================摄像机管理模块结束===================")
wait 1	
 datatable.Export("E:\Jangboer201705\UFT-YeWu\ACS3.0-New\BaseProcess\Excel\基础流程数据表.xls")
WriteLogs("数据表导出成功")
Wait 2
 @@ hightlight id_;_3277248_;_script infofile_;_ZIP::ssf28.xml_;_