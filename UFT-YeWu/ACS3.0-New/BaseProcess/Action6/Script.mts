WriteLogs("===================收费规则模块开始====================")
WriteLogs("前置初始化操作！")
Do While True
If SwfWindow("视频识别出入口管理系统").Exist(1) Then
    Exit Do
End If
Loop 
SwfWindow("视频识别出入口管理系统").SwfObject("系统设置").Click
Wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu4").Click
Do	While True
If  SwfWindow("收费规则管理").Exist(1) Then
	Exit Do
End If
Loop
Wait 1 @@ hightlight id_;_920786_;_script infofile_;_ZIP::ssf37.xml_;_
SwfWindow("收费规则管理").SwfObject("rdoChannelType").Click 106,12
Wait 1
SwfWindow("收费规则管理").SwfObject("添加(A)").Click
'======================================================================================================
'收费规则信息添加
WriteLogs("=================收费规则信息添加开始==================") @@ hightlight id_;_2032994_;_script infofile_;_ZIP::ssf4.xml_;_
Wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("cmbChargeType").Click
Wait 1
i=0
Do While True @@ hightlight id_;_655550_;_script infofile_;_ZIP::ssf55.xml_;_
	SwfWindow("SwfWindow").SwfObject("SwfObject").Click 91,12
	i=i+14
	If SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("cmbChargeType").GetROProperty("Text")="时间收费" Then
		Exit Do
	End If
Loop 
Wait 1
Dim RuleName
RuleName=Datatable.GetSheet("临时收费规则").GetParameter("规则名称").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit").Set RuleName @@ hightlight id_;_1050292_;_script infofile_;_ZIP::ssf7.xml_;_
Wait 1
Dim RuleType
RuleType=Datatable.GetSheet("临时收费规则").GetParameter("规则类型").ValueByRow(1)
Select Case RuleType
Case "分时"
   	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("rdoChargeType").Click 20,12
Case "时段"
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("rdoChargeType").Click 113,14
Case "时长"
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("rdoChargeType").Click 200,10
Case "按次" @@ hightlight id_;_1313902_;_script infofile_;_ZIP::ssf9.xml_;_
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("rdoChargeType").Click 296,12
End Select
Wait 1
Dim NoChargeTime @@ hightlight id_;_1313902_;_script infofile_;_ZIP::ssf11.xml_;_
NoChargeTime=Datatable.GetSheet("临时收费规则").GetParameter("不收费时间").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtStopTime").Click 
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_2").Type NoChargeTime @@ hightlight id_;_788208_;_script infofile_;_ZIP::ssf12.xml_;_
Wait 1
Dim TotalMoney
TotalMoney=Datatable.GetSheet("临时收费规则").GetParameter("总限额").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtAllDayCharge").Click
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_3").Type TotalMoney @@ hightlight id_;_461430_;_script infofile_;_ZIP::ssf13.xml_;_
Wait 1 @@ hightlight id_;_2819448_;_script infofile_;_ZIP::ssf16.xml_;_
Dim IsCharge
IsCharge=Datatable.GetSheet("临时收费规则").GetParameter("是否免费时间收费").ValueByRow(1)
If IsCharge="是" Then
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("chkFreeTimeCharge").Click
End If @@ hightlight id_;_855018_;_script infofile_;_ZIP::ssf45.xml_;_
Wait 1
Dim UnitTime
UnitTime=Datatable.GetSheet("临时收费规则").GetParameter("单位时间").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtEveryHour").Click
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_5").Type UnitTime @@ hightlight id_;_1180812_;_script infofile_;_ZIP::ssf17.xml_;_
Dim UnitMoney
UnitMoney=Datatable.GetSheet("临时收费规则").GetParameter("单位金额").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtChargeMoney").Click
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_6").Type UnitMoney @@ hightlight id_;_461414_;_script infofile_;_ZIP::ssf18.xml_;_
Wait 1
Dim MaximumCharge
MaximumCharge=Datatable.GetSheet("临时收费规则").GetParameter("最多收费").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtMaxCharge").Click
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_4").Type MaximumCharge
Wait 1 @@ hightlight id_;_2557960_;_script infofile_;_ZIP::ssf50.xml_;_
'==============================================================
'时段字符串切割
Dim StartTime
Dim EndTime
tempTime=Datatable.GetSheet("临时收费规则").GetParameter("时段").ValueByRow(1)
tempArray=Split(tempTime,"/")
tempStartTime=tempArray(0)
tempEndTime=tempArray(1)
StartTime=Split(tempStartTime,":")
EndTime=Split(tempEndTime,":")
'==============================================================
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtStartTime").DblClick 5,5,micLeftBtn @@ hightlight id_;_1312512_;_script infofile_;_ZIP::ssf52.xml_;_
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_7").Type StartTime(0) @@ hightlight id_;_2428058_;_script infofile_;_ZIP::ssf53.xml_;_
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_7").Type micRight @@ hightlight id_;_2428058_;_script infofile_;_ZIP::ssf53.xml_;_
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_7").Type StartTime(1) @@ hightlight id_;_2428058_;_script infofile_;_ZIP::ssf53.xml_;_
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_7").Type micRight
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_7").Type StartTime(2)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_7").Type micRight
Wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtStopTime").DblClick 5,5,micLeftBtn
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_8").Type EndTime(0)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_8").Type micRight
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_8").Type EndTime(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_8").Type micRight
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_8").Type EndTime(2)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_8").Type micRight
Wait 1
Dim Sub_Rule
Sub_Rule=Datatable.GetSheet("临时收费规则").GetParameter("子规则").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfComboBox("cmb_ChargeType").Select Sub_Rule
'i=0
'Do While True @@ hightlight id_;_3673498_;_script infofile_;_ZIP::ssf57.xml_;_
'	SwfWindow("SwfWindow").SwfObject("SwfObject").Click 91,12+i
' @@ hightlight id_;_3673498_;_script infofile_;_ZIP::ssf58.xml_;_
'	i=i+14
'	If SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("cmbChargeType").GetROProperty("Text")="时间收费" Then
'		Exit Do
'	End If
'Loop 
Wait 1

SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("添加").Click 
Wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("保存").Click @@ hightlight id_;_1641556_;_script infofile_;_ZIP::ssf20.xml_;_
passFlag=false
Wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("SwfWindow").SwfObject("OK").Click
If  SwfWindow("收费规则管理").SwfWindow("编辑收费规则").Exist(1) Then
	passFlag=true
	WriteLogs("添加收费规则返回=成功！")
Else
	WriteLogs("添加收费规则返回=失败！")
End If
If passFlag Then
	reporter.ReportEvent micPass,"Add","添加成功！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("临时收费规则").SetCurrentRow(1)
	datatable.Value("添加结果","临时收费规则")="成功"
	WriteLogs("数据表修改成功")
Else
	reporter.ReportEvent micPass,"Add","添加失败！"
	datatable.LocalSheet.AddParameter "添加结果"," "
	datatable.GetSheet("临时收费规则").SetCurrentRow(1)
	datatable.Value("添加结果","临时收费规则")="失败"
	WriteLogs("数据表修改成功")
End If
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").Close()
'========================================================================================================
 '操作员信息修改
WriteLogs("=================收费规则信息修改开始==================")
Wait 1

'SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("SwfWindow").SwfObject("OK").Click 40,8 @@ hightlight id_;_2884984_;_script infofile_;_ZIP::ssf21.xml_;_
'SwfWindow("收费规则管理").SwfWindow("编辑收费规则").Close
'SwfWindow("收费规则管理").SwfTable("gridControlChargingRules").SetView "" @@ hightlight id_;_920100_;_script infofile_;_ZIP::ssf22.xml_;_
'SwfWindow("收费规则管理").SwfTable("gridControlChargingRules").SelectCell 5,"收费规则名称" @@ hightlight id_;_920100_;_script infofile_;_ZIP::ssf23.xml_;_
'SwfWindow("收费规则管理").SwfObject("编辑").Click 57,9 @@ hightlight id_;_1902792_;_script infofile_;_ZIP::ssf24.xml_;_
'SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfTable("gridControl1").SetView "" @@ hightlight id_;_1181886_;_script infofile_;_ZIP::ssf25.xml_;_
'SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfTable("gridControl1").SelectCell 0,"金额" @@ hightlight id_;_1181886_;_script infofile_;_ZIP::ssf26.xml_;_
'SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit").SetSelection 0,5 @@ hightlight id_;_1379438_;_script infofile_;_ZIP::ssf27.xml_;_
'SwfWindow("收费规则管理").SwfWindow("编辑收费规则").Close
'SwfWindow("收费规则管理").SwfTable("gridControlChargingRules").SelectCell 5,"收费规则名称" @@ hightlight id_;_920100_;_script infofile_;_ZIP::ssf28.xml_;_
'SwfWindow("收费规则管理").SwfObject("删除(D)").Click 34,16 @@ hightlight id_;_1574184_;_script infofile_;_ZIP::ssf29.xml_;_
'SwfWindow("收费规则管理").SwfObject("rdoChannelType").Click 12,16 @@ hightlight id_;_1771352_;_script infofile_;_ZIP::ssf30.xml_;_
'SwfWindow("收费规则管理").SwfObject("添加(A)").Click 52,11 @@ hightlight id_;_2032994_;_script infofile_;_ZIP::ssf31.xml_;_
'SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("cmbChargeType").Click 161,9 @@ hightlight id_;_920604_;_script infofile_;_ZIP::ssf32.xml_;_
'SwfWindow("收费规则管理").SwfTable("gridControlChargingRules").SetFilter "收费规则名称","按期收费" @@ hightlight id_;_920100_;_script infofile_;_ZIP::ssf33.xml_;_
'SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("cmbChargeRuleType").Click 38,9 @@ hightlight id_;_855064_;_script infofile_;_ZIP::ssf34.xml_;_
'SwfWindow("收费规则管理").SwfWindow("编辑收费规则").Close
'SwfWindow("收费规则管理").Close