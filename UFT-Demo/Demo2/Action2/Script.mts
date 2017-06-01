'===========================================================================================================================
' 登录部分
'If  SwfWindow("登录界面").Exist(2) Then
' SwfWindow("登录界面").Close
'End If
'wait 1
'systemutil.Run("E:\Project\标准版Test_不提交\广州粤电大厦\V20170502\PakingVideo_Login\bin\Debug\PakingVideo_Login.exe")
'wait	2
'SwfWindow("登录界面").SwfEdit("SwfEdit_2").Set "admin"
'SwfWindow("登录界面").SwfEdit("SwfEdit").SetSecure "5924e4b99935029c317c8fdbcdda0b6b"
'wait 1
'
'SwfWindow("登录界面").SwfObject("登录").Click
'
'do While true
'	if(SwfWindow("视频识别出入口管理系统").Exist(1)) then
'		Exit do
'	end if
'loop
'wait 1
'===========================================================================================================================


SwfWindow("视频识别出入口管理系统").SwfObject("系统设置").Click @@ hightlight id_;_132940_;_script infofile_;_ZIP::ssf23.xml_;_
wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu4").Click @@ hightlight id_;_67462_;_script infofile_;_ZIP::ssf24.xml_;_
WriteLogs("收费规则管理入口！")
WriteLogs("-------------------------------------------------------")

Do While true
	if(SwfWindow("收费规则管理").Exist(1)) then
		Exit do
	end if
loop
wait 1
'点击临时用户选择项
SwfWindow("收费规则管理").SwfObject("rdoChannelType").Click 124,12 @@ hightlight id_;_1509150_;_script infofile_;_ZIP::ssf25.xml_;_
'SwfWindow("收费规则管理").SwfObject("rdoChannelType").Click
wait 1
SwfWindow("收费规则管理").SwfObject("添加(A)").Click
do While true
	if(SwfWindow("收费规则管理").SwfWindow("编辑收费规则").Exist(1)) then
		Exit do
	end if
loop
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("cmbChargeType").Click
wait 1
' 点击下拉框的时间收费选项
SwfWindow("SwfWindow").SwfObject("SwfObject").Click 65,8
WriteLogs("临时收费选项参数设置！")
wait 1
ruleName=datatable.GetSheet("临时时长收费").GetParameter("规则名称").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_8").Set ruleName
WriteLogs("临时收费规则名称参数设置！")
wait 1
' 点击分时收费RadioButton
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("rdoChargeType").Click 36,13
WriteLogs("分时收费参数设置！")

wait 1
parkTimeFree=datatable.GetSheet("临时时长收费").GetParameter("停车时间不收费").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_7").type parkTimeFree
WriteLogs("临时收费时长参数设置！")

wait 1
totalCount=datatable.GetSheet("临时时长收费").GetParameter("总限额").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_6").type totalCount
WriteLogs("临时收费总限额参数设置！")

'checkBox 部分
fullCharge=datatable.GetSheet("临时时长收费").GetParameter("足时收费").ValueByRow(1)
freeTimeCharge=datatable.GetSheet("临时时长收费").GetParameter("免费时间收费").ValueByRow(1)
chargeByDays=datatable.GetSheet("临时时长收费").GetParameter("超24小时按天收费").ValueByRow(1)
chargeAllByDay=datatable.GetSheet("临时时长收费").GetParameter("按天").ValueByRow(1)
mergeSpanTime=datatable.GetSheet("临时时长收费").GetParameter("跨段合并").ValueByRow(1)
secondes=datatable.GetSheet("临时时长收费").GetParameter("向上进位").ValueByRow(1)
inner=datatable.GetSheet("临时时长收费").GetParameter("内部免费").ValueByRow(1)
WriteLogs("临时收费checkBox部分参数设置！")

If fullCharge <>"" Then
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("chkFullCharge").Click @@ hightlight id_;_393898_;_script infofile_;_ZIP::ssf16.xml_;_
End If

If  freeTimeCharge <>""Then
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("chkFreeTimeCharge").Click @@ hightlight id_;_459400_;_script infofile_;_ZIP::ssf17.xml_;_
End If

If chargeByDays <>"" Then @@ hightlight id_;_393898_;_script infofile_;_ZIP::ssf16.xml_;_
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("chkChargeByDays").Click @@ hightlight id_;_4849744_;_script infofile_;_ZIP::ssf18.xml_;_
End If

If  chargeAllByDay <>"" Then @@ hightlight id_;_459400_;_script infofile_;_ZIP::ssf17.xml_;_
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("chkChargeAllByDay").Click @@ hightlight id_;_328356_;_script infofile_;_ZIP::ssf19.xml_;_
End If
 @@ hightlight id_;_4849744_;_script infofile_;_ZIP::ssf18.xml_;_
If mergeSpanTime <>"" Then @@ hightlight id_;_328356_;_script infofile_;_ZIP::ssf19.xml_;_
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("chkMergeSpanTime").Click @@ hightlight id_;_2753832_;_script infofile_;_ZIP::ssf20.xml_;_
End If

If secondes <>"" Then @@ hightlight id_;_2753832_;_script infofile_;_ZIP::ssf20.xml_;_
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("chkSecondes").Click @@ hightlight id_;_8196046_;_script infofile_;_ZIP::ssf21.xml_;_
End If

If inner <>"" Then @@ hightlight id_;_8196046_;_script infofile_;_ZIP::ssf21.xml_;_
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("chkInner").Click @@ hightlight id_;_8586970_;_script infofile_;_ZIP::ssf22.xml_;_
End If

' 规则明细部分
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtStartTime").Click
For i=0 to 2 
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtStartTime").Type micLeft 
Next
startTime=datatable.GetSheet("临时时长收费").GetParameter("开始时间").ValueByRow(1)
startTimeArr = split(startTime,":")
For i=0 to ubound(startTimeArr)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtStartTime").Type startTimeArr(i)
wait 1
 SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtStartTime").Type micRight  
Next

wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtEndTime").Click
For i=0 to 2 
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtEndTime").Type micLeft 
Next
wait 1
endTime=datatable.GetSheet("临时时长收费").GetParameter("结束时间").ValueByRow(1)
endTimeArr = split(endTime,":")
For i=0 to ubound(endTimeArr)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtEndTime").Type endTimeArr(i)
wait 1
 SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtEndTime").Type micRight  
Next
wait 1

maxCharge=datatable.GetSheet("临时时长收费").GetParameter("最多收费").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtMaxCharge").DblClick 5,5,micLeftBtn
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtMaxCharge").Type micDel 
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtMaxCharge").Type maxCharge

wait 1
everyHour=datatable.GetSheet("临时时长收费").GetParameter("每小时").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtEveryHour").DblClick 5,5,micLeftBtn
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtEveryHour").Type micDel 
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtEveryHour").Type everyHour 

wait 1
chargeMoney=datatable.GetSheet("临时时长收费").GetParameter("收费").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtChargeMoney").DblClick 5,5,micLeftBtn
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtChargeMoney").Type micDel 
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txtChargeMoney").Type chargeMoney 

wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("添加").Click

wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("保存").Click

do While true
	if(SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("SwfWindow").SwfObject("OK").Exist(1)) then
		wait 1
		SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("SwfWindow").SwfObject("OK").Click @@ hightlight id_;_4589700_;_script infofile_;_ZIP::ssf15.xml_;_
		wait 1
		SwfWindow("收费规则管理").SwfWindow("编辑收费规则").Close
	Exit do
	end if
loop
wait 1

passFlag=false
tableRowCount=SwfWindow("收费规则管理").SwfTable("gridControlChargingRules").RowCount
For i=0 to tableRowCount -1
	tempRuleName=SwfWindow("收费规则管理").SwfTable("gridControlChargingRules").GetCellData(i,0)
	If  tempRuleName=ruleName Then
			passFlag=true
			Exit for
	End If
Next

SwfWindow("收费规则管理").Close
WriteLogs("收费规则管理出口！")
WriteLogs("-------------------------------------------------------")
wait 1

If passFlag Then
	reporter.ReportEvent micPass,"Add","添加成功！"
	datatable.LocalSheet.AddParameter "结果"," "
	datatable.GetSheet("临时时长收费").SetCurrentRow(1)
	datatable.Value("结果","临时时长收费")="成功"
else
	reporter.ReportEvent  micFail ,"Add","添加失败！"
	datatable.LocalSheet.AddParameter "结果"," "
	datatable.GetSheet("临时时长收费").SetCurrentRow(1)
	datatable.Value("结果","临时时长收费")="失败"
End If

datatable.Export("E:\Jangboer201705\UFT-Demo\Excel\临时时长收费.xls")
WriteLogs("收费规则数据表Excel导出！")