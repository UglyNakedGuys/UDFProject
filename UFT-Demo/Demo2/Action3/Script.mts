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
'点击长期用户选择项
SwfWindow("收费规则管理").SwfObject("rdoChannelType").Click 39,12
wait 1
SwfWindow("收费规则管理").SwfObject("添加(A)").Click @@ hightlight id_;_788162_;_script infofile_;_ZIP::ssf5.xml_;_

do While true
	if(SwfWindow("收费规则管理").SwfWindow("编辑收费规则").Exist(1)) then
		Exit do
	end if
loop
wait 1

SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("cmbChargeType").Click @@ hightlight id_;_1507416_;_script infofile_;_ZIP::ssf6.xml_;_
wait 1
'点击下拉框 按期收费选项
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 66,8 @@ hightlight id_;_788372_;_script infofile_;_ZIP::ssf7.xml_;_
WriteLogs("长期收费选项参数设置")

wait 1
ruleName=datatable.GetSheet("长期按期收费").GetParameter("规则名称").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEdit("SwfEdit_3").Set ruleName
WriteLogs("长期收费规则名称参数设置")

wait 1
' 规则类型处理
chargeType =datatable.GetSheet("长期按期收费").GetParameter("规则类型").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("cmbChargeRuleType").Click @@ hightlight id_;_329820_;_script infofile_;_ZIP::ssf8.xml_;_

Select Case chargeType
	Case "按天" SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("按期收费规则").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 48,29
	Case "按月" SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("按期收费规则").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 46,42
	WriteLogs("长期收费类型参数设置")
End Select
 @@ hightlight id_;_1116184_;_script infofile_;_ZIP::ssf11.xml_;_


txtCharge=datatable.GetSheet("长期按期收费").GetParameter("每期金额").ValueByRow(1)
WriteLogs("长期收费详细参数设置")
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txt_Charge").DblClick 5,5,micLeftBtn
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txt_Charge").Type micDel 
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txt_Charge").Type txtCharge

wait 1
txtDay =datatable.GetSheet("长期按期收费").GetParameter("有效期限").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txt_Days").DblClick 5,5,micLeftBtn
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txt_Days").Type micDel 
wait 1
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("txt_Days").Type txtDay

wait 1

isFixed=datatable.GetSheet("长期按期收费").GetParameter("固定车位").ValueByRow(1)
If isFixed <>"" Then
	If  SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("chkIsFixed").GetROProperty("Checked")<>True Then
		SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("chkIsFixed").Click
	End If
End If 

wait 1
parkingType =datatable.GetSheet("长期按期收费").GetParameter("停车类型").ValueByRow(1)
SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("cmbParkingType").Click @@ hightlight id_;_329820_;_script infofile_;_ZIP::ssf8.xml_;_
Select Case parkingType
			Case "全天包月"  SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("按期收费规则").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 55,25
			Case "白天包月"  SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("按期收费规则").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 40,37
			Case "晚上包月" SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("按期收费规则").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 34,55
			Case "其他"  SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("按期收费规则").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 34,75
End Select
wait 1

txRule =datatable.GetSheet("长期按期收费").GetParameter("规则描述").ValueByRow(1)
If txRule<> "" Then
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEditor("SwfEditor").DblClick 5,5,micLeftBtn
	wait 1
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEditor("SwfEditor").Type micDel 
	wait 1
	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfEditor("SwfEditor").Type txRule
	wait 1
End If

SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfObject("确定").Click

Do While true
	'保存系统参数提示
	if(SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("提示信息").Exist(1)) then
		wait 1
		SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("提示信息").SwfObject("OK").Click @@ hightlight id_;_4589700_;_script infofile_;_ZIP::ssf15.xml_;_
		wait 1
		If  (SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("提示信息").Exist(1)) Then
			wait 1
			SwfWindow("收费规则管理").SwfWindow("编辑收费规则").SwfWindow("提示信息").SwfObject("OK").Click @@ hightlight id_;_4589700_;_script infofile_;_ZIP::ssf15.xml_;_
		End If
'	SwfWindow("收费规则管理").SwfWindow("编辑收费规则").Close
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
	datatable.GetSheet("长期按期收费").SetCurrentRow(1)
	datatable.Value("结果","长期按期收费")="成功"
else
	reporter.ReportEvent  micFail ,"Add","添加失败！"
	datatable.LocalSheet.AddParameter "结果"," "
	datatable.GetSheet("长期按期收费").SetCurrentRow(1)
	datatable.Value("结果","长期按期收费")="失败"
End If

datatable.Export("E:\Jangboer201705\UFT-Demo\Excel\长期按期收费.xls")


