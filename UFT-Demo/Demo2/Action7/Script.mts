SwfWindow("视频识别出入口管理系统").SwfObject("实时监控").Click @@ hightlight id_;_2624792_;_script infofile_;_ZIP::ssf1.xml_;_
wait 1
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu1").Click @@ hightlight id_;_1969446_;_script infofile_;_ZIP::ssf2.xml_;_

 @@ hightlight id_;_132046_;_script infofile_;_ZIP::ssf1.xml_;_
Do While true
	if(SwfWindow("实时监控界面").Exist(1)) then
		Exit do
	end if
loop
wait 1

If (SwfWindow("实时监控界面").SwfObject("查询").GetROProperty("Visible")=true) Then
	wait 2
	SwfWindow("实时监控界面").SwfObject("查询").Click
End If

'打印小票的CheckBox去掉不要
wait 1
If SwfWindow("实时监控界面").SwfObject("chkPrint").GetROProperty("Checked") Then
	SwfWindow("实时监控界面").SwfObject("chkPrint").Click @@ hightlight id_;_986666_;_script infofile_;_ZIP::ssf7.xml_;_
	wait 1
End If

tableRowCount=SwfWindow("实时监控界面").SwfTable("gridControlInRecord").RowCount
If  tableRowCount>0 Then
	' 可以弄一个循环点击===
	wait 2
	SwfWindow("实时监控界面").SwfTable("gridControlInRecord").SetView "" @@ hightlight id_;_1707426_;_script infofile_;_ZIP::ssf5.xml_;_
	wait 1
	SwfWindow("实时监控界面").SwfTable("gridControlInRecord").SelectCell 0,"确定" @@ hightlight id_;_1707426_;_script infofile_;_ZIP::ssf6.xml_;_
	wait 1
	SwfWindow("实时监控界面").SwfObject("缴费").Click @@ hightlight id_;_3279462_;_script infofile_;_ZIP::ssf8.xml_;_
End If


SystemUtil.CloseProcessByName("PakingVideo_Login.exe")
