WriteLogs("模块头：操作员管理测试模块！")

'初始化操作
If SwfWindow("视频识别出入口管理系统").Exist(1) Then
	SystemUtil.CloseProcessByName("PakingVideo_Login.exe")
	WriteLogs("初始化关闭操作！")
End If

'程序启动
Wait 2
SystemUtil.Run("C:\Users\Administrator\Desktop\兔巢acs3.1\软件201705181100-Debug\Debug\PakingVideo_Login.exe")
WriteLogs("启动程序")

'登录操作
Wait 2
SwfWindow("登录界面").SwfEdit("SwfEdit").Set "admin"
SwfWindow("登录界面").SwfEdit("SwfEdit_2").Set "admin"
SwfWindow("登录界面").SwfObject("登录").Click

'操作员管理操作
Wait 3
SwfWindow("视频识别出入口管理系统").Activate
WriteLogs("操作员管理入口")
SwfWindow("视频识别出入口管理系统").SwfObject("系统设置").Click
SwfWindow("视频识别出入口管理系统").SwfObject("btnMenu2").Click 

WriteLogs("=======================================================")
WriteLogs("模块体：添加操作迭代！")
Wait 1
SwfWindow("操作员管理").Activate
SwfWindow("操作员管理").SwfObject("添加(A)").Click 

WriteLogs("添加数据迭代开始！")
'对象赋值
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit").Set Datatable("Name","AddUser")
Wait 0.5
SwfWindow("操作员管理").SwfWindow("操作员编辑").Activate
SwfWindow("操作员管理").Activate

Dim txRole
txRole=Datatable("Role","AddUser")
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("cmbRole").Click
Wait 1
If txRole="系统管理员" Then
	SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 110,10
	'SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("cmbRole").Click 110,10
Else
	'SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("cmbRole").Click 110,27
	SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 110,27
End If 
Wait 0.5
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit_2").Set Datatable("LoginName","AddUser")
Wait 0.5
SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit_3").Set Datatable("LoginPwd","AddUser")
Wait 1

SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("保存(S)").Click 
'点击提示框操作
If SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("提示信息").Exist(1) Then
	Wait 1
	SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("提示信息").SwfObject("OK").Click
'判断是否为添加失败操作
	If SwfWindow("操作员管理").SwfWindow("操作员编辑").Exist(1) Then
		WriteLogs("添加操作失败，数据库中已存在"&Datatable("Name","AddUser")&"该用户！")
		Wait 1
		SwfWindow("操作员管理").SwfWindow("操作员编辑").Close()
	Else
		WriteLogs("操作员"& Datatable("Name","AddUser") &"添加成功！")
	End If
End If

WriteLogs("添加数据迭代第"& Datatable.GetSheet("AddUser").GetParameter("Num").ValueByRow(1) &"次完毕")
WriteLogs("-------------------------------------------------------")
'检查添加操作时间
Wait 3
WriteLogs("模块体：编辑操作迭代！")
Wait 1
Dim Name

'定义需要编辑的条目名
Name=Datatable("Name","AddUSer")

For Iterator = 0 To SwfWindow("操作员管理").SwfTable("gridControlOperator").RowCount-1

	If SwfWindow("操作员管理").SwfTable("gridControlOperator").GetCellData(Iterator,1)=Name Then
		
		WriteLogs("编辑数据迭代开始！")
		SwfWindow("操作员管理").SwfTable("gridControlOperator").ActivateCell Iterator,1
'对象赋值
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit").Set Datatable("EditName","AddUser")
		Wait 0.5
		txRole=Datatable("EditRole","AddUser")
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("cmbRole").Click
		If txRole="系统管理员" Then
			SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 110,10
			'SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("cmbRole").Click 110,10
		Else
			'SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("cmbRole").Click 110,27
			SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("SwfWindow").SwfObject("SwfObject").Click 110,27
		End If 
		Wait 0.5
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit_2").Set Datatable("EditLoginName","AddUser")
		Wait 0.5
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfEdit("SwfEdit_3").Set Datatable("EditLoginPwd","AddUser")
		Wait 1
		SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfObject("保存(S)").Click 
'点击提示框操作
		If SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("提示信息").Exist(1) Then
			Wait 1
			SwfWindow("操作员管理").SwfWindow("操作员编辑").SwfWindow("提示信息").SwfObject("OK").Click
'判断是否为编辑失败操作 
			If SwfWindow("操作员管理").SwfWindow("操作员编辑").Exist(1) Then
				WriteLogs("编辑操作失败"&"用户"&Datatable("Name","AddUser"))
				Wait 1
				SwfWindow("操作员管理").SwfWindow("操作员编辑").Close()
			Else
				WriteLogs("操作员"& Datatable("EditName","AddUser") &"编辑成功！")
			End If
		End If
	Else
		WriteLogs("当前Table第"&Iterator&"行Name参数不对，未能进入循环!")
	End If
Next
WriteLogs("编辑数据迭代第"& Datatable.GetSheet("AddUser").GetParameter("Num").ValueByRow(1) &"次完毕")
WriteLogs("-------------------------------------------------------")
'迭代次数参数记录
Datatable.GetSheet("AddUser").GetParameter("Num").ValueByRow(1)=Datatable.GetSheet("AddUser").GetParameter("Num").ValueByRow(1)+1


'删除操作
For Iterator = 0 To SwfWindow("操作员管理").SwfTable("gridControlOperator").RowCount-1
	If SwfWindow("操作员管理").SwfTable("gridControlOperator").GetCellData(Iterator,1)=Name Then
		SwfWindow("操作员管理").SwfTable("gridControlOperator").SelectCell Iterator,1
		Wait 1
		SwfWindow("操作员管理").SwfObject("删除(D)").Click 
		Wait 1
		If SwfWindow("操作员管理").SwfWindow("确认信息").Exist(1) Then
			WriteLogs("删除相应操作员，Name为"&Name)
		      SwfWindow("操作员管理").SwfWindow("确认信息").SwfObject("Yes").Click
		      If SwfWindow("操作员管理").SwfWindow("提示信息").Exist(1)  Then
		      	    	WriteLogs("删除成功！")
		      Else
		     		WriteLogs("删除失败！")
		      End If
              End If
	End If
Next

SwfWindow("操作员管理").Close()


 @@ hightlight id_;_460468_;_script infofile_;_ZIP::ssf2.xml_;_