'调用鼠标键盘操作函数
'KeyBoard通过键盘输入数据
'DelExist清除以存在数据
'RightKey右键点击
ExecuteFile("D:\MouseKeyboard.vbs")

'连接mysql数据库
Function MySQLConn(sqlQuary, uuid)
	MySQLConnRet = False
	Dim Cnn, Rst, strCnn
	strCnn = "Driver={MySQL ODBC 5.1 Driver};DATABASE=facecore;PWD=1qazXSW@;PORT=3306;SERVER=192.168.0.15;UID=root"
	Set Cnn = CreateObject("ADODB.Connection")
	Cnn.Open strCnn
	Set Rst = CreateObject("ADODB.Recordset")
	Rst.Open sqlQuary, Cnn, 1, 1
	Rst.MoveFirst

	If Rst(0) = uuid Then
		'print("True")
		MySQLConnRet = True
	Else
		Reporter.ReportEvent micFail, "详细信息", "uuid查询无结果"
		ExitAction(0)
	End If
	Rst.Close
	Cnn.Close
	
	Set Rst = Nothing
	Set Cnn = Nothing

End Function

'核查反馈
Function VerificationCheck()
	VerificationCheckRet = False
	WpfWindow("窗体列表").WpfRadioButton("全部").Click
	'查询需要核查的报警
	Set children = WpfWindow("窗体列表").WpfObject("SofaContainerContent1").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		'print(children(i).GetROProperty("text"))
		If children(i).GetROProperty("text") = "核查反馈" Then
			children(i).Click
			VerificationCheckRet = True
			Exit For
		End If
	Next
	If not VerificationCheckRet Then
		Reporter.ReportEvent micFail, "预警历史", "无需要核查的报警"
		ExitAction(0)
	End If
	VerificationCheckRet = False
	Set children = Nothing
	Wait(1)
	'核查
	WpfWindow("核查记录").WpfEdit("txtT").Set("目标以确认")
	Wait(1)
	WpfWindow("核查记录").WpfButton("确认").Click
	VerificationCheckRet = True
End Function

'详细信息
Function AlertInformation()
	AlertInformationRet = False
	'查询详细信息
	Set children = WpfWindow("窗体列表").WpfObject("SofaContainerContent1").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		'print(children(i).GetROProperty("text"))
		If children(i).GetROProperty("text") = "详细信息" Then
			children(i).Click
			AlertInformationRet = True
			Exit For
		End If
	Next
	If not AlertInformationRet Then
		Reporter.ReportEvent micFail, "预警历史", "无需要核查的报警"
		ExitAction(0)
	End If
	AlertInformationRet = False
	Set children = Nothing
	Wait(1)
	
	If not WpfWindow("popwindow").Exist(3) Then
		Reporter.ReportEvent micFail, "详细信息", "无详细信息界面"
		ExitAction(0)
	End If
	
	WpfWindow("popwindow").WpfTabStrip("WpfTabStrip").Select(0)
	Set children = WpfWindow("popwindow").WpfTabStrip("WpfTabStrip").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
'		print(i)
'		print(children(i).GetROProperty("text")) 
		If children(i).GetROProperty("text") = "告警时间:" Then
			uuid = children(i - 1).GetROProperty("text") 
			Exit For
		End If
	Next
	
	sqlQuary = "select uuid from original_alerts where uuid = '" & uuid & "'"
	Call MySQLConn(sqlQuary, uuid)
	AlertInformationRet = True
	
	WpfWindow("popwindow").WpfTabStrip("WpfTabStrip").Select(1)
	Wait(1)
	WpfWindow("popwindow").WpfTabStrip("WpfTabStrip").Select(2)
	Wait(1)
	WpfWindow("popwindow").WpfTabStrip("WpfTabStrip").Select(3)
	Wait(1)
	WpfWindow("popwindow").Close
End Function


'未复核个数确认
Function UnCheckNum()
	UnCheckNumRet = False
	WpfWindow("窗体列表").WpfRadioButton("WpfRadioButton").Click
	Set children = WpfWindow("窗体列表").WpfRadioButton("WpfRadioButton").ChildObjects

	startIndex = instr(children(1).GetROProperty("text"), "(")
	endIndex = instr(children(1).GetROProperty("text"), ")")

	srcNum = mid(children(1).GetROProperty("text"), startIndex + 1, endIndex - 2)
	Dim strArr
	strArr = split(WpfWindow("窗体列表").WpfObject("导出数据").GetVisibleText, " ")
	'dstNum = strArr(2)
	For i = LBound(strArr) To UBound(strArr) - 1 Step 1
		'print("--->" & CStr(i))
		If srcNum = strArr(i) Then
			UnCheckNumRet = True
			Exit For
		End If
	Next
	print("		状态栏未复核告警个数：" & CStr(srcNum))
'	print("		下角标个数：" & CStr(dstNum))
'	If srcNum = dstNum Then
'		UnCheckNumRet = True
'	End If

End Function


'抓取异常
'On Error Resume Next
On Error Goto 0
WpfWindow("窗体列表").Maximize
WpfWindow("窗体列表").WpfObject("预警历史").Click

Select Case Parameter("handle")
	
	Case "verificationCheck"
		VerificationCheckRet = False
		Call VerificationCheck()
		If VerificationCheckRet Then
			Reporter.ReportEvent micPass, "核查反馈", "核查反馈成功"
			ExitAction(1)
		Else
			Reporter.ReportEvent micFail, "核查反馈", "核查反馈失败"
			ExitAction(0)
		End If
	
	Case "alertInformation"	
		AlertInformationRet = False
		Call AlertInformation()
		If AlertInformationRet Then
			Reporter.ReportEvent micPass, "详细信息", "查询详细信息成功"
			ExitAction(1)
		Else
			Reporter.ReportEvent micFail, "详细信息", "查询详细信息失败"
			ExitAction(0)
		End If
	
	Case "unCheckNum"
		UnCheckNumRet = False
		Call UnCheckNum()
		If UnCheckNumRet Then
			Reporter.ReportEvent micPass, "未复核告警个数确认", "未复核告警个数确认成功"
			ExitAction(1)
		Else
			Reporter.ReportEvent micPass, "未复核告警个数确认", "未复核告警个数确认失败"
			ExitAction(0)
		End If		
	
	Case Else
		Reporter.ReportEvent micFail, "预警历史", "输入预警历史操作类型错误"
		ExitAction(0)
End Select

If Err.Number <> 0 Then
	errMessage = "错误代码： " & CStr(Err.Number) & ", 错误信息： " & Err.Description & "."
	print(errMessage)
	ExitAction(-1)
End If