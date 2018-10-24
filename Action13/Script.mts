'调用鼠标键盘操作函数
'KeyBoard通过键盘输入数据
'DelExist清除以存在数据
'RightKey右键点击
vbsPath = DataTable("A", dtGlobalSheet)
excelPath = DataTable("B", dtGlobalSheet)
ExecuteFile(vbsPath)

'新建布控任务
Function NewTask(taskName, taskType, taskDescription, storesArr, regionsArr, strategyName)
	NewTaskRet = False
	WpfWindow("窗体列表").WpfButton("+新建任务").Click
	If not WpfWindow("布控任务").Exist(1) Then
		Reporter.ReportEvent micFail, "新建布控任务", "无新建布控任务列表"
		ExitAction(0)
	End If
	'新建布控任务
	WpfWindow("布控任务").WpfComboBox("cbPlan").Select(taskType)
	WpfWindow("布控任务").WpfEdit("任务名称").Set(CStr(taskName))
	WpfWindow("布控任务").WpfEdit("任务描述").Set(CStr(taskDescription))
	'选择目标人
	WpfWindow("布控任务").WpfButton("选择目标人").Click
	If not WpfWindow("选择目标人").Exist(1) Then
		Reporter.ReportEvent micFail, "新建布控任务", "无目标人列表"
		ExitAction(0)
	End If
	WpfWindow("选择目标人").InsightObject("InsightObject").Click
	Set children = WpfWindow("选择目标人").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		For arrIndex = LBound(storesArr) To UBound(storesArr) Step 1
			If children(i).GetVisibleText = storesArr(arrIndex) Then
				children(i).Click
				NewTaskRet = True
			End If
		Next
	Next
	If not NewTaskRet Then
		Reporter.ReportEvent micFail, "新建布控任务", "无所要目标库"
		ExitAcion(0)
	End If
	Set children = Nothing
	NewTaskRet = False
	WpfWindow("选择目标人").WpfButton("确认").Click

	Wait(1)
	
	'选择区域通道
	WpfWindow("布控任务").WpfButton("选择通道").Click
	If not WpfWindow("区域通道").Exist(1) Then
		Reporter.ReportEvent micFail, "新建布控任务", "无区域通道列表"
		ExitAction(0)
	End If
	WpfWindow("区域通道").InsightObject("InsightObject").Click
	Set children = WpfWindow("区域通道").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		For arrIndex = LBound(regionsArr) To UBound(regionsArr) Step 1
			If children(i).GetVisibleText = regionsArr(arrIndex) Then
				children(i).Click
				NewTaskRet = True
			End If
		Next
	Next
	If not NewTaskRet Then
		Reporter.ReportEvent micFail, "新建布控任务", "无所要区域"
		ExitAction(0)
	End If
	Set children = Nothing
	NewTaskRet = False
	WpfWindow("区域通道").WpfButton("确定").Click
	Wait(1)
	
	WpfWindow("布控任务").WpfButton("保存").Click
	'校验是否新建完成
	Wait(1)
	WpfWindow("窗体列表").WpfButton("刷新").Click
	WpfWindow("窗体列表").WpfEdit("txtKeyWord").Click
	DelExist()
	KeyBoard(CStr(taskName))
	KeyBoard("{ENTER}")
	Wait(1)
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("value") = CStr(taskName) Then
			NewTaskRet = True
			Exit For
		End If
	Next
	If not NewTaskRet Then
		Reporter.ReportEvent micFail, "新建布控任务", "无所要策略"
		ExitAction(0)
	End If

End Function

'启动布控任务
Function StartTask(taskName)
	StartTaskRet = False
	WpfWindow("窗体列表").WpfEdit("txtKeyWord").Click
	DelExist()
	KeyBoard(CStr(taskName))
	KeyBoard("{ENTER}")
	Wait(1)
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("value") = CStr(taskName) Then
			children(i).Click
			WpfWindow("窗体列表").WpfButton("启动").Click
			StartTaskRet = True
			Exit For
		End If
	Next
	If not StartTaskRet Then
		Reporter.ReportEvent micFail, "启动布控任务", "无所要布控任务"
		ExitAction(0)
	End If
	StartTaskRet = False
	Wait(1)
	
	WpfWindow("窗体列表").WpfEdit("txtKeyWord").Click
	DelExist()
	KeyBoard(CStr(taskName))
	KeyBoard("{ENTER}")
	Wait(1)
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("text") = "布控中" Then
			StartTaskRet = True
			Exit For
		End If
	Next
End Function

'关闭布控任务
Function StopTask(taskName)
	StopTaskRet = False
	WpfWindow("窗体列表").WpfEdit("txtKeyWord").Click
	DelExist()
	KeyBoard(CStr(taskName))
	KeyBoard("{ENTER}")
	Wait(1)
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("value") = CStr(taskName) Then
			children(i).Click
			WpfWindow("窗体列表").WpfButton("关闭").Click
			StopTaskRet = True
			Exit For
		End If
	Next
	If not StopTaskRet Then
		Reporter.ReportEvent micFail, "关闭布控任务", "无所要布控任务"
		ExitAction(0)
	End If
	StopTaskRet = False
	Wait(1)
	
	WpfWindow("窗体列表").WpfEdit("txtKeyWord").Click
	DelExist()
	KeyBoard(CStr(taskName))
	KeyBoard("{ENTER}")
	Wait(1)
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("text") = "关闭" Then
			StopTaskRet = True
			Exit For
		End If
	Next
End Function

'删除布控任务
Function DeleteTask(taskName)
	DeleteTaskRet = False
	WpfWindow("窗体列表").WpfEdit("txtKeyWord").Click
	DelExist()
	KeyBoard(CStr(taskName))
	KeyBoard("{ENTER}")
	Wait(1)
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("value") = CStr(taskName) Then
			children(i).Click
			WpfWindow("窗体列表").WpfButton("关闭").Click
			DeleteTaskRet = True
			Exit For
		End If
	Next
	If not DeleteTaskRet Then
		Reporter.ReportEvent micFail, "删除布控任务", "无所要布控任务"
		ExitAction(0)
	End If
	DeleteTaskRet = False
	Wait(1)
	
	WpfWindow("窗体列表").WpfEdit("txtKeyWord").Click
	DelExist()
	KeyBoard(CStr(taskName))
	KeyBoard("{ENTER}")
	Wait(1)
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("text") = "更多" Then
			Children(i).Click
			For j = 1 To 3 Step 1
				KeyBoard("{DOWN}")
			Next
			KeyBoard("{ENTER}")
			DeleteTaskRet = True
			Exit For
		End If
	Next
	
	Wait(1)
	If WpfWindow("提示").Exist(1) Then
		WpfWindow("提示").WpfButton("是").Click
	End If
End Function

'编辑布控任务
Function ModifyTask(taskName, taskDescription, regionsArr, storesArr, strategyName)
	ModifyTaskRet = False
	'按任务名搜索布控任务
	WpfWindow("窗体列表").WpfEdit("txtKeyWord").Click
	DelExist()
	KeyBoard(CStr(taskName))
	KeyBoard("{ENTER}")
	Wait(1)
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("text") = "更多" Then
			Children(i).Click
			For j = 1 To 2 Step 1
				KeyBoard("{DOWN}")
			Next
			KeyBoard("{ENTER}")
			ModifyTaskRet = True
			Exit For
		End If
	Next
	If not ModifyTaskRet Then
		Reporter.ReportEvent micFail, "编辑布控任务", "无所要的布控任务"
		ExitAction(0)
	End If
	ModifyTaskRet = False
	Wait(1)
	WpfWindow("布控任务").WpfEdit("任务描述").Set(CStr(taskDescription))

	
	If not WpfWindow("布控任务").Exist(1) Then
		Reporter.ReportEvent micFail, "编辑布控任务", "无布控任务列表"
		ExitAction(0)
	End If
	
	'选择目标人
	WpfWindow("布控任务").WpfButton("选择目标人").Click
	If not WpfWindow("选择目标人").Exist(1) Then
		Reporter.ReportEvent micFail, "编辑布控任务", "无选择目标人列表"
		ExitAction(0)
	End If
	Wait(1)
	WpfWindow("选择目标人").InsightObject("InsightObject").Click
	
	Set children = WpfWindow("选择目标人").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		For arrIndex = LBound(storesArr) To UBound(storesArr) Step 1
		
			templateTmp = storesArr(arrIndex)
			If children(i).GetROProperty("value") = templateTmp Then
				children(i).Click
				ModifyTaskRet = True
			End If
		Next
	Next
	Set children = nothing
	If not ModifyTaskRet Then
		Reporter.ReportEvent micFail, "编辑布控任务", "无所要的目标人"
		ExitAction(0)
	End If
	ModifyTaskRet = False
	WpfWindow("选择目标人").WpfButton("确认").Click
	Wait(1)
	
	'选择区域通道
	WpfWindow("布控任务").WpfButton("选择通道").Click
	If not WpfWindow("区域通道").Exist(1) Then
		Reporter.ReportEvent micFail, "编辑布控任务", "无选择区域通道列表"
		ExitAction(0)
	End If
	Wait(1)
	WpfWindow("区域通道").InsightObject("InsightObject").Click
	Set child = WpfWindow("区域通道").WpfObject("PART_ItemsScrollViewer").ChildObjects
	'Set child = WpfWindow("区域通道").ChildObjects
	count = child.Count

	For i = 0 To count - 1 Step 1
		For arrIndex = LBound(regionsArr) To UBound(regionsArr) Step 1
			regionTmp = regionsArr(arrIndex)
			If child(i).GetVisibleText = regionTmp Then
				child(i).Click
				ModifyTaskRet = True
			End If
		Next
		
	Next
	Set child = nothing
	If not ModifyTaskRet Then
		Reporter.ReportEvent micFail, "编辑布控任务", "无所要的区域通道"
		ExitAction(0)
	End If
	WpfWindow("区域通道").WpfButton("确定").Click
	Wait(3)

	'选择比对策略
	WpfWindow("布控任务").RefreshObject
	countStrategy = WpfWindow("布控任务").WpfComboBox("cbCmpList").GetItemsCount
	For i = 0 To countStrategy - 1 Step 1
		WpfWindow("布控任务").WpfComboBox("cbCmpList").Select(i)
		If WpfWindow("布控任务").WpfComboBox("cbCmpList").GetROProperty("selection") = strategyName Then
			ModifyTaskRet = True
			Exit For
		End If
	Next
	If not ModifyTaskRet Then
		Reporter.ReportEvent micFail, "编辑布控任务", "无所要的比对策略"
		ExitAction(0)
	End If
	
	WpfWindow("布控任务").WpfButton("保存").Click
	
End Function


'抓取异常
'On Error Resume Next
On Error Goto 0

WpfWindow("窗体列表").Maximize
WpfWindow("窗体列表").WpfObject("布控任务").Click
WpfWindow("窗体列表").WpfTabStrip("WpfTabStrip").Select(0)

'NewTaskRet = False
'taskName = Parameter("taskName")
'taskType = Parameter("taskType")
'handle = Parameter("handle")
'If taskType = 0 Then
'	taskTypeName = "永久任务"
'ElseIf taskType = 1 Then
'	taskTypeName = "自定义任务"
'End If
'0：新建永久任务；1：新建自定义任务
'Call NewTask(taskName, taskType)
'If NewTaskRet Then
'	Reporter.ReportEvent micPass, "新建布控任务", "新建" & taskTypeName & "成功"
'	ExitAction(1)
'Else
'	Reporter.ReportEvent micFail, "新建布控任务", "新建" & taskTypeName & "失败"
'	ExitAction(0)
'End If
Dim storesArr, regionsArr
Select Case Parameter("handle")
	
	Case "newTask"
		print("新建布控任务")
		countExcute = ReadExcel(3, 7, "新建布控任务", excelPath)
		For caseIndex = 0 To countExcute - 1 Step 1
			If ReadExcel(3 + caseIndex, 1, "新建布控任务", excelPath) = "永久任务" Then
				taskType = 0
			ElseIf ReadExcel(3 + caseIndex, 1, "新建布控任务", excelPath) = "自定义任务" Then
				taskType = 1
			End If
			taskName = ReadExcel(3 + caseIndex, 2, "新建布控任务", excelPath)
			taskDescription = ReadExcel(3 + caseIndex, 3, "新建布控任务", excelPath)
			
			storesArr = Split(ReadExcel(3 + caseIndex, 4, "新建布控任务", excelPath), ",")
'			For arrIndex = LBound(storesArr) To UBound(storesArr) Step 1
'				print(storesArr(arrIndex))
'			Next
'			ExitRun
			regionsArr = Split(ReadExcel(3 + caseIndex, 5, "新建布控任务", excelPath), ",")
			strategyName = ReadExcel(3, 6, "新建布控任务", excelPath)
			NewTaskRet = False
			Call NewTask(taskName, taskType, taskDescription, storesArr, regionsArr, strategyName)
			If NewTaskRet Then
				Reporter.ReportEvent micPass, "新建布控任务", "新建" & taskTypeName & "成功"
				'ExitAction(1)
			Else
				Reporter.ReportEvent micFail, "新建布控任务", "新建" & taskTypeName & "失败"
				'ExitAction(0)
			End If
		Next
		If NewTaskRet Then
			ExitAction(1)
		Else
			ExitAction(0)
		End If
		
		
	Case "startTask"
		countExcute = ReadExcel(3, 2, "启动布控任务", excelPath)
		For caseIndex = 0 To countExcute - 1 Step 1
			StartTaskRet = False
			taskName = ReadExcel(3 + caseIndex, 1, "启动布控任务", excelPath)
			Call StartTask(taskName)
			If StartTaskRet Then
				Reporter.ReportEvent micPass, "启动布控任务", "启动布控任务-" & taskName & "-成功"
			Else
				Reporter.ReportEvent micFail, "启动布控任务", "启动布控任务-" & taskName & "-失败"
			End If
		Next
		If StartTaskRet Then
			ExitAction(1)
		Else
			ExitAction(0)
		End If
	
		
	Case "stopTask"
		countExcute = ReadExcel(3, 2, "关闭布控任务", excelPath)
		For caseIndex = 0 To countExcute - 1 Step 1
			StopTaskRet = False
			taskName = ReadExcel(3 + caseIndex, 1, "关闭布控任务", excelPath)
			Call StopTask(taskName)
			If StopTaskRet Then
				Reporter.ReportEvent micPass, "关闭布控任务", "关闭布控任务-" & taskName & "-成功"
			Else
				Reporter.ReportEvent micFail, "关闭布控任务", "关闭布控任务-" & taskName & "-失败"
			End If
		Next
		If StopTaskRet Then
			ExitAction(1)
		Else
			ExitAction(0)
		End If
		
		
	Case "deleteTask"
		countExcute = ReadExcel(3, 2, "删除布控任务", excelPath)
		For caseIndex = 0 To countExcute - 1 Step 1
			DeleteTaskRet = False
			taskName = ReadExcel(3 + caseIndex, 1, "删除布控任务", excelPath)
			Call DeleteTask(taskName)
			If DeleteTaskRet Then
				Reporter.ReportEvent micPass, "删除布控任务", "删除布控任务-" & taskName & "-成功"
			Else
				Reporter.ReportEvent micFail, "删除布控任务", "删除布控任务-" & taskName & "-失败"
			End If
		Next
		If DeleteTaskRet Then
			ExitAction(1)
		Else
			ExitAction(0)
		End If
		
	
	Case "modifyTask"
		countExcute = ReadExcel(3, 6, "修改布控任务", excelPath)
		For caseIndex = 0 To countExcute - 1 Step 1
			taskName = ReadExcel(3 + caseIndex, 1, "修改布控任务", excelPath)
			taskDescription = ReadExcel(3 + caseIndex, 2, "修改布控任务", excelPath)
			storesArr = Split(ReadExcel(3 + caseIndex, 3, "修改布控任务", excelPath), ",")
			regionsArr = Split(ReadExcel(3 + caseIndex, 4, "修改布控任务", excelPath), ",")
			strategyName = ReadExcel(3, 5, "修改布控任务", excelPath)
			ModifyTaskRet = False
			Call ModifyTask(taskName, taskDescription, regionsArr, storesArr, strategyName)
			If ModifyTaskRet Then
				Reporter.ReportEvent micPass, "修改布控任务", "修改布控任务-" & taskName & "-成功"
			Else
				Reporter.ReportEvent micFail, "修改布控任务", "修改布控任务-" & taskName & "-失败"
			End If
		Next
		If ModifyTaskRet Then
			ExitAction(1)
		Else
			ExitAction(0)
		End If
		

	Case Else
		Reporter.ReportEvent micFail, "布控任务", "输入布控任务操作类型错误"
		ExitAction(0)
End Select


If Err.Number <> 0 Then
	errMessage = "错误代码： " & CStr(Err.Number) & ", 错误信息： " & Err.Description & "."
	print(errMessage)
	ExitAction(-1)
End If