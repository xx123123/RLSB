'调用鼠标键盘操作函数
'KeyBoard通过键盘输入数据
'DelExist清除以存在数据
'RightKey右键点击
vbsPath = DataTable("A", dtGlobalSheet)
excelPath = DataTable("B", dtGlobalSheet)
ExecuteFile(vbsPath)

Function TaskDetails(taskName)
	TaskDetailsRet = False
	WpfWindow("窗体列表").WpfTabStrip("TaskList").Select(0)
	'选择所要查看的布控任务
	Set children = WpfWindow("窗体列表").WpfList("taskList").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("text") = CStr(taskName) Then
			children(i).Click
			TaskDetailsRet = True
			Exit For
		End If
	Next
	If not TaskDetailsRet Then
		Reporter.ReportEvent micFail, "布控详情", "无所要查看的布控任务"
		ExitAction(0)
	End If

	If WpfWindow("窗体列表").InsightObject("CancelAlert").Exist(1) Then
		WpfWindow("窗体列表").InsightObject("CancelAlert").Click
		WpfWindow("提示").WpfButton("确定").Click
	End If

	WpfWindow("窗体列表").WpfButton("修改").Click
	If WpfWindow("布控任务").Exist(1) Then
		WpfWindow("布控任务").WpfButton("保存").Click
	End If
	Wait(1)
	
	WpfWindow("窗体列表").WpfButton("调整策略参数").Click
	If WpfWindow("比对策略").Exist(1) Then
		WpfWindow("比对策略").WpfButton("保存").Click
	End If
	Wait(5)

	WpfWindow("窗体列表").InsightObject("OrderAlert").Click
	WpfWindow("提示").WpfButton("确定").Click
End Function

'视频预览
Function VideoPreview(regionName, channelName)
	VideoPreviewRet = False
	WpfWindow("窗体列表").WpfTabStrip("TaskList").Select(1)
	Wait(1)
	
	'选择区域
	Set children = WpfWindow("窗体列表").WpfTabStrip("TaskList").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If InStr(children(i).GetVisibleText, regionName) = 1 Then
			children(i).Click
			Wait(1)
			children(i).DblClick 2, 2
			VideoPreviewRet = True
			Wait(1)
			Exit For
		End If
	Next
	'ExitRun
	Set children = nothing
	If not VideoPreviewRet Then
		Reporter.ReportEvent micFail, "视频预览", "无所要区域"
		ExitAction(0)
	End If
	
	'选择通道
	WpfWindow("窗体列表").WpfTabStrip("TaskList").RefreshObject
	Set children = WpfWindow("窗体列表").WpfTabStrip("TaskList").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		'print(children(i).GetVisibleText)
		If InStr(children(i).GetVisibleText, channelName) = 1 Then
			children(i).DblClick 2, 2
			VideoPreviewRet = True
			Wait(1)
			Exit For
		End If
	Next
	Set children = nothing
	If not VideoPreviewRet Then
		Reporter.ReportEvent micFail, "视频预览", "无所要通道"
		ExitAction(0)
	End If
End Function

'推送告警
Function AlarmPushed()
	Wait(60)
	AlarmPushedRet = False
	'选择未推出告警
	Set children = WpfWindow("窗体列表").WpfList("alertList").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("text") = "未复核" Then
			children(i).Click
			AlarmPushedRet = True
			Exit For
		End If
	Next
	If not AlarmPushedRet Then
		Reporter.ReportEvent micFail, "推送告警", "无告警输出"
		ExitAction(0)
	End If
	AlarmPushedRet = False
	Wait(1)
	'推送告警
	WpfWindow("窗体列表").WpfButton("确认此目标").Click
	If WpfWindow("提示").Exist(1) Then
		WpfWindow("提示").WpfButton("确定").Click
		AlarmPushedRet = True
	End If
	AlarmPushedRet = False
	Wait(1)
	WpfWindow("窗体列表").WpfButton("推送").Click
	If WpfWindow("popwindow").Exist(1) Then
		WpfWindow("popwindow").WpfButton("推送").Click
		AlarmPushedRet = True
	End If
	WpfWindow("popwindow").Close
End Function

'修改布控任务
Function ModifyTask(taskName, regionsArr, storesArr, strategyName)
	ModifyTaskRet = False
	
	'选择布控任务
	Set children = WpfWindow("窗体列表").WpfList("taskList").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("text") = CStr(taskName) Then
			children(i).Click
			ModifyTaskRet = True
			Exit For
		End If
	Next
	Set children = Nothing
	If not ModifyTaskRet Then
		Reporter.ReportEvent micFail, "首页修改布控任务", "无所要布控任务"
		ExitAction(0)
	End If
	ModifyTaskRet = False
	
	WpfWindow("窗体列表").WpfButton("修改").Click
	If not WpfWindow("布控任务").Exist(1) Then
		Reporter.ReportEvent micFail, "首页修改布控任务", "无布控任务列表"
		ExitAction(0)
	End If
	
	'修改目标人
	WpfWindow("布控任务").WpfButton("选择目标人").Click
	If not WpfWindow("选择目标人").Exist(1) Then
		Reporter.ReportEvent micFail, "首页修改布控任务", "无目标人列表"
		ExitAction(0)
	End If
	WpfWindow("选择目标人").WpfLink("取消全选").Click
	Set children = WpfWindow("选择目标人").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		'print(children(i).GetVisibleText)
		For arrIndex = LBound(storesArr) To UBound(storesArr) Step 1
			If children(i).GetVisibleText = storesArr(arrIndex) Then
				children(i).Click
				ModifyTaskRet = True
			End If
		Next
	Next
	If not ModifyTaskRet Then
		Reporter.ReportEvent micFail, "首页修改布控任务", "无所要目标人"
		ExitAction(0)
	End If
	Set children = Nothing
	ModifyTaskRet = False
	WpfWindow("选择目标人").WpfButton("确认").Click
	Wait(1)
	
	'修改区域
	WpfWindow("布控任务").WpfButton("选择通道").Click
	If not WpfWindow("区域通道").Exist(1) Then
		Reporter.ReportEvent micFail, "首页修改布控任务", "无区域通道列表"
		ExitAction(0)
	End If
	WpfWindow("区域通道").WpfLink("取消全选").Click
	Set children = WpfWindow("区域通道").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		'print(children(i).GetVisibleText)
		For arrIndex = LBound(regionsArr) To UBound(regionsArr) Step 1
			If children(i).GetVisibleText = regionsArr(arrIndex) Then
				children(i).Click
				ModifyTaskRet = True
			End If
		Next
	Next
	If not ModifyTaskRet Then
		Reporter.ReportEvent micFail, "首页修改布控任务", "无所要区域"
		ExitAction(0)
	End If
	Set children = Nothing
	ModifyTaskRet = False
	WpfWindow("区域通道").WpfButton("确定").Click
	Wait(1)
	
	'选择比对策略
	count = WpfWindow("布控任务").WpfComboBox("cbCmpList").GetItemsCount
	For i = 0 To count - 1 Step 1
		WpfWindow("布控任务").WpfComboBox("cbCmpList").Select(i)
		If WpfWindow("布控任务").WpfComboBox("cbCmpList").GetROProperty("selection") = strategyName Then
			ModifyTaskRet = True
			Exit For
		End If
	Next
	If not ModifyTaskRet Then
		Reporter.ReportEvent micFail, "首页修改布控任务", "无要比对策略"
		ExitAction(0)
	End If
	Set children = Nothing
	WpfWindow("布控任务").WpfButton("保存").Click
	Wait(1)
	WpfWindow("窗体列表").WpfButton("告警订阅").Click
	If WpfWindow("提示").Exist(1) Then
		WpfWindow("提示").WpfButton("确定").Click
	End If
	Wait(1)
	WpfWindow("窗体列表").WpfButton("告警订阅").Click
	If WpfWindow("提示").Exist(1) Then
		WpfWindow("提示").WpfButton("确定").Click
	End If
	Wait(1)
End Function


'更改比对策略
Function ModifyStrategy(taskName, strategyType, threshold, countTotal, thresholdCount, countHit)
	ModifyStrategyRet = False
	
	'选择布控任务
	Set children = WpfWindow("窗体列表").WpfList("taskList").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("text") = CStr(taskName) Then
			children(i).Click
			ModifyStrategyRet = True
			Exit For
		End If
	Next
	Set children = Nothing
	If not ModifyStrategyRet Then
		Reporter.ReportEvent micFail, "首页修改布控任务", "无所要布控任务"
		ExitAction(0)
	End If
	ModifyStrategyRet = False
	
	WpfWindow("窗体列表").WpfButton("调整策略参数").Click
	If not WpfWindow("比对策略").Exist(1) Then
		Reporter.ReportEvent micFail, "首页修改比对策略", "无比对策略列表"
		ExitAction(1)
	End If
	WpfWindow("比对策略").WpfButton("选择比对方法").Click
	If not WpfWindow("比对策略").Exist(1) Then
		Reporter.ReportEvent micFail, "首页修改比对策略", "无选择比对方法列表"
		ExitAction(1)
	End If
	WpfWindow("选择比对方法 （拖动更改项顺序）").WpfLink("全选").Click
	WpfWindow("选择比对方法 （拖动更改项顺序）").WpfButton("确认").Click
	Wait(1)
	
	If strategyType = "threshold" Then
		WpfWindow("比对策略").WpfButton("countClose").Click
		WpfWindow("比对策略").WpfEdit("txtScore").Set(CStr(threshold))
	ElseIf strategyType = "count" Then
		WpfWindow("比对策略").WpfButton("thresholdClose").Click
		WpfWindow("比对策略").WpfEdit("countTotal").Set(CStr(countTotal))
		WpfWindow("比对策略").WpfEdit("Score").Set(CStr(thresholdCount))
		WpfWindow("比对策略").WpfEdit("countCompare").Set(CStr(countHit))
	ElseIf strategyType = "both" Then
		WpfWindow("比对策略").WpfEdit("txtScore").Set(CStr(threshold))
		WpfWindow("比对策略").WpfEdit("countTotal").Set(CStr(countTotal))
		WpfWindow("比对策略").WpfEdit("Score").Set(CStr(thresholdCount))
		WpfWindow("比对策略").WpfEdit("countCompare").Set(CStr(countHit))
	End If
	WpfWindow("比对策略").WpfButton("保存").Click
	ModifyStrategyRet = True
	Wait(1)
	
	WpfWindow("窗体列表").WpfButton("告警订阅").Click
	If WpfWindow("提示").Exist(1) Then
		WpfWindow("提示").WpfButton("确定").Click
	End If
	Wait(1)
	WpfWindow("窗体列表").WpfButton("告警订阅").Click
	If WpfWindow("提示").Exist(1) Then
		WpfWindow("提示").WpfButton("确定").Click
	End If
	Wait(1)
End Function


'抓取异常
On Error Resume Next
'On Error Goto 0

WpfWindow("窗体列表").Maximiz
WpfWindow("窗体列表").WpfObject("首页").Click


Select Case Parameter("handle")
	Case "taskDetails"
		TaskDetailsRet = False
		Call TaskDetails("CustomTask")
		If TaskDetailsRet Then
			Reporter.ReportEvent micPass, "布控详情", "查看布控详情成功"
			ExitAction(1)
		Else
			Reporter.ReportEvent micFail, "布控详情", "查看布控详情失败"
			ExitAction(0)
		End If
	
	Case "videoPreview"
		VideoPreviewRet = False
		Call VideoPreview("test111", "VideoChannel")
		If VideoPreviewRet Then
			Reporter.ReportEvent micPass, "视频预览", "视频预览成功"
			ExitAction(1)
		Else
			Reporter.ReportEvent micFail, "视频预览", "视频预览失败"
			ExitAction(0)
		End If
		
	Case "alarmPushed"
		AlarmPushedRet = False
		Call AlarmPushed()
		If AlarmPushedRet Then
			Reporter.ReportEvent micPass, "告警推送", "告警推送成功"
			ExitAction(1)
		Else
			Reporter.ReportEvent micFail, "告警推送", "告警推送失败"
			ExitAction(0)
		End If
	
	Case "modifyTaskPage"
		countExcute = ReadExcel(3, 5, "首页修改布控任务", excelPath)
		For caseIndex = 0 To countExcute - 1 Step 1
			ModifyTaskRet = False
			taskName = ReadExcel(3 + caseIndex, 1, "首页修改布控任务", excelPath)
			regionsArr = Split(ReadExcel(3 + caseIndex, 2, "首页修改布控任务", excelPath), ",")
			storesArr = Split(ReadExcel(3 + caseIndex, 3, "首页修改布控任务", excelPath), ",")
			strategyName = ReadExcel(3 + caseIndex, 4, "首页修改布控任务", excelPath)
			Call ModifyTask(taskName, regionsArr, storesArr, strategyName)
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
		
	
	Case "modifyStrategyPage"
		countExcute = ReadExcel(3, 7, "首页修改比对策略", excelPath)
		For caseIndex = 0 To countExcute - 1 Step 1
			ModifyStrategyRet = False
			taskName = ReadExcel(3 + caseIndex, 1, "首页修改比对策略", excelPath)
			strategyType = ReadExcel(3 + caseIndex, 2, "首页修改比对策略", excelPath)
			threshold = ReadExcel(3 + caseIndex, 3, "首页修改比对策略", excelPath)
			countTotal = ReadExcel(3 + caseIndex, 4, "首页修改比对策略", excelPath)
			thresholdCount = ReadExcel(3 + caseIndex, 5, "首页修改比对策略", excelPath)
			countHit = ReadExcel(3 + caseIndex, 6, "首页修改比对策略", excelPath)
			Call ModifyStrategy(taskName, strategyType, threshold, countTotal, thresholdCount, countHit)
			If ModifyStrategyRet Then
				Reporter.ReportEvent micPass, "修改比对策略", "修改比对策略-" & taskName & "-成功"
			Else
				Reporter.ReportEvent micFail, "修改比对策略", "修改比对策略-" & taskName & "-失败"
			End If
		Next
		If ModifyStrategyRet Then
			ExitAction(1)
		Else
			ExitAction(0)
		End If
		

	Case Else
		Reporter.ReportEvent micFail, "首页", "输入首页操作类型错误"
		ExitAction(0)
End Select

If Err.Number <> 0 Then
	errMessage = "错误代码： " & CStr(Err.Number) & ", 错误信息： " & Err.Description & "."
	print(errMessage)
	ExitAction(-1)
End If