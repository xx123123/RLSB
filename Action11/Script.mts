'调用鼠标键盘操作函数
'KeyBoard通过键盘输入数据
'DelExist清除以存在数据
'RightKey右键点击
vbsPath = DataTable("A", dtGlobalSheet)
excelPath = DataTable("B", dtGlobalSheet)
ExecuteFile(vbsPath)

'新建阈值法策略
Function NewThresholdStrategy(strategyName, threshold)
	NewStrategyRet = False
	WpfWindow("窗体列表").WpfButton("+新建策略").Click
	If not WpfWindow("比对策略").Exist(1) Then
		ExitAction(0)
	End If
	
	WpfWindow("比对策略").WpfEdit("策略名称").Set(CStr(strategyName))
	WpfWindow("比对策略").WpfEdit("策略描述").Set("ThresholdTest")
	
	WpfWindow("比对策略").WpfEdit("txtScore").Set(CStr(threshold))
	WpfWindow("比对策略").WpfButton("countClose").Click
	If WpfWindow("比对策略").WpfButton("保存").Exist(1) Then
		WpfWindow("比对策略").WpfButton("保存").Click
	End If
	Wait(1)
	
	WpfWindow("窗体列表").WpfButton("刷新").Click
	WpfWindow("窗体列表").WpfEdit("txtStrategyKeyWord").Click
	DelExist()
	KeyBoard(CStr(strategyName))
	KeyBoard("{ENTER}")
	Set children = WpfWindow("窗体列表").WpfTabStrip("WpfTabStrip").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("text") = CStr(strategyName) Then
			NewStrategyRet = True
			Exit For
		End If
	Next

End Function

'新建计数法策略
Function NewCountStrategy(strategyName, countTotal, thresholdCmp, countHit)
	NewStrategyRet = False
	WpfWindow("窗体列表").WpfButton("+新建策略").Click
	If not WpfWindow("比对策略").Exist(1) Then
		ExitAction(0)
	End If
	
	WpfWindow("比对策略").WpfEdit("策略名称").Set(CStr(strategyName))
	WpfWindow("比对策略").WpfEdit("策略描述").Set("CountTest")
	
	WpfWindow("比对策略").WpfButton("thresholdClose").Click
	WpfWindow("比对策略").WpfEdit("countTotal").Set(CStr(countTotal))
	WpfWindow("比对策略").WpfEdit("Score").Set(CStr(thresholdCmp))
	WpfWindow("比对策略").WpfEdit("countCompare").Set(CStr(countHit))
	If WpfWindow("比对策略").WpfButton("保存").Exist(1) Then
		WpfWindow("比对策略").WpfButton("保存").Click
	End If
	Wait(1)

	WpfWindow("窗体列表").WpfButton("刷新").Click
	WpfWindow("窗体列表").WpfEdit("txtStrategyKeyWord").Click
	DelExist()
	KeyBoard(CStr(strategyName))
	KeyBoard("{ENTER}")
	Set children = WpfWindow("窗体列表").WpfTabStrip("WpfTabStrip").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("text") = CStr(strategyName) Then
			NewStrategyRet = True
			Exit For
		End If
	Next
	
End Function

'新建两种比对方式策略
Function NewBothStrategy(strategyName, threshold, countTotal, thresholdCmp, countHit)
	NewStrategyRet = False
	WpfWindow("窗体列表").WpfButton("+新建策略").Click
	If not WpfWindow("比对策略").Exist(1) Then
		ExitAction(0)
	End If
	
	WpfWindow("比对策略").WpfEdit("策略名称").Set(CStr(strategyName))
	WpfWindow("比对策略").WpfEdit("策略描述").Set("BothTest")
	
	WpfWindow("比对策略").WpfEdit("txtScore").Set(CStr(threshold))
	WpfWindow("比对策略").WpfEdit("countTotal").Set(CStr(countTotal))
	WpfWindow("比对策略").WpfEdit("Score").Set(Cstr(thresholdCmp))
	WpfWindow("比对策略").WpfEdit("countCompare").Set(CStr(countHit))
	If WpfWindow("比对策略").WpfButton("保存").Exist(1) Then
		WpfWindow("比对策略").WpfButton("保存").Click
	End If
	Wait(1)

	WpfWindow("窗体列表").WpfButton("刷新").Click
	WpfWindow("窗体列表").WpfEdit("txtStrategyKeyWord").Click
	DelExist()
	KeyBoard(CStr(strategyName))
	KeyBoard("{ENTER}")
	Wait(1)
	Set children = WpfWindow("窗体列表").WpfTabStrip("WpfTabStrip").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("text") = CStr(strategyName) Then
			NewStrategyRet = True
			Exit For
		End If
	Next
End Function


'抓取异常
On Error Resume Next

WpfWindow("窗体列表").Maximize
WpfWindow("窗体列表").WpfObject("布控任务").Click
WpfWindow("窗体列表").WpfTabStrip("WpfTabStrip").Select(1)

countExcute = ReadExcel(3, 7, "新建比对策略", excelPath)
For caseIndex = 0 To countExcute - 1 Step 1
	strategyType = ReadExcel(3 + caseIndex, 1, "新建比对策略", excelPath)
	strategyName = ReadExcel(3 + caseIndex, 2, "新建比对策略", excelPath)
	threshold = ReadExcel(3 + caseIndex, 3, "新建比对策略", excelPath)
	countTotal = ReadExcel(3 + caseIndex, 4, "新建比对策略", excelPath)
	thresholdCmp = ReadExcel(3 + caseIndex, 5, "新建比对策略", excelPath)
	countHit = ReadExcel(3 + caseIndex, 6, "新建比对策略", excelPath)
	NewStrategyRet = False
	
	Select Case strategyType
		
		Case "Threshold"
			Call NewThresholdStrategy(strategyName, threshold)
			If NewStrategyRet Then
				Reporter.ReportEvent micPass, "新建阈值法策略", "新建阈值法策略-" & strategyName & "-成功"
			Else
				Reporter.ReportEvent micFail, "新建阈值法策略", "新建阈值法策略-" & strategyName & "-失败"
			End If
			
		Case "Count"
			Call NewCountStrategy(strategyName, countTotal, thresholdCmp, countHit)
			If NewStrategyRet Then
				Reporter.ReportEvent micPass, "新建计数法策略", "新建计数法策略-" & strategyName & "-成功"
			Else
				Reporter.ReportEvent micFail, "新建计数法策略", "新建计数法策略-" & strategyName & "-失败"
			End If
			
		Case "Both"
			Call NewBothStrategy(strategyName, threshold, countTotal, thresholdCmp, countHit)
			If NewStrategyRet Then
				Reporter.ReportEvent micPass, "新建双方法策略", "新建双方法策略-" & strategyName & "-成功"
			Else
				Reporter.ReportEvent micFail, "新建双方法策略", "新建双方法策略-" & strategyName & "-失败"
			End If
		
		Case Else
			Reporter.ReportEvent micFail, "新建比对策略", "输入参数错误"
			ExitAction(0)
	End Select
Next
If NewStrategyRet Then
	ExitAction(1)
Else
	ExitAction(0)
End If

'strategyName = Parameter("strategyName")
'Select Case Parameter("StrategyType")
'	
'	Case "Threshold"
'		NewThresholdStrategyRet = False
'		Call NewThresholdStrategy(strategyName)
'		If NewThresholdStrategyRet Then
'			Reporter.ReportEvent micPass, "新建阈值法策略", "新建阈值法策略成功"
'			ExitAction(1)
'		Else
'			Reporter.ReportEvent micFail, "新建阈值法策略", "新建阈值法策略失败"
'			ExitAction(0)
'		End If
'	
'	Case "Count"
'		NewCountStrategyRet = False
'		Call NewCountStrategy(strategyName)
'		If NewCountStrategyRet Then
'			Reporter.ReportEvent micPass, "新建计数法策略", "新建计数法策略成功"
'			ExitAction(1)
'		Else
'			Reporter.ReportEvent micFail, "新建计数法策略", "新建计数法策略失败"
'			ExitAction(0)
'		End If
'	
'	Case "Both"
'		NewBothStrategyRet = False
'		Call NewBothStrategy(strategyName)
'		If NewBothStrategyRet Then
'			Reporter.ReportEvent micPass, "新建双方法策略", "新建双方法策略成功"
'			ExitAction(1)
'		Else
'			Reporter.ReportEvent micFail, "新建双方法策略", "新建双方法策略失败"
'			ExitAction(0)
'		End If
'	
'	Case Else
'		Reporter.ReportEvent micFail, "新建策略", "输入新建的策略类型错误"
'		ExitAction(0)
'End Select
'
''Call NewThresholdStrategy("ThresholdTest", "65.00")
''If NewThresholdStrategy Then
''	Reporter.ReportEvent micPass, "新建阈值法策略", "新建阈值法策略成功"
''	ExitAction(1)
''Else
''	Reporter.ReportEvent micFail, "新建阈值法策略", "新建阈值法策略失败"
''	ExitAction(0)
''End If

If Err.Number <> 0 Then
	errMessage = "错误代码： " & CStr(Err.Number) & ", 错误信息： " & Err.Description & "."
	print(errMessage)
	ExitAction(-1)
End If