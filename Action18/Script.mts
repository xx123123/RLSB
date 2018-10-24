'调用鼠标键盘操作函数
'KeyBoard通过键盘输入数据
'DelExist清除以存在数据
'RightKey右键点击
vbsPath = DataTable("A", dtGlobalSheet)
excelPath = DataTable("B", dtGlobalSheet)
ExecuteFile(vbsPath)


'新建目标库
Function NewStore(storeName, storeCount)
	NewStoreRet = False
	WpfWindow("窗体列表").WpfButton("新建").Click
	Wait(1)
	WpfWindow("窗体列表").InsightObject("storeName").Click
	DelExist()
	KeyBoard(storeName) @@ hightlight id_;_329184_;_script infofile_;_ZIP::ssf3.xml_;_
	WpfWindow("窗体列表").InsightObject("storeCount").Click
	DelExist()
	KeyBoard(CStr(storeCount))
	WpfWindow("窗体列表").WpfButton("保存").Click
	Wait(1)
	If WpfWindow("提示").WpfEdit("txtMessage").GetROProperty("text") = "确认要新建目标库吗？" Then
		WpfWindow("提示").WpfButton("是").Click
		NewStoreRet = True
	Else
		NewStoreRet = False
	End If
End Function


'修改目标库
Function ModifyStore(storeName, newStoreName, newCount)
	ModifyStoreRet = False
	'选择所要添加的目标库
	countStore = WpfWindow("窗体列表").WpfComboBox("WpfComboBox").GetItemsCount
	For i = 0 To countStore - 1 Step 1
		WpfWindow("窗体列表").WpfComboBox("WpfComboBox").Select(i)
		If WpfWindow("窗体列表").WpfComboBox("WpfComboBox").GetROProperty("selection") = CStr(storeName) Then
			ModifyStoreRet = True
			Exit For
		End If
	Next
	If not ModifyStoreRet Then
		Reporter.ReportEvent micFail, "修改目标库", "所选目标库不存在"
		ExtiAction(0)
	End If
	ModifyStoreRet = False
	
	'修改目标库
	WpfWindow("窗体列表").InsightObject("storeName").Click
	DelExist()
	KeyBoard(newStoreName) @@ hightlight id_;_329184_;_script infofile_;_ZIP::ssf3.xml_;_
	WpfWindow("窗体列表").InsightObject("storeCount").Click
	DelExist()
	KeyBoard(CStr(newCount))
	WpfWindow("窗体列表").WpfButton("保存").Click
	If WpfWindow("提示").WpfButton("是").Exist(1) Then
		WpfWindow("提示").WpfButton("是").Click
	End If
	If WpfWindow("提示").WpfButton("确定").Exist(1) Then
		WpfWindow("提示").WpfButton("确定").Click
	End If
	ModifyStoreRet = True
End Function


'新建目标人
Function NewTemplate(storeName, templateName, imgPath)
	NewTemplateRet = False
	'选择所要添加的目标库
	countStore = WpfWindow("窗体列表").WpfComboBox("WpfComboBox").GetItemsCount
	For i = 0 To countStore - 1 Step 1
		WpfWindow("窗体列表").WpfComboBox("WpfComboBox").Select(i)
		If WpfWindow("窗体列表").WpfComboBox("WpfComboBox").GetROProperty("selection") = CStr(storeName) Then
			NewTemplateRet = True
			Exit For
		End If
	Next
	If not NewTemplateRet Then
		Reporter.ReportEvent micFail, "新建目标人", "所选目标库不存在"
		ExtiAction(0)
	End If
	NewTemplateRet = False
	
	'右键新建目标人
	Set children = WpfWindow("窗体列表").WpfTabStrip("RadTabControl1").ChildObjects
	count = children.Count
	For i = 1 To count - 1 Step 1
		If children(i).GetROProperty("text") = "照片" Then
			children(i).Click 0, 0, 1
			RightKey(1)
			Exit For
		End If
	Next
	'新建目标人
	WpfWindow("窗体列表").WpfEdit("WpfEdit").Set(templateName)
	WpfWindow("窗体列表").WpfButton("WpfButton").Click
	If WpfWindow("窗体列表").Dialog("请选择模板照片").Exist(3) Then
		WpfWindow("窗体列表").Dialog("请选择模板照片").WinEdit("文件名(N):").Set(imgPath)
		WpfWindow("窗体列表").Dialog("请选择模板照片").WinButton("打开(&O)").Click
	End If
	Wait(1)
	WpfWindow("窗体列表").WpfButton("btnSave").Click
	If WpfWindow("提示").Exist(1) Then
		WpfWindow("提示").WpfButton("是").Click
	End If
	WpfWindow("提示").WpfButton("确定").Click
	NewTemplateRet = True
End Function


'修改目标人
Function ModifyTemplate(storeName, templateName, newTemplateName, newImgPath)
	ModifyTemplateRet = False
	'print(storeName)
	'选择所要添加的目标库
	countStore = WpfWindow("窗体列表").WpfComboBox("WpfComboBox").GetItemsCount
	For i = 0 To countStore - 1 Step 1
		WpfWindow("窗体列表").WpfComboBox("WpfComboBox").Select(i)
		'print(WpfWindow("窗体列表").WpfComboBox("WpfComboBox").GetROProperty("selection"))
		If WpfWindow("窗体列表").WpfComboBox("WpfComboBox").GetROProperty("selection") = CStr(storeName) Then
			ModifyTemplateRet = True
			Exit For
		End If
	Next
	If not ModifyTemplateRet Then
		Reporter.ReportEvent micFail, "修改目标人", "所选目标库不存在"
		ExtiAction(0)
	End If
	ModifyTemplateRet = False
	
	'选择要修改的目标人
	Set children = WpfWindow("窗体列表").WpfTabStrip("RadTabControl1").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("value") = CStr(templateName) Then
			children(i).DblClick 2, 2
			ModifyTemplateRet = True
			Exit For
		End If
	Next
	If not ModifyTemplateRet Then
		Reporter.ReportEvent micFail, "修改目标人", "所选目标人不存在"
		ExitAction(0)
	End If
	ModifyTemplateRet = False
	
	Wait(1)
	WpfWindow("窗体列表").WpfEdit("WpfEdit").Set(newTemplateName)
	WpfWindow("窗体列表").WpfButton("WpfButton").Click
	If WpfWindow("窗体列表").Dialog("请选择模板照片").Exist(3) Then
		WpfWindow("窗体列表").Dialog("请选择模板照片").WinEdit("文件名(N):").Set(newImgPath)
		WpfWindow("窗体列表").Dialog("请选择模板照片").WinButton("打开(&O)").Click
	End If
	WpfWindow("窗体列表").WpfButton("btnSave").Click
	If WpfWindow("提示").WpfButton("是").Exist(1) Then
		WpfWindow("提示").WpfButton("是").Click
		ModifyTemplateRet = True
	End If
End Function


'抓取异常
On Error Resume Next
'最大化窗口
WpfWindow("窗体列表").Maximize
WpfWindow("窗体列表").WpfObject("目标库").Click


Select Case Parameter("handle")
	
	Case "newStore"
		countExcute = ReadExcel(3, 3, "新建目标库", excelPath)
		For caseIndex = 0 To countExcute - 1 Step 1
			NewStoreRet = False
			storeName = ReadExcel(3 + caseIndex, 1, "新建目标库", excelPath)
			storeCount = ReadExcel(3 + caseIndex, 2, "新建目标库", excelPath)
			Call NewStore(storeName, storeCount)
			If NewStoreRet Then
				Reporter.ReportEvent micPass, "新建目标库", "新建目标库-" & storeName & "-成功"
				ExitAction(1)
			Else
				Reporter.ReportEvent micFail, "新建目标库", "新建目标库-" & storeName & "失败"
				ExitAction(0)
			End If
		Next
		
	
	Case "modifyStore"
		countExcute = ReadExcel(3, 4, "修改目标库", excelPath)
		For caseIndex = 0 To countExcute - 1 Step 1
			ModifyStoreRet = False
			storeName = ReadExcel(3 + caseIndex, 1, "修改目标库", excelPath)
			newStoreName = ReadExcel(3 + caseIndex, 2, "修改目标库", excelPath)
			newCount = ReadExcel(3 + caseIndex, 3, "修改目标库", excelPath)
			'print("name->" & storeName & ", newName->" & newStoreName)
			Call ModifyStore(storeName, newStoreName, newCount)
			If ModifyStoreRet Then
				Reporter.ReportEvent micPass, "修改目标库", "修改目标库-" & storeName & "-成功"
				'ExitAction(1)
			Else
				Reporter.ReportEvent micFail, "修改目标库", "修改目标库-" & storeName & "-失败"
				'ExitAction(0)
			End If
		Next
		If ModifyStoreRet Then
			ExitAction(1)
		Else
			ExitAction(0)
		End If		
	
	Case "newTemplate"
		countExcute = ReadExcel(3, 4, "新建目标人", excelPath)
		For caseIndex = 0 To countExcute - 1 Step 1
			NewTemplateRet = False
			storeName = ReadExcel(3 + caseIndex, 1, "新建目标人", excelPath)
			templateName = ReadExcel(3 + caseIndex, 2, "新建目标人", excelPath)
			imgPath = ReadExcel(3 + caseIndex, 3, "新建目标人", excelPath)
			Call NewTemplate(storeName, templateName, imgPath)
			If NewTemplateRet Then
				Reporter.ReportEvent micPass, "新建目标人", "新建目标人-" & template & "-成功"
				'ExitAction(1)
			Else
				Reporter.ReportEvent micFail, "新建目标人", "新建目标人-" & template & "-失败"
				'ExitAction(0)
			End If
		Next
		If NewTemplateRet Then
			ExitAction(1)
		Else
			ExitAction(0)
		End If
		
	Case "modifyTemplate"
		countExcute = ReadExcel(3, 5, "修改目标人", excelPath)
		For i = 0 To countExcute - 1 Step 1
			ModifyTemplateRet = False
			storeName = ReadExcel(3 + caseIndex, 1, "修改目标人", excelPath)
			templateName = ReadExcel(3 + caseIndex, 2, "修改目标人", excelPath)
			newTemplateName = ReadExcel(3 + caseIndex, 3, "修改目标人", excelPath)
			newImgPath = ReadExcel(3 + caseIndex, 4, "修改目标人", excelPath)
			Call ModifyTemplate(storeName, templateName, newTemplateName, newImgPath)
			If ModifyTemplateRet Then
				Reporter.ReportEvent micPass, "修改目标人", "修改目标人-" & template & "-成功"
				'ExitAction(1)
			Else
				Reporter.ReportEvent micFail, "修改目标人", "修改目标人-" & template & "-失败"
				'ExitAction(0)
			End If
		Next
		If ModifyTemplateRet Then
			ExitAction(1)
		Else
			ExitAction(0)
		End If
		
	Case Else
		Reporter.ReportEvent micFail, "目标库", "输入操作方法错误"
		ExitAction(0)
End Select


If Err.Number <> 0 Then
	errMessage = "错误代码： " & CStr(Err.Number) & ", 错误信息： " & Err.Description & "."
	print(Err.Number)
	print(Err.Source)
	print(Err.Description)
	ExitAction(-1)
End If