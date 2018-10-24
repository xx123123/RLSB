'调用鼠标键盘操作函数
'KeyBoard通过键盘输入数据
'DelExist清除以存在数据
'RightKey右键点击
vbsPath = DataTable("A", dtGlobalSheet)
excelPath = DataTable("B", dtGlobalSheet)
ExecuteFile(vbsPath)

'查询选择项抓拍数据
Function CaptureChosen(dateTime, regionsArr)
	CaptureChosenRet = False
	WpfWindow("窗体列表").WpfButton("清空查询").Click
	'选择查询时间
	Select Case dateTime
		Case -1
			WpfWindow("窗体列表").WpfLink("昨天").Click
		Case 0
			WpfWindow("窗体列表").WpfLink("今天").Click
		Case 7
			WpfWindow("窗体列表").WpfLink("7天").Click
		Case Else
			Reporter.ReportEvent micFail, "抓拍历史", "输入时间信息错误"
			ExitAction(0)
	End Select
	
	'选择通道
	WpfWindow("窗体列表").WpfButton("选择区域").Click
	If not WpfWindow("区域通道").Exist(3) Then
		Reporter.ReportEvent micFail, "抓拍历史", "无通道列表"
		ExitAction(0)
	End If
	WpfWindow("区域通道").WpfLink("取消全选").Click
	Set children = WpfWindow("区域通道").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	'print("count:" & CStr(count))
	For i = 0 To count - 1 Step 1
		For arrIndex = LBound(regionsArr) To UBound(regionsArr) Step 1
			print(regionsArr(arrIndex))
			If children(i).GetVisibleText = regionsArr(arrIndex) Then
				children(i).Click
				CaptureChosenRet = True
			End If
		Next
	Next
	If not CaptureChosenRet Then
		Reporter.ReportEvent micFail, "抓拍历史", "无所要通道"
		ExitAction(0)
	End If
	'CaptureChosenRet = False
	WpfWindow("区域通道").WpfButton("确定").Click
	Wait(1)
	
	WpfWindow("窗体列表").WpfButton("查询").Click
	Wait(10)
End Function

'抓取异常
On Error Resume Next @@ hightlight id_;_2051838304_;_script infofile_;_ZIP::ssf2.xml_;_
WpfWindow("窗体列表").Maximize
WpfWindow("窗体列表").WpfObject("抓拍历史").Click

Select Case Parameter("handle")
	Case "captureChosen"
		countExcute = ReadExcel(3, 3, "抓拍历史", excelPath)
		For caseIndex = 0 To countExcute - 1 Step 1
			CaptureChosenRet = False
			dateTime = ReadExcel(3 + caseIndex, 1, "抓拍历史", excelPath)
			regionsArr = Split(ReadExcel(3 + caseIndex, 2, "抓拍历史", excelPath), ",")
			Call CaptureChosen(dateTime, regionsArr)
			If CaptureChosenRet Then
				Reporter.ReportEvent micPass, "抓拍历史", "查询选择项抓拍数据成功"
			Else
				Reporter.ReportEvent micFail, "抓拍历史", "查询选择项抓拍数据失败"
			End If
		Next
		If CaptureChosenRet Then
			ExitAction(1)
		Else
			ExitAction(0)
		End If
		
End Select

If Err.Number <> 0 Then
	errMessage = "错误代码： " & CStr(Err.Number) & ", 错误信息： " & Err.Description & "."
	print(errMessage)
	ExitAction(-1)
End If