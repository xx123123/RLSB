'调用鼠标键盘操作函数
'KeyBoard通过键盘输入数据
'DelExist清除以存在数据
'RightKey右键点击
vbsPath = DataTable("A", dtGlobalSheet)
excelPath = DataTable("B", dtGlobalSheet)
ExecuteFile(vbsPath)

'新建区域
Function NewRegion(regionName)
	NewRegionRet = False
	' 选择最后一个区域，建立同级区域
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(count - 1 - i).GetROProperty("helptext") = "GridViewCell" Then
			children(count - 1 - i).Click
			children(count - 1 - i).Click 0, 0, 1
			RightKey(2)
			NewRegionRet = True
			Exit For
		End If
	Next
	Set children = nothing
	If not NewRegionRet Then
		Reporter.ReportEvent micFail, "新建区域", "无区域"
		ExitAction(0)
	End If
	NewRegionRet = False
	
	WpfWindow("窗体列表").Click
	
	'选择最后一个区域改名
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(count - 1 - i).GetROProperty("helptext") = "GridViewCell" Then
			children(count - 1 - i).Click
			children(count - 1 - i).Click 0, 0, 1
			RightKey(3)
			WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").RefreshObject
			Set child = children(count - 1 - i).ChildObjects
			For j = 0 To child.Count - 1 Step 1
				If child(j).GetROProperty("text") = "区域" Then
					child(j).Click
					DelExist()
					Wait(1)
					KeyBoard(regionName)
					NewRegionRet = True
					Exit For
				End IF
			Next
			Exit For
		End If
	Next
	Set children = nothing
	
	'NewRegions = False
End Function

'抓取异常
'On Error Resume Next
On Error Goto 0

WpfWindow("窗体列表").Maximize
WpfWindow("窗体列表").WpfObject("区域通道").Click

countExcute = ReadExcel(3, 2, "新建区域", excelPath)
For caseIndex = 0 To countExcute - 1 Step 1
	NewRegionRet = False
	regionName = ReadExcel(3 + caseIndex, 1, "新建区域", excelPath)
	Call NewRegion(regionName)
	If NewRegionRet Then
		Reporter.ReportEvent micPass, "新建区域", "新建区域-" & CStr(regionName) & "-成功"
	Else
		Reporter.ReportEvent micFail, "新建区域", "新建区域-" & CStr(regionName) & "-失败"
	End If
Next
If NewRegionRet Then
	ExitAction(1)
Else
	ExitAction(0)
End If


If Err.Number <> 0 Then
	errMessage = "错误代码： " & CStr(Err.Number) & ", 错误信息： " & Err.Description & "."
	print(errMessage)
	ExitAction(-1)
End If