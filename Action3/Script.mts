'调用鼠标键盘操作函数
'KeyBoard通过键盘输入数据
'delExist清除以存在数据
'RightKey右键点击
ExecuteFile("D:\MouseKeyboard.vbs")

'新建目标人
Function NewTemplate(storeName, templateName, imageUrl)
	NewTemplate = False
	'选择所要添加的目标库
	countStore = WpfWindow("窗体列表").WpfComboBox("WpfComboBox").GetItemsCount
	For i = 0 To countStore - 1 Step 1
		WpfWindow("窗体列表").WpfComboBox("WpfComboBox").Select(i)
		If WpfWindow("窗体列表").WpfComboBox("WpfComboBox").GetROProperty("selection") = CStr(storeName) Then
			'print("1111")
			NewTemplate = True
			Exit For
		End If
	Next
	If not NewTemplate Then
		'print(retTmp)
		Reporter.ReportEvent micFail, "新建目标人", "所选目标库不存在"
		ExtiAction(0)
	End If
	NewTemplate = False
	
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
		WpfWindow("窗体列表").Dialog("请选择模板照片").WinEdit("文件名(N):").Set(imageUrl)
		WpfWindow("窗体列表").Dialog("请选择模板照片").WinButton("打开(&O)").Click
	End If
	Wait(1)
	WpfWindow("窗体列表").WpfButton("btnSave").Click
	If WpfWindow("提示").Exist(1) Then
		WpfWindow("提示").WpfButton("是").Click
	End If
	WpfWindow("提示").WpfButton("确定").Click
	NewTemplate = True
	
End Function

'抓取异常
On Error Resume Next
WpfWindow("窗体列表").Maximize
WpfWindow("窗体列表").WpfObject("目标库").Click

'调用新建目标人函数
Call NewTemplate("test111", "template111", "D:\2.jpg")
If NewTemplate Then
	Reporter.ReportEvent micPass, "新建目标人", "新建目标库成功"
	ExitAction(1)
Else
	Reporter.ReportEvent micFail, "新建目标人", "新建目标库失败"
	ExitAction(0)
End If

If Err.Number <> 0 Then
	errMessage = "错误代码： " & CStr(Err.Number) & ", 错误信息： " & Err.Description & "."
	print(Err.Number)
	print(Err.Source)
	print(Err.Description)
	ExitAction(-1)
End If
