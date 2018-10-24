'调用鼠标键盘操作函数
'KeyBoard通过键盘输入数据
'DelExist清除以存在数据
'RightKey右键点击
ExecuteFile("D:\MouseKeyboard.vbs")

'新建目标库
Function NewStore(storeName)	
	NewStore = False
	WpfWindow("窗体列表").WpfButton("新建").Click
	Wait(1)
	WpfWindow("窗体列表").InsightObject("storeName").Click
	DelExist()
	KeyBoard(storeName) @@ hightlight id_;_329184_;_script infofile_;_ZIP::ssf3.xml_;_
	WpfWindow("窗体列表").InsightObject("storeCount").Click
	DelExist()
	KeyBoard("10000")
	WpfWindow("窗体列表").WpfButton("保存").Click
	Wait(1)
	If WpfWindow("提示").WpfEdit("txtMessage").GetROProperty("text") = "确认要新建目标库吗？" Then
		WpfWindow("提示").WpfButton("是").Click
		NewStore = True
	Else
		NewStore = False
	End If
End Function



'抓取异常
On Error Resume Next
'最大化窗口
WpfWindow("窗体列表").Maximize
WpfWindow("窗体列表").WpfObject("目标库").Click

Dim retTmp 
retTmp = False
'Call NewTemplate("111", "111")


'调用新建目标库函数
retTmp = NewStore("test111")
If retTmp Then
	Reporter.ReportEvent micPass, "新建目标库", "新建目标人成功"
	WpfWindow("窗体列表").RefreshObject
	Wait(3)
	'Call NewTemplate("test111", "test111", "D:\1.jpg")
	'RunAction("新建目标人", oneIteration)
	retTmp = RunAction("新建目标人", oneIteration)
	print("新建目标库执行结果：" & CStr(retTmp))
	ExitAction(1)
Else
	Reporter.ReportEvent micFail, "新建目标库", "新建目标人失败"
	ExitAction(0)
End If

If Err.Number <> 0 Then
	errMessage = "错误代码： " & CStr(Err.Number) & ", 错误信息： " & Err.Description & "."
	print(Err.Number)
	print(Err.Source)
	print(Err.Description)
	ExitAction(-1)
End If