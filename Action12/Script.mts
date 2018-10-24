'调用鼠标键盘操作函数
'KeyBoard通过键盘输入数据
'DelExist清除以存在数据
'RightKey右键点击
vbsPath = DataTable("A", dtGlobalSheet)
excelPath = DataTable("B", dtGlobalSheet)
ExecuteFile(vbsPath)

'新建视频平台通道
Function NewVideoChannel(channelName, channelNo, channelAddr, channelPort, channelUser, channelPwd)
	NewChannelRet = False
	' 选择最后一个区域，添加视频平台通道
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(count - 1 - i).GetROProperty("helptext") = "GridViewCell" Then
			children(count - 1 - i).Click
			children(count - 1 - i).Click 0, 0, 1
			RightKey(5)
			NewChannelRet = True
			Exit For
		End If
	Next
	Set children = nothing
	If not NewChannelRet Then
		Reporter.ReportEvent micFail, "新建视频平台通道", "无区域"
		ExitAction(0)
	End If
	NewChannelRet = False
	
	WpfWindow("窗体列表").RefreshObject
	'Wait(1)
	'新建视频平台通道
	If WpfWindow("窗体列表").WpfEdit("TbChannelName").Exist(1) Then
		WpfWindow("窗体列表").WpfEdit("TbChannelName").Set(channelName)
	Else
		Reporter.ReportEvent micFail, "新建视频平台通道", "无添加视频平台通道界面"
		ExitAction(0)
	End If

	WpfWindow("窗体列表").InsightObject("ChannelNo").Click
	DelExist()
	KeyBoard(CStr(channelNo))
	
	WpfWindow("窗体列表").InsightObject("VideoAddr").Click
	DelExist()
	KeyBoard(CStr(channelAddr))
	
	WpfWindow("窗体列表").InsightObject("VideoPort").Click
	DelExist()
	KeyBoard(CStr(channelPort))

	WpfWindow("窗体列表").InsightObject("VideoUserName").Click
	DelExist()
	KeyBoard(CStr(channelUser))
	
	WpfWindow("窗体列表").InsightObject("VideoPassword").Click
	DelExist()
	KeyBoard(CStr(channelPwd))
	
	WpfWindow("窗体列表").WpfButton("保存").Click
	Wait(1)
	If WpfWindow("提示").WpfEdit("txtMessage").GetROProperty("text") = "保存通道成功！" Then
		WpfWindow("提示").WpfButton("确定").Click
		NewChannelRet = True
	End If

End Function

'新建RTSP视频通道
Function NewRTSPChannel(channelName, channelNo)
	NewChannelRet = False
	' 选择最后一个区域，添加视频平台通道
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(count - 1 - i).GetROProperty("helptext") = "GridViewCell" Then
			children(count - 1 - i).Click
			children(count - 1 - i).Click 0, 0, 1
			RightKey(5)
			NewChannelRet = True
			Exit For
		End If
	Next
	Set children = nothing
	If not NewChannelRet Then
		Reporter.ReportEvent micFail, "新建视频平台通道", "无区域"
		ExitAction(0)
	End If
	NewChannelRet = False
	
	WpfWindow("窗体列表").RefreshObject
	
	'新建RTSP视频通道，将新建通道选择改为RTSP视频
	If WpfWindow("窗体列表").WpfEdit("TbChannelName").Exist(1) Then
		WpfWindow("窗体列表").WpfEdit("TbChannelName").Set(channelName)
	Else
		Reporter.ReportEvent micFail, "新建RTSP视频通道", "无添加视频平台通道界面"
		ExitRun
	End If
	Set children = WpfWindow("窗体列表").WpfObject("通道详细信息维护").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("helptext") = "RadComboBox" Then
			countItem = children(i).GetItemsCount
			For j = 0 To countItem - 1 Step 1
				children(i).Select(j)
				If children(i).GetROProperty("selection") = "RTSP协议" Then
					NewChannelRet = True
					Exit For
				End If
			Next
			If NewChannelRet Then
				NewChannelRet = False
				Exit For
			End If
		End If
	Next
	
	Wait(1)
	WpfWindow("窗体列表").InsightObject("RTSPAddr").Click
	DelExist()
	KeyBoard(channelNo)
	
	WpfWindow("窗体列表").WpfButton("保存").Click
	Wait(1)
	If WpfWindow("提示").WpfEdit("txtMessage").GetROProperty("text") = "保存通道成功！" Then
		WpfWindow("提示").WpfButton("确定").Click
		NewChannelRet = True
	End If
End Function

'新建GB28181视频通道
Function NewGBChannel(channelName, channelNo, channelAddr, channelPort)
	NewChannelRet = False
	' 选择最后一个区域，添加视频平台通道
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(count - 1 - i).GetROProperty("helptext") = "GridViewCell" Then
			children(count - 1 - i).Click
			children(count - 1 - i).Click 0, 0, 1
			RightKey(5)
			NewChannelRet = True
			Exit For
		End If
	Next
	Set children = nothing
	If not NewChannelRet Then
		Reporter.ReportEvent micFail, "新建视频平台通道", "无区域"
		ExitRun
	End If
	NewChannelRet = False
	
	WpfWindow("窗体列表").RefreshObject
	
	'新建GB视频通道，将新建通道选择改为GB视频
	If WpfWindow("窗体列表").WpfEdit("TbChannelName").Exist(1) Then
		WpfWindow("窗体列表").WpfEdit("TbChannelName").Set(channelName)
	Else
		Reporter.ReportEvent micFail, "新建GB视频通道", "无添加视频平台通道界面"
		ExitAction(0)
	End If
	Set children = WpfWindow("窗体列表").WpfObject("通道详细信息维护").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("helptext") = "RadComboBox" Then
			countItem = children(i).GetItemsCount
			For j = 0 To countItem - 1 Step 1
				children(i).Select(j)
				If children(i).GetROProperty("selection") = "GB/T 28181（2011）" Then
					NewChannelRet = True
					Exit For
				End If
			Next
			If NewChannelRet Then
				NewChannelRet = False
				Exit For
			End If
		End If
	Next
	
	Wait(1)
	WpfWindow("窗体列表").InsightObject("ChannelCode").Click
	DelExist()
	KeyBoard(channelNo)
	
	WpfWindow("窗体列表").InsightObject("CommonCode").Click
	DelExist()
	KeyBoard(channelNo)
	
	WpfWindow("窗体列表").InsightObject("VideoAddr").Click
	DelExist()
	KeyBoard(CStr(channelAddr))

	WpfWindow("窗体列表").InsightObject("VideoPort").Click
	DelExist()
	KeyBoard(CStr(channelPort))
	
	WpfWindow("窗体列表").WpfButton("保存").Click
	Wait(1)
	If WpfWindow("提示").WpfEdit("txtMessage").GetROProperty("text") = "保存通道成功！" Then
		WpfWindow("提示").WpfButton("确定").Click
		NewChannelRet = True
	End If
End Function

'新建离线视频通道
Function NewFileChannel(channelName, channelNo)
	NewChannelRet = False
	' 选择最后一个区域，添加视频平台通道
	Set children = WpfWindow("窗体列表").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(count - 1 - i).GetROProperty("helptext") = "GridViewCell" Then
			children(count - 1 - i).Click
			children(count - 1 - i).Click 0, 0, 1
			RightKey(5)
			NewChannelRet = True
			Exit For
		End If
	Next
	Set children = nothing
	If not NewChannelRet Then
		Reporter.ReportEvent micFail, "新建视频平台通道", "无区域"
		ExitAction(0)
	End If
	NewChannelRet = False
	
	WpfWindow("窗体列表").RefreshObject
	
	'新建离线视频通道，将新建通道选择改为离线视频
	If WpfWindow("窗体列表").WpfEdit("TbChannelName").Exist(1) Then
		WpfWindow("窗体列表").WpfEdit("TbChannelName").Set(channelName)
	Else
		Reporter.ReportEvent micFail, "新建离线视频通道", "无添加视频平台通道界面"
		ExitRun
	End If
	Set children = WpfWindow("窗体列表").WpfObject("通道详细信息维护").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		If children(i).GetROProperty("helptext") = "RadComboBox" Then
			countItem = children(i).GetItemsCount
			For j = 0 To countItem - 1 Step 1
				children(i).Select(j)
				If children(i).GetROProperty("selection") = "离线视频文件（测试版）" Then
					NewChannelRet = True
					Exit For
				End If
			Next
			If NewChannelRet Then
				NewChannelRet = False
				Exit For
			End If
		End If
	Next
	
	Wait(1)
	WpfWindow("窗体列表").InsightObject("FileAddr").Click
	DelExist()
	KeyBoard(channelNo)
	
	WpfWindow("窗体列表").WpfButton("保存").Click
	Wait(1)
	If WpfWindow("提示").WpfEdit("txtMessage").GetROProperty("text") = "保存通道成功！" Then
		WpfWindow("提示").WpfButton("确定").Click
		NewChannelRet = True
	Else
		WpfWindow("提示").Close
	End If
End Function


'抓取异常
On Error Resume Next
'On Error Goto 0

WpfWindow("窗体列表").Maximize


countExcute = ReadExcel(3, 8, "新建通道", excelPath)
'print(countExcute)
print(countExcute - 1)
For caseIndex = 0 To Int(countExcute) - 1 Step 1
	WpfWindow("窗体列表").WpfObject("区域通道").Click
	WpfWindow("窗体列表").WpfObject("区域通道").RefreshObject
	newChannelRet = False
	channelType = ReadExcel(3 + caseIndex, 1, "新建通道", excelPath)
	channelName = ReadExcel(3 + caseIndex, 2, "新建通道", excelPath)
	channelNo = ReadExcel(3 + caseIndex, 3, "新建通道", excelPath)
	channelAddr = ReadExcel(3 + caseIndex, 4, "新建通道", excelPath)
	channelPort = ReadExcel(3 + caseIndex, 5, "新建通道", excelPath)
	channelUser = ReadExcel(3 + caseIndex, 6, "新建通道", excelPath)
	channelPwd = ReadExcel(3 + caseIndex, 7, "新建通道", excelPath)
	'print(channelType)
	print("i--->" & CStr(caseIndex))
	Select Case channelType
		Case "Video"
			Call NewVideoChannel(channelName, channelNo, channelAddr, channelPort, channelUser, channelPwd)
			If NewChannelRet Then
				Reporter.ReportEvent micPass, "新建视频平台通道", "新建视频平台通道-" & channelName & "-成功"
				'print("新建视频平台通道-" & channelName & "-成功")
			Else
				Reporter.ReportEvent micFail, "新建视频平台通道", "新建视频平台通道-" & channelName & "-失败"
				'print("新建视频平台通道-" & channelName & "-失败")
			End If
		
		Case "RTSP"
			Call NewRTSPChannel(channelName, channelNo)
			If NewChannelRet Then
				Reporter.ReportEvent micPass, "新建RTSP通道", "新建RTSP通道-" & channelName & "-成功"
			Else
				Reporter.ReportEvent micFail, "新建RTSP通道", "新建RTSP通道-" & channelName & "-失败"
			End If
			
		Case "GB28181"
			Call NewGBChannel(channelName, channelNo, channelAddr, channelPort)
			If NewChannelRet Then
				Reporter.ReportEvent micPass, "新建国标通道", "新建国标通道-" & channelName & "-成功"
			Else
				Reporter.ReportEvent micFail, "新建国标通道", "新建国标通道-" & channelName & "-失败"
			End If
		
		Case "File"
			Call NewFileChannel(channelName, channelNo)
			If NewChannelRet Then
				Reporter.ReportEvent micPass, "新建离线视频通道", "新建离线视频通道-" & channelName & "-成功"
			Else
				Reporter.ReportEvent micFail, "新建离线视频通道", "新建离线视频通道-" & channelName & "-失败"
			End If
		
		Case Else
			Reporter.ReportEvent micPass, "新建通道", "通道类型错误"
			ExitAction(0)
	End Select
Next

If NewChannelRet Then
	ExitAction(1)
Else
	ExitAction(0)
End If

'Select Case Parameter("ChannelType")
'	
'	Case "Video" 
'		NewVideoChannelRet = False
'		Call NewVideoChannel(channelName, channelNo)
'		If NewVideoChannelRet Then
'			Reporter.ReportEvent micPass, "新建视频平台通道", "新建视频平台通道成功"
'			ExitAction(1)
'		Else
'			Reporter.ReportEvent micFail, "新建视频平台通道", "新建视频平台通道失败"
'			ExitAction(0)
'		End If
'	
'	Case "RTSP"
'		NewRTSPChannelRet = False
'		Call NewRTSPChannel(channelName, channelNo)
'		If NewRTSPChannelRet Then
'			Reporter.ReportEvent micPass, "新建RTSP视频通道", "新建RTSP视频通道成功"
'			ExitAction(1)
'		Else
'			Reporter.ReportEvent micFail, "新建RTSP视频通道", "新建RTSP视频通道失败"
'			ExitAction(0)
'		End If
'		
'	Case "GB28181"
'		NewGBChannelret = False
'		Call NewGBChannel(channelName, channelNo)
'		If NewGBChannelret Then
'			Reporter.ReportEvent micPass, "新建GB28181视频通道", "新建GB28181视频通道成功"
'			ExitAction(1)
'		Else
'			Reporter.ReportEvent micFail, "新建GB28181视频通道", "新建GB28181视频通道失败"
'			ExitAction(0)
'		End If
'	
'	Case "File"
'		NewFileChannelret = False
'		Call NewFileChannel(channelName, channelNo)
'		If NewFileChannelret Then
'			Reporter.ReportEvent micPass, "新建离线视频通道", "新建离线视频通道成功"
'			ExitAction(1)
'		Else
'			Reporter.ReportEvent micFail, "新建离线视频通道", "新建离线视频通道失败"
'			ExitAction(0)
'		End If
'		
'	Case else
'		Reporter.ReportEvent micFail, "新建通道", "输入新建的通道类型错误"
'		ExitAction(0)
'End Select

If Err.Number <> 0 Then
	errMessage = "错误代码： " & CStr(Err.Number) & ", 错误信息： " & Err.Description & "."
	print(errMessage)
	ExitAction(-1)
End If