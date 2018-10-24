'调用鼠标键盘操作函数
'KeyBoard通过键盘输入数据
'DelExist清除以存在数据
'RightKey右键点击
vbsPath = DataTable("A", dtGlobalSheet)
excelPath = DataTable("B", dtGlobalSheet)
ExecuteFile(vbsPath)


'目标库查询
Function StoreQuary(imgPath, storesArr, scoreThreshold)
	WpfWindow("窗体列表").WpfTabStrip("WpfTabStrip").Select(0)
	StoreQuaryRet = False
	WpfWindow("窗体列表").WpfButton("清除").Click
	
	'选择目标库
	WpfWindow("窗体列表").WpfButton("选择目标库").Click
	If not WpfWindow("目标库").Exist(3) Then
		Reporter.ReportEvent micFail, "目标库查询", "无目标库列表"
		ExitAction(0)
	End If
	WpfWindow("目标库").WpfLink("取消全选").Click
	Set children = WpfWindow("目标库").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		'print(children(i).GetVisibleText)
		For arrIndex = LBound(storesArr) To UBound(storesArr) Step 1
			'print(storesArr(arrIndex))
			If children(i).GetVisibleText = storesArr(arrIndex) Then
				children(i).Click
				StoreQuaryRet = True
			End If
		Next
	Next
	Set children = Nothing
	If not StoreQuaryRet Then
		Reporter.ReportEvent micFail, "目标库查询", "无所要的通道"
		ExitAction(0)
	End If
	WpfWindow("目标库").WpfButton("确定").Click
	Wait(1)
	StoreQuaryRet = False
	
	'选择所要比对的图片
	WpfWindow("窗体列表").InsightObject("InsightObject").Click
	If not WpfWindow("窗体列表").Dialog("请选择人脸照片").Exist(3) Then
		Reporter.ReportEvent micFail, "目标库查询", "无选择图片列表"
		ExitAction(0)
	End If
	WpfWindow("窗体列表").Dialog("请选择人脸照片").WinEdit("文件名(N):").Set(imgPath)
	WpfWindow("窗体列表").Dialog("请选择人脸照片").WinButton("打开(&O)").Click
	Wait(3)
	
	WpfWindow("窗体列表").InsightObject("thresholdStore").Click
	DelExist()
	KeyBoard(CStr(scoreThreshold))
	
	'查询
	WpfWindow("窗体列表").WpfButton("查询").Click
	If WpfWindow("提示").Exist(5) Then
		StoreQuaryRet = False
		WpfWindow("提示").WpfButton("确定").Click
	Else
		StoreQuaryRet = True
	End If

End Function


'抓拍库查询
Function CaptureQuary(dateTime, imgPath, regionsArr, scoreThreshold)
	WpfWindow("窗体列表").WpfTabStrip("WpfTabStrip").Select(1)
	CaptureQuaryRet = False
	WpfWindow("窗体列表").WpfButton("清除").Click
	Select Case dateTime
		Case 3
			WpfWindow("窗体列表").WpfLink("3天").Click
		Case 7
			WpfWindow("窗体列表").WpfLink("7天").Click
		Case 30
			WpfWindow("窗体列表").WpfLink("30天").Click
		Case Else
			Reporter.ReportEvent micFail, "抓拍库查询", "输入查询日期错误"
			ExitAction(0)
	End Select
	
	'选择通道
	WpfWindow("窗体列表").WpfButton("选择区域").Click
	If not WpfWindow("区域通道").Exist(3)  Then
		Reporter.ReportEvent micFail, "抓拍库查询", "无选择区域通道界面"
		ExitAction(0)
	End If
	WpfWindow("区域通道").WpfLink("取消全选").Click
	Set children = WpfWindow("区域通道").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	
	For i = 0 To count - 1 Step 1
		For arrIndex = LBound(regionsArr) To UBound(regionsArr) Step 1
			If children(i).GetVisibleText = regionsArr(arrIndex) Then			
				children(i).Click
				CaptureQuaryRet = True
			End If
		Next
	Next
	Set children = Nothing
	If not CaptureQuaryRet Then
		Reporter.ReportEvent micFail, "抓拍库查询", "无所要的通道"
		ExitAction(0)
	End If
	WpfWindow("区域通道").WpfButton("确定").Click
	CaptureQuaryRet = False
	Wait(1)
	
	WpfWindow("窗体列表").InsightObject("thresholdCapture").Click
	DelExist()
	KeyBoard(CStr(scoreThreshold))


	'选择所要比对的图片
	WpfWindow("窗体列表").InsightObject("InsightObject").Click
	If not WpfWindow("窗体列表").Dialog("请选择人脸照片").Exist(3) Then
		Reporter.ReportEvent micFail, "抓拍库查询", "无选择图片列表"
		ExitAction(0)
	End If
	WpfWindow("窗体列表").Dialog("请选择人脸照片").WinEdit("文件名(N):").Set(imgPath)
	WpfWindow("窗体列表").Dialog("请选择人脸照片").WinButton("打开(&O)").Click
	Wait(1)
	
	'查询
	WpfWindow("窗体列表").WpfButton("查询").Click
	If WpfWindow("提示").Exist(5) Then
		WpfWindow("提示").WpfButton("确定").Click
		CaptureQuaryRet = False
	Else
		CaptureQuaryRet = True
	End If
End Function


'1：1比对
Function Compare(srcPath, dstPath)
	WpfWindow("窗体列表").WpfTabStrip("WpfTabStrip").Select(2)
	CompareRet = False
	
	'选择原比对图片
	WpfWindow("窗体列表").InsightObject("srcImg").Click
	If not WpfWindow("窗体列表").Dialog("请选择模板照片").Exist(3) Then
		Reporter.ReportEvent micFail, "1：1比对", "无选择图片列表"
		ExitAction(0)
	End If
	WpfWindow("窗体列表").Dialog("请选择模板照片").WinEdit("文件名(N):").Set(srcPath)
	WpfWindow("窗体列表").Dialog("请选择模板照片").WinButton("打开(&O)").Click
	Wait(1)
	
	'选择原比对图片
	WpfWindow("窗体列表").InsightObject("dstImg").Click
	If not WpfWindow("窗体列表").Dialog("请选择模板照片").Exist(3) Then
		Reporter.ReportEvent micFail, "1：1比对", "无选择图片列表"
		ExitAction(0)
	End If
	WpfWindow("窗体列表").Dialog("请选择模板照片").WinEdit("文件名(N):").Set(dstPath)
	WpfWindow("窗体列表").Dialog("请选择模板照片").WinButton("打开(&O)").Click
	Wait(1)
	
	WpfWindow("窗体列表").WpfButton("1:1分析比对").Click
	If WpfWindow("提示").Exist(5) Then
		WpfWindow("提示").WpfButton("确定").Click
		CompareRet = False
	Else
		CompareRet = True
	End If
	
End Function


'目标库批量查询
Function StoreBatchQuary(dirPath, storesArr, scoreThreshold, savePath)
	WpfWindow("窗体列表").WpfTabStrip("WpfTabStrip").Select(3)
	StoreBatchQuaryRet = False
	WpfWindow("窗体列表").WpfButton("清除").Click
	
	'选择目标库
	WpfWindow("窗体列表").WpfButton("选择目标库").Click
	If not WpfWindow("目标库").Exist(3) Then
		Reporter.ReportEvent micFail, "目标库批量查询", "无目标库列表"
		ExitAction(0)
	End If
	WpfWindow("目标库").WpfLink("取消全选").Click
	Set children = WpfWindow("目标库").WpfObject("PART_ItemsScrollViewer").ChildObjects
	count = children.Count
	For i = 0 To count - 1 Step 1
		'print(children(i).GetVisibleText)
		For arrIndex = LBound(storesArr) To UBound(storesArr) Step 1
			If children(i).GetVisibleText = storesArr(arrIndex) Then
				children(i).Click
				StoreBatchQuaryRet = True
			End If
		Next
	Next
	Set children = Nothing
	WpfWindow("目标库").WpfButton("确定").Click
	If not StoreBatchQuaryRet Then
		Reporter.ReportEvent micFail, "目标库批量查询", "无所要的通道"
		ExitAction(0)
	End If
	StoreBatchQuaryRet = False
	Wait(1)
	
	'选择批量目标库
	WpfWindow("窗体列表").WpfButton("选择文件夹").Click
	If not WpfWindow("窗体列表").Dialog("选择文件夹").Exist(3) Then
		Reporter.ReportEvent micFail, "目标库批量查询", "无选择文件夹界面"
		ExitAction(0)
	End If
	WpfWindow("窗体列表").Dialog("选择文件夹").WinEdit("文件夹:").Set(dirPath)
	WpfWindow("窗体列表").Dialog("选择文件夹").WinButton("选择文件夹").Click
	Wait(1)
	
	WpfWindow("窗体列表").InsightObject("thresholdStore").Click
	DelExist()
	KeyBoard(CStr(scoreThreshold))
	
	'查询
	WpfWindow("窗体列表").WpfButton("查询").Click
	If WpfWindow("窗体列表").WpfButton("下载查询结果").Exist(20) Then
		StoreBatchQuaryRet = True
	End If
	WpfWindow("窗体列表").WpfButton("下载查询结果").Click
	If not WpfWindow("窗体列表").Dialog("另存为").Exist(3) Then
		Reporter.ReportEvent micFail, "目标库批量查询", "无另存为界面"
		ExitAction(0)
	End If
	WpfWindow("窗体列表").Dialog("另存为").WinEdit("文件名:").Set(savePath)
	WpfWindow("窗体列表").Dialog("另存为").WinButton("保存(&S)").Click
	If Dialog("另存为").WinButton("是(&Y)").Exist(1) Then
		Dialog("另存为").WinButton("是(&Y)").Click
	End If
End Function

'抓取异常
On Error Resume Next
WpfWindow("窗体列表").Maximize
WpfWindow("窗体列表").WpfObject("人像检索").Click


Dim storesArr, regionsArr
Select Case Parameter("handle")
	
	Case "storeQuary"
	countExcute = ReadExcel(3, 4, "目标库查询", excelPath)
	For caseIndex = 0 To countExcute - 1 Step 1
		StoreQuaryRet = False
		storesArr = Split(ReadExcel(3 + caseIndex, 1, "目标库查询", excelPath), ",")
'		For Iterator = LBound(storesArr) To UBound(storesArr) Step 1
'			print(storesArr(Iterator))
'		Next
'		ExitRun
		imgPath = ReadExcel(3 + caseIndex, 2, "目标库查询", excelPath)
		scoreThreshold =  ReadExcel(3 + caseIndex, 3, "目标库查询", excelPath)
		Call StoreQuary(imgPath, storesArr, scoreThreshold)
		If StoreQuaryRet Then
			Reporter.ReportEvent micPass, "目标库查询", "目标库查询成功"
		Else
			Reporter.ReportEvent micFail, "目标库查询", "目标库查询失败"
		End If
	Next
	If StoreQuaryRet Then
		ExitAction(1)
	Else
		ExitAction(0)
	End If
	
	
	Case "captureQuary"
	countExcute = ReadExcel(3, 5, "抓拍库查询", excelPath)
	For caseIndex = 0 To countExcute - 1 Step 1
		CaptureQuaryRet = False
		dateTime = ReadExcel(3 + caseIndex, 1, "抓拍库查询", excelPath)
		regionsArr = Split(ReadExcel(3 + caseIndex, 2, "抓拍库查询", excelPath), ",")
		imgPath = ReadExcel(3 + caseIndex, 3, "抓拍库查询", excelPath)
		scoreThreshold = ReadExcel(3 + caseIndex, 4, "抓拍库查询", excelPath)
		Call CaptureQuary(dateTime, imgPath, regionsArr, scoreThreshold)
		If CaptureQuaryRet Then
			Reporter.ReportEvent micPass, "抓拍库查询", "抓拍库查询成功"
		Else
			Reporter.ReportEvent micFail, "抓拍库查询", "抓拍库查询失败"
		End If
	Next
	If CaptureQuaryRet Then
		ExitAction(1)
	Else
		ExitAction(0)
	End If
	
	
	Case "compare"
	countExcute = ReadExcel(3, 3, "比对", excelPath)
	For caseIndex = 0 To countExcute - 1 Step 1
		CompareRet = False
		srcPath = ReadExcel(3 + caseIndex, 1, "比对", excelPath)
		dstPath = ReadExcel(3 + caseIndex, 2, "比对", excelPath)
		Call Compare(srcPath, dstPath)
		If CompareRet Then
			Reporter.ReportEvent micPass, "1：1比对", "1：1比对成功"
		Else
			Reporter.ReportEvent micFail, "1：1比对", "1：1比对失败"
		End If
	Next
	If CompareRet Then
		ExitAction(1)
	Else
		ExitAction(0)
	End If
	
	
	Case "storeBatchQuary"
	countExcute = ReadExcel(3, 5, "目标库批量查询", excelPath)
	For caseIndex = 0 To countExcute - 1 Step 1
		StoreBatchQuaryRet = False
		storesArr = Split(ReadExcel(3 + caseIndex, 1, "目标库批量查询", excelPath), ",")
		dirPath = ReadExcel(3 + caseIndex, 2, "目标库批量查询", excelPath)
		scoreThreshold =  ReadExcel(3 + caseIndex, 3, "目标库批量查询", excelPath)
		savePath = ReadExcel(3 + caseIndex, 4, "目标库批量查询", excelPath)
		Call StoreBatchQuary(dirPath, storesArr, scoreThreshold, savePath)
		If StoreBatchQuaryRet Then
			Reporter.ReportEvent micPass, "目标库查询", "目标库批量查询成功"
		Else
			Reporter.ReportEvent micFail, "目标库查询", "目标库批量查询失败"
		End If
	Next
	If StoreBatchQuaryRet Then
		ExitAction(1)
	Else
		ExitAction(0)
	End If
	
	Case Else
		Reporter.ReportEvent micFail, "人像检索", "输入预警历史操作类型错误"
		ExitAction(0)
End Select

If Err.Number <> 0 Then
	errMessage = "错误代码： " & CStr(Err.Number) & ", 错误信息： " & Err.Description & "."
	print(errMessage)
	ExitAction(-1)
End If