vbsPath = DataTable("A", dtGlobalSheet)
excelPath = DataTable("B", dtGlobalSheet)
ExecuteFile(vbsPath)
caseExcute = ReadExcel(2, 3, "MainAction", excelPath)
Dim caseExcuteArr
caseExcuteArr = Split(caseExcute, ",")
Dim retMain
For arrIndex = LBound(caseExcuteArr) To UBound(caseExcuteArr) Step 1
	print(caseExcuteArr(arrIndex))
	Select Case caseExcuteArr(arrIndex)
		'新建目标库
		Case 1
			handle = "newStore"
			retMain = RunAction("目标库", oneIteration, handle)
			print("新建目标库执行结果：" & CStr(retMain))			
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'修改目标库		
		Case 2
			handle = "modifyStore"
			retMain = RunAction("目标库", oneIteration, handle)
			print("修改目标库执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'新建目标人
		Case 3
			handle = "newTemplate"
			retMain = RunAction("目标库", oneIteration, handle)
			print("新建目标人执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'修改目标人
		Case 4
			handle = "modifyTemplate"
			retMain = RunAction("目标库", oneIteration, handle)
			print("修改目标人执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'新建区域
		Case 5
			handle = "newRegion"
			retMain = RunAction("新建区域", oneIteration, handle)
			print("新建区域执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'新建通道
		Case 6
			handle = "newChannel" 
			retMain = RunAction("新建通道", oneIteration)
			print("新建通道执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'新建比对策略
		Case 7
			handle = "newStrategy"
			retMain = RunAction("新建策略", oneIteration, handle)
			print("新建比对策略执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
			
		'新建布控任务
		Case 8
			handle = "newTask"
			retMain = RunAction("布控任务", oneIteration, handle)
			print("新建布控任务执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'启动布控任务
		Case 9
			handle = "startTask"
			retMain = RunAction("布控任务", oneIteration, handle)
			print("启动布控任务执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'关闭布控任务
		Case 10
			handle = "stopTask"
			retMain = RunAction("布控任务", oneIteration, handle)
			print("关闭布控任务执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'修改布控任务
		Case 11
			handle = "modifyTask"
			retMain = RunAction("布控任务", oneIteration, handle)
			print("修改布控任务执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
			
		'删除布控任务
		Case 12
			handle = "deleteTask"
			retMain = RunAction("布控任务", oneIteration, handle)
			print("删除布控任务执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
			
		'首页修改布控任务
		Case 13
			handle = "modifyTaskPage"
			retMain = RunAction("首页", oneIteration, handle)
			print("首页修改布控任务执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'首页修改比对策略
		Case 14
			handle = "modifyStrategyPage"
			retMain = RunAction("首页", oneIteration, handle)
			print("首页修改比对策略执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'推送告警
		Case 15
			handle = "alarmPushed"
			retMain = RunAction("首页", oneIteration, handle)
			print("首页修改比对策略执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
			
		'告警详情
		Case 16
			handle = "alertInformation"
			retMain = RunAction("预警历史", oneIteration, handle)
			print("告警详情执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
			
		'核查反馈
		Case 17
			handle = "verificationCheck"
			retMain = RunAction("预警历史", oneIteration, handle)
			print("核查反馈执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
			
		'未复核告警个数检查
		Case 18
			handle = "unCheckNum"
			retMain = RunAction("预警历史", oneIteration, handle)
			print("未复核告警个数检查执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
			
		'抓拍历史
		Case 19
			handle = "captureChosen"
			retMain = RunAction("抓拍历史", oneIteration, handle)
			print("抓拍历史执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
			
		'目标库查询
		Case 20
			handle = "storeQuary"
			retMain = RunAction("人像检索", oneIteration, handle)
			print("目标库查询执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
			
		'抓拍库查询
		Case 21
			handle = "captureQuary"
			retMain = RunAction("人像检索", oneIteration, handle)
			print("抓拍库查询执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
			
		'1：1比对
		Case 22
			handle = "compare"
			retMain = RunAction("人像检索", oneIteration, handle)
			print("1：1比对执行结果：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
		
		'目标库批量查询
		Case 23
			handle = "storeBatchQuary"
			retMain = RunAction("人像检索", oneIteration, handle)
			print("目标库批量查询：" & CStr(retMain))
			If WpfWindow("提示").Exist(1) Then
				WpfWindow("提示").WpfButton("确定").Click
			End If
			Wait(3)
			
		Case Else
			Reporter.ReportEvent micFail, "用例选择", "用例选择参数错误"
			ExitRun
	End Select
Next

'Dim retMain 
'
'handle = "newStore"
'retMain = RunAction("目标库", oneIteration, handle)
'print("新建目标库执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "modifyStore"
'retMain = RunAction("目标库", oneIteration, handle)
'print("修改目标库执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "newTemplate"
'retMain = RunAction("目标库", oneIteration, handle)
'print("新建目标人执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "modifyTemplate"
'retMain = RunAction("目标库", oneIteration, handle)
'print("修改目标人执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'retMain = RunAction("新建区域", oneIteration)
'print("新建区域执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'Dim ChannelTypes(4)
'ChannelTypes(0) = "Video"
'ChannelTypes(1) = "RTSP"
'ChannelTypes(2) = "GB28181"
'ChannelTypes(3) = "File"
'
'Dim channelNames(4)
'channelNames(0) = "VideoChannel"
'channelNames(1) = "RTSPChannel"
'channelNames(2) = "GBChannel"
'channelNames(3) = "FileChannel"
'
'Dim channelNos(4)
'channelNos(0) = "52_#_51000001_1_102"
'channelNos(1) = "rtsp://192.168.1.176/9.sdp"
'channelNos(2) = "34020000001320000003"
'channelNos(3) = "/data/video/IMG_0229.MOV"
'For i = 0 To 3 Step 1
'	ChannelType = ChannelTypes(i)
'	channelName = channelNames(i)
'	channelNo = channelNos(i)
'	retMain = RunAction("新建通道", oneIteration, ChannelType, channelName, channelNo)
'	print("新建" & CStr(ChannelType) & "视频通道执行结果：" & CStr(retMain))
'	Wait(1)
'Next
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'Dim StrategyTypes(3)
'StrategyTypes(0) = "Threshold"
'StrategyTypes(1) = "Count"
'StrategyTypes(2) = "Both"
'
'Dim strategyNames(3)
'strategyNames(0) = "ThresholdStrategy"
'strategyNames(1) = "CountStrategy"
'strategyNames(2) = "BothStrategy"
'
'For i = 0 To 2 Step 1
'	StrategyType = StrategyTypes(i)
'	strategyName = strategyNames(i)
'	retMain = RunAction("新建策略", oneIteration, StrategyType, strategyName)
'	print("新建" & StrategyType & "策略执行结果：" & CStr(retMain))
'	Wait(3)
'Next
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "New"
'taskName = "ForverTask"
'taskType = 0
'retMain = RunAction("布控任务", oneIteration, taskName, taskType, handle)
'print("新建布控任务执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "New"
'taskName = "CustomTask"
'taskType = 1
'retMain = RunAction("布控任务", oneIteration, taskName, taskType, handle)
'print("新建布控任务执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
''handle：操作方式，new：新建布控任务，start：启动布控任务，stop：关闭布控任务，delete：删除布控任务
'handle = "Start"
'taskName = "ForverTask"
'retMain = RunAction("布控任务", oneIteration, taskName, taskType, handle)
'print("启动布控任务执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "Stop"
'taskName = "ForverTask"
'retMain = RunAction("布控任务", oneIteration, taskName, taskType, handle)
'print("关闭布控任务执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "Delete"
'taskName = "ForverTask"
'retMain = RunAction("布控任务", oneIteration, taskName, taskType, handle)
'print("删除布控任务执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "Modify"
'taskName = "CustomTask"
'retMain = RunAction("布控任务", oneIteration, taskName, taskType, handle)
'print("编辑布控任务执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "taskDetails"
'retMain = RunAction("首页", oneIteration, handle)
'print("查看任务详情执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "videoPreview"
'retMain = RunAction("首页", oneIteration, handle)
'print("视频预览执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If

'handle = "modifyTask"
'retMain = RunAction("首页", oneIteration, handle)
'print("修改布控任务执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If

'handle = "modifyStrategy"
'retMain = RunAction("首页", oneIteration, handle)
'print("修改比对策略行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "alarmPushed"
'retMain = RunAction("首页", oneIteration, handle)
'print("告警推送执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "verificationCheck"
'retMain = RunAction("预警历史", oneIteration, handle)
'print("核查反馈执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "alertInformation"
'retMain = RunAction("预警历史", oneIteration, handle)
'print("查询详细信息执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "uncheckNum"
'retMain = RunAction("预警历史", oneIteration, handle)
'print("未复核告警个数确认执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "captureChosen"
'retMain = RunAction("抓拍历史", oneIteration, handle)
'print("查询选择项抓拍数据执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "storeQuary"
'retMain = RunAction("人像检索", oneIteration, handle)
'print("目标库查询执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "captureQuary"
'retMain = RunAction("人像检索", oneIteration, handle)
'print("抓拍库查询执行结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "compare"
'retMain = RunAction("人像检索", oneIteration, handle)
'print("1：1比对结果：" & CStr(retMain))
'Wait(3)
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
'handle = "storeBatchQuary"
'retMain = RunAction("人像检索", oneIteration, handle)
'print("目标库批量查询执行结果：" & CStr(retMain))
'
'If WpfWindow("提示").Exist(1) Then
'	WpfWindow("提示").WpfButton("确定").Click
'End If
'
ExitRun