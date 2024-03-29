gloExecuteFlag = True
'*************************************************************************************************
'Function : Broswer_Launch
'Purpose: Launch broswer
'Inputs : URL
'Output: Broswer launch successfully
'*************************************************************************************************
Function Broswer_Launch()
	
	'Check execute flag
	If gloExecuteFlag = False Then
		Call ReportNoRun("Broswer_Launch", True)
		Exit Function
	End If
	
	On Error Resume Next
	
	Dim strURL
	Dim strTempURL
	
	Print_Log "Broswer_Launch"
	
	'retrieve Broswer URL
	strTempURL = RetrieveTestData(CStr(Environment.Value("ENV_SCENARIO_NAME")), "URL")
	strURL = Environment.Value(strTempURL)
	
	'Launch IE browser and navagite to URL
	Systemutil.Run "iexplore.exe", strURL,,, 3
	
	'Clear cache and delete cookies
	Browser("IE").ClearCache
	Browser("IE").DeleteCookies
	
	If Browser("IE").Page("BaiDu").WebArea("BaiduLogo").Exist(5) Then
		Call ReportPass(1, "Broswer_Launch", "URL launched successfully", 2)
	Else
		Call ReportError("Broswer_Launch", "URL launched failed", True)
	End If
	
	On Error Goto 0
	Err.Clear
	
	
End Function

'*************************************************************************************************
'Function : CloseBrowser
'Purpose: Close broswer
'Inputs : NA
'Output: Broswer close successfully
'*************************************************************************************************
Function CloseBrowser()

	On Error Resume Next
	
	Print_Log "CloseBrowser"
	
	Dim objWMIService, objprocess, colProcess
	Dim strComputer, strProcessKill
	strComputer = "."
	strProcessKill = "'iexplore.exe'"
	
	Set objWMIService = GetObject("winmgmts:" _ 
		&"{impersonationlevel = impersonate}!\\" _
			& strComputer & "\root\cimv2")
			
	Set colProcess = objWMIService.ExeQuery _
		("Select * from win32_Process Where Name = " & strProcessKill)
		
	For Each objprocess in colProcess
		objprocess.Terminate()
	Next
	
	WSCript.Echo "Just killed process " & strProcessKill_ & " on " & strComputer
	
	WSCript.Quit
	
	Err.Clear
	
	If Err.Number <> 0 Then
		Call ReportError("CloseBrowser", "URL launched failed", True)
	Else
		Call ReportPass(1, "CloseBrowser", "URL launched successfully", 2)
	End If
	On Error Goto 0
	Err.Clear
End Function