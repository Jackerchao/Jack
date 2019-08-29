'*************************************************************************************************
'File Name : Common Function.vbs
'Description: Contains the functions for Reporting
'*************************************************************************************************

Dim objXMLroot
Dim objXMLBP
Dim objXMLTestSuite
Dim objXMLTestCase
Dim objXMLCustomReport
Dim strPassFail
Dim newNameOfFile
Dim deleteoldXML
Dim fso
Dim iFlagFailChk
Dim glostatus
Dim gloTestCase
Dim gloStepName

Dim gloScrshotNameList
Set gloScrshotNameList = CreateObject("System.Collections.ArrayList")


'*************************************************************************************************
'Function : SaveReport
'Purpose: Saves the custom report file
'Inputs : Name of file
'*************************************************************************************************
Function SaveReport(nameOfFile)
	
	objXMLCustomReport.SaveFile nameOfFile
	If iFlagFailChk = "F" Then
		newNameOfFile = Replace(nameOfFile, "\P_", "\F_")
		objXMLCustomReport.SaveFile newNameOfFile
'		Set fso = CreateObject("Scripting.FileSystemObject")
'		Set deleteoldXML = fso.GetFile(nameOfFile)
'		deleteoldXML.Delete
	End If
	
End Function



'*************************************************************************************************
'Function : CreateCutomReportFile
'Purpose: Create the document and the nodes Report and Batch
'Inputs : Batch name
'*************************************************************************************************
Function createCustomReportFile(Desc)
	
	Set ctr = CreateObject("Scripting.FileSystemObject")
	Dim logFileName
	logFileName = FN_RELATIVE_PATH(Environment.Value("ENV_RESULTS_FILE_PATH") &"\_images\testreport.txt")
	strPassFail = "P"
	
	If not ctr.FileExists(logFileName) Then
		fileMode = 2
		Set logfile = ctr.OpenTextFile(logFileName,fileMode,True)
		logfile.WriteLine("<?xml version='1.0'?>")
		logfile.WriteLine("<?xml-stylesheet href='Report.xsl' type='text/xsl'?>")
		logfile.WriteLine("<Report>")
		logfile.WriteLine("</Report>")
		logfile.Close
	End If
'	loading the xml file
'	Set doc = XMLUTIL.CreateXML()
	Set objXMLCustomReport = XMLUtil.CreateXML()
	objXMLCustomReport.LoadFile logFileName
	
	Set objXMLroot = objXMLCustomReport.GetRootElement()
	objXMLroot.AddChildElementByName "TestSuite", Desc
	Set objXMLTestSuite = objXMLroot.ChildElements().Item(1)
	objXMLTestSuite.AddAttribute "StartTime", cstr(Now)
	objXMLTestSuite.AddAttribute "Desc", Desc
	objXMLTestSuite.AddAttribute "EndTime", cstr(Now)
	
	strTimeStamp = cstr(Now)
	strTimeStamp = Replace(strTimeStamp, "/", "_")
	strTimeStamp = Replace(strTimeStamp, ":", "_")
	
	'Save the Report file
	'Environment.Value("REPORT_NAME") = FN_RELATIVE_PATH(Environment.Value("ENV_RESULTS_FILE_PATH") &"\" & strPassFail & "_" & strTestSuiteName & "_"& strTimeStamp &".xml")
	Environment.Value("REPORT_NAME") = FN_RELATIVE_PATH(Environment.Value("ENV_RESULTS_FILE_PATH") &"\" & Desc& "_" & strTimeStamp &".xml")
	Environment.Value("XML_NAME") = Desc & "_" & strTimeStamp & ".xml"
	
	Call SaveReport(Environment.Value("REPORT_NAME"))
	
End Function

'*************************************************************************************************
'Function : addTCNode
'Purpose: Create the testcase node.Adds child elements to the node batch
'Inputs : Test case name
'*************************************************************************************************
Function addTCNode(Desc)

	objXMLTestSuite.AddChildElementByName "TestCase", Desc
	Set resultSet = objXMLTestSuite.ChildElements()
	numberOfTestCases = resultSet.Count()
	Set objXMLTestCase = resultSet.Item(numberOfTestCases)
	objXMLTestCase.AddAttribute "StartTime", cstr(Now)
	objXMLTestCase.AddAttribute "Desc", Desc
	objXMLTestCase.AddAttribute "EndTime", cstr(Now)
	objXMLTestCase.AddAttribute "Status", "1"
End Function

'*************************************************************************************************
'Function : addAggBPNode
'Purpose: Create the BP node.Adds child elements to the node Testcase
'Inputs : BP name
'*************************************************************************************************
Function addAggBPNode(Desc)
	
	objXMLTestCase.AddChildElementByName "ABP", Desc
	Set BPs = objXMLTestCase.ChildElements()
	numberOfBPs = BPs.Count()
	Set objXMLBP = BPs.Item(numberOfBPs)
	objXMLBP.AddAttribute "StartTime", cstr(Now)
	objXMLBP.AddAttribute "Desc", Desc
	objXMLBP.AddAttribute "EndTime", cstr(Now)
	Call SaveReport(Environment.Value("REPORT_NAME"))
	
End Function

'*************************************************************************************************
'Function : addBPNode
'Purpose: Create the BP node.Adds child elements to the node Testcase
'Inputs : BP name
'*************************************************************************************************
Function addBPNode(Desc)
	
	objXMLTestCase.AddChildElementByName "BP", Desc
	Set BPs = objXMLTestCase.ChildElements()
	numberOfBPs = BPs.Count()
	Set objXMLBP = BPs.Item(numberOfBPs)
	objXMLBP.AddAttribute "StartTime", cstr(Now)
	objXMLBP.AddAttribute "Desc", Desc
	objXMLBP.AddAttribute "EndTime", cstr(Now)

	
End Function

'*************************************************************************************************
'Function : addResultNode
'Purpose: Create the result node.Adds child elements to the node BP
'Inputs : BP name
'*************************************************************************************************
Function addResultNode(status,result)
	
	objXMLBP.AddChildElementByName "Result", result
	Set resultNew = objXMLBP.ChildElements()
	stepCount = resultNew.Count()
	Set resultAddAttr = resultNew.Item(stepCount)
	resultAddAttr.AddAttribute "Status", cstr(status)
	Call StatusTimestampUpdate(cstr(status))
	Call SaveReport(Environment.Value("REPORT_NAME")) 
End Function

'*************************************************************************************************
'Function : CaptureTestCaseBmp
'Purpose: Capture bitmap for failed batch
'Inputs : TestCaseId
'*************************************************************************************************
Sub CaptureTestCaseBmp(strScreenShotName,blnScreenShotType)
	
	Dim strFoldername, strTimeStamp, strScreenShotPath
	
	On error resume next
	strTimeStamp = cstr(Now)
	strTimeStamp = Replace(strTimeStamp, "/", "_")
	strTimeStamp = Replace(strTimeStamp, ":", "_")
	
	strScreenShotName = Environment.Value("ENV_SCENARIO_NAME") & "_" & strScreenShotName & "_" & strTimeStamp & ".png"
	
	
	If blnScreenShotType = 0 Then
		strFoldername = "_errorImages"
	Else
		strFoldername = "_exeImages"
	End If
	
	strScreenShotPath = FN_RELATIVE_PATH(Environment.Value("ENV_RESULTS_FILE_PATH")) &"\" & strFoldername & "\" & strScreenShotName
	strScreenShotRelativePath = "..\Reports\" & strFoldername & "\" & strScreenShotName
	
	Environment.Value("SCREEN_PATH") = strScreenShotPath
	'Just capture
	Desktop.CaptureBitmap strScreenShotPath, True
	
	'Call addScreenShotLinkInReport(strScreenShotPath,blnScreenShotType)
	Call addScreenShotLinkInReport(strScreenShotRelativePath,blnScreenShotType)
	
	gloScrshotNameList.Add strScreenShotName
	
End Sub

'*************************************************************************************************
'Function : addScreenShotLinkInReport
'Purpose: Capture the Link to the Screenshot in custom report
'Inputs : strImagePath - path of the image
'		  blnScreenShotType - Type of screen shot(Pass or fail)
'*************************************************************************************************
Function addScreenShotLinkInReport(strImagePath,blnScreenShotType)
	
	'add the result element
	objXMLBP.AddChildElementByName "Result", ""
	'access the newly added node and add the ImagePath attribute
	Set objResultNodes = objXMLBP.ChildElements()
	intChildItems = objResultNodes.Count()
	Set objNewResultNode = objResultNodes.Item(intChildItems)
	
	If blnScreenShotType <> 0 Then
		objNewResultNode.AddAttribute "ScreenShotPath", strImagePath
	Else
		objNewResultNode.AddAttribute "ErrorScreenShotPath", strImagePath
	End If
	
	Call SaveReport(Environment.Value("REPORT_NAME"))
End Function

'*************************************************************************************************
'Function : ReportError
'Purpose: This Reports error
'Inputs : strFunctionName- The name of the function
'		  strCustomError- Error message
'		  blnCaptureScreenShot- Ture - Capture screenshot
'*************************************************************************************************
Function ReportError(strFunctionName,strCustomError,blnCaptureScreenShot)
	
	Dim strReportResult
	
	'Add the error number and description
	If Err.Number <> 0 Then
		strReportResult = "Error# : "& Err.Number &" Description : "& Err.Description &" : " & strCustomError
		'Msgbox Err.Number & Err.Description
	Else
		strReportResult = strCustomError		
	End If
	'On Error Resume Next
	'Report in QTP
	Reporter.ReportEvent micFail,strFunctionName,strReportResult
	
	'Report in the custom xml report file
'	Call addResultNode(3,strCustomError& " " & Err.Description)
	Call addResultNode(3,strCustomError)
	
	If blnCaptureScreenShot = True Then
		'Capture the error screen shot
		Call CaptureTestCaseBmp(Glbl_Error_ScreenShotName,0)
	End If
	'glostatus = glostatus & ";" & "Failed"
	glostatus = "Failed"
	'Return the value
	ReportError = strCustomError
	err.clear
	
End Function


'*************************************************************************************************
'Function : ReportNoRun
'Purpose: This Reports the No run
'Inputs : strFunctionName- The name of the function
'		  blnCaptureScreenShot- Ture - Capture screenshot
'*************************************************************************************************
Function ReportNoRun(strFunctionName,blnCaptureScreenShot)
	
	Dim strReportResult
	
	strReportResult = "No Run due to previous keyword failed."
	
	'On Error Resume Next
	''Report in QTP
	Reporter.ReportEvent micFail,strFunctionName,strReportResult
	
	'report in the custom xml report file
	Call addResultNode(4,strReportResult)
	
	If blnCaptureScreenShot = True Then
		'Capture the error screen shot
		Call CaptureTestCaseBmp(Glbl_Error_ScreenShotName,0)
	End If
	'glostatus = glostatus & ";" & "No Run"
	glostatus = "No Run"
	'Return the value
	ReportNoRun = strReportResult
	err.clear
	
End Function

'*************************************************************************************************
'Function : ReportPass
'Purpose: Function to report pass
'Inputs : strFunctionName- The name of the function
'		  intStatus - Event Status, 0-Pass, 2-Done
'		  intReportLocation - which reporting needs to be done. 1-QTP only, 2-Both, 3-XNL only
'*************************************************************************************************
Function ReportPass(intStatus, strFunctionName, strMessage, intRptLocn)
	
	On Error Resume Next
	Dim strReportResult
	
	Select Case intRptLocn
		Case 1 'report in QTP
			Reporter.ReportEvent 0, strFunctionName, strMessage
		Case 2 'report in QTP
			Reporter.ReportEvent 0, strFunctionName, strMessage
			'report in the custom xml report file.
			Call addResultNode(intStatus,strMessage)
		Case 3 'report in the custom xml report file.
			Call addResultNode(intStatus,strMessage)
	End Select
	
	'glostatus = glostatus & ";" & "Passed"
	glostatus = "Passed"
	'return the value
	Call CaptureTestCaseBmp(Glbl_Error_ScreenShotName,0)
	ReportPass = True
	
End Function

'*************************************************************************************************
'Sub : StatusTimestampUpdate
'Purpose: For the status and the timestamp Update
'Inputs : 
'*************************************************************************************************
Sub StatusTimestampUpdate(status)
	
	On error resume next
	strTimeStamp = Cstr(Now)
	
	If status = "3" Then
		objXMLTestCase.RemoveAttribute "status"
		objXMLTestCase.AddAttribute "status", status
	End If
	
	objXMLTestCase.RemoveAttribute "EndTime"
	objXMLTestCase.AddAttribute "EndTime" ,strTimeStamp
	
	objXMLTestSuite.RemoveAttribute "EndTime"
	objXMLTestSuite.AddAttribute "EndTime", strTimeStamp
	
	objXMLBP.RemoveAttribute "EndTime"
	objXMLBP.AddAttribute "EndTime", strTimeStamp
	
	If status = "1" Then
		strPassFail = "p"
	ElseIf status = "3" Then
		strPassFail = "F"
		iFlagFailChk = "F"
	End If
	
End Sub

'*************************************************************************************************
'Function : addScreenShotPassLinkInReport
'Purpose: Capture bitmap for the failed batch into xml report
'Inputs : bmpath
'*************************************************************************************************

Function addScreenShotPassLinkInReport(strImagePath)
	
	'add the result element
	objXMLBP.AddChildElementByName "Result", ""
	'access the newly added node and add the ImagePath attribute
	Set objResultNodes = objXMLBP.ChildElements()
	intChildItems = objResultNodes.Count()
	Set objNewResultNode = objResultNodes.Item(intChildItems)
	objNewResultNode.AddAttribute "ImagePathDone", strImagePath
End Function

'*************************************************************************************************
'Function : AttachToQC
'Purpose: Attach file to QC folder
'Inputs : strQCPath - folder in QC
'		  strSourceFilePath - File which needs to be uploaded
'*************************************************************************************************
Function AttachToQC()
	
	Set objAttachments = QCUtil.CurrentRun.Attachments
	Set objAttachNew = objAttachments.AddItem(Null)
	objAttachNew.FileName = Environment.Value("REPORT_NAME")
	objAttachNew.Type = 1
	objAttachNew.Post()
End Function
'*************************************************************************************************
'Function : ExportXMLReportContent
'Purpose: Export execution results to a temp txt file stored under DataSheet Folder
'Inputs : FileContent - test execution result
'Output: Testresult.txt will be generated under datasheet folder
'*************************************************************************************************
Function ExportXMLReportContent(FIleContent)
	
	Dim objFile
	Dim strImageFile
	Dim myFile
	strImageFile = FN_RELATIVE_PATH("..\..\DataSheets\TestResults.txt")
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If objFSO.FileExists(strImageFile) Then
		Set objFile = objFSO.OpenTextFile (strImageFile, ForAppending, True)
		objFile.WriteLine FileContent
		objFile.Close
	Else
		Set objFile = objFSO.CreateTextFile(strImageFile, True)
		objFile.WriteLine FileContent
		objFile.Close
	End If
End Function