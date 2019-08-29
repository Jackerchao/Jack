'Call Intilization File
Environment.LoadFromFile(FN_RELATIVE_PATH("..\..\Library\Config.xml"))

'Load Intilization Scripts
ExecuteFile FN_RELATIVE_PATH(FN_RELATIVE_PATH(Environment.Value("ENV_INTILIZATION_FILES")))

'*************************************************************************************************
'Type : Generic Function
'Name:	FN_RELATIVE_PATH
'Owner:
'Objective:To get the absolute path
'Prerequisites:
'Parmeters:strFileName(variable having relative path)
'OutParmeters:The Absolute path
'*************************************************************************************************
Function FN_RELATIVE_PATH(strFileName)
	
	On Error Resume Next
	
	strFileName = Replace(strFileName,"/","\")
	If Instr(1, strFileName, "..\" ) > 0 Then
		
		strTemp2Env_Set = Split(Environment.Value("TestDir"),"\")
		
		strTemp1Env = Split(strFileName, "\")
		intCoutVal = 0
		
		For i = 0 To Ubound(strTemp1Env)
			If strTemp1Env(i) = ".." Then
				intCoutVal = intCoutVal + 1
			End If
		Next
		
		StartVal = Ubound(strTemp2Env_Set) - intCoutVal + 1
		AddVal = Ubound(strTemp1Env) - intCoutVal - 1
		ReDim Preserve strTemp2Env_Set(StartVal + AddVal + 1 )
		
		If intCoutVal > Ubound(strTemp2Env_Set) Then
			Exit Function
		End If
		
		For j = StartVal To (StartVal + AddVal + 1 )
			strTemp2Env_Set(j) = strTemp1Env(intCoutVal)
			intCoutVal = intCoutVal + 1
		Next
		
		FN_RELATIVE_PATH = Join(strTemp2Env_Set, "\")
	Else
	
		FN_RELATIVE_PATH = strFileName
		
	End If
	
End Function

'*************************************************************************************************
'File Name: Master Script.vbs
'Type : Generic Function
'Name:	FN_RELATIVE_PATH
'Owner:
'Objective:To get the absolute path
'Prerequisites:
'Parmeters:strFileName(variable having relative path)
'OutParmeters:The Absolute path
'*************************************************************************************************
Err.Clear
On Error Resume Next
Environment.Value("ENV_BP_ERR_FLAG") = 1

Dim strTestScenarioName, strTestSuiteName, strTestScenarioDescription
Dim strTestCasePath
Dim strTempReportName
Dim strReportName
Dim strTestResult
Dim strResultFile

strTestResult = FN_RELATIVE_PATH("..\..\DataSheets\TestResult.txt")
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(strTestResult) Then
	Set strResultFile = objFSO.GetFile(strTestResult)
	strResultFile.Delete
End If

'add a sheet for storing the Batch Details
DataTable.AddSheet("BATCH_SUMMARY")

'import the Batch details sheet from the env path to the local sheet
strFileName = FN_RELATIVE_PATH(Environment.Value("ENV_BATCH_DETAILS_FILE"))

strTempReportName = Right(Environment.Value("ENV_BATCH_DETAILS_FILE"), Len(Environment.Value("ENV_BATCH_DETAILS_FILE")) - 17)
strReportName = Left(strTempReportName, Len(strTempReportName) - 5)

DataTable.ImportSheet strFileName, "TestCase", "BATCH_SUMMARY"

'Code to retrieve Test Suite Name
Dim arrFolder
Dim objFileSystem
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
Set objFile = objFileSystem.GetFile(strFileName)
arrFolder = Split(objFile.ParentFolder, "\")
strTestSuiteName = arrFolder(UBound(arrFolder) - 1) 
Set objFileSystem = Nothing

Dim GloTempRowCountStart 'global parameter, the start row of the keyword
Dim GloTempRowCountEnd 'global parameter, the end row of the keyword
Dim TempCurrentRow 'temporary current row of the test case sheet
Dim intSheetCount 'total sheet count
Dim intTestDataRowCount 'total row count for the test data sheet
Dim intNoOfRowInTaCounter 'test data sheet row couser
Dim GloCurrentActiveRow 'current active row
Dim intReportCaseCount 'test case number in the report

Environment.Value("ENV_Order") = 0
Environment.Value("ENV_Record") = 0

'get the row count
intRowCount = DataTable.GetSheet("BATCH_SUMMARY").GetRowCount
DataTable.GetSheet("BATCH_SUMMARY").SetCurrentRow(1)
Environment.Value("PRV_ENV_BP_TEST_CASE") = " "

'create the custom report file
Call createCustomReportFile(strReportName)
intReportCaseCount = 1

i = 1

Do
	DataTable.GetSheet("BATCH_SUMMARY").SetCurrentRow(i)
	
	Environment.Value("ENV_SCENARIO_NAME") = DataTable.Value("Auto_Test_Case_Name", "BATCH_SUMMARY")
	
	'If the BPC marks as "Yes", it will be executed, otherwise, it will not.
	If DataTable.Value("PRM_EXECUTE_SCRIPT", "BATCH_SUMMARY") = "Yes" Then
		
		Reporter.ReportEvent micDone, "-----------"& Environment.Value("ENV_SCENARIO_NAME") & "--------------", "Start Test Case Execution: '" & Environment.Value("ENV_SCENARIO_NAME") & "'"
		
		intSheetCount = DataTable.GetSheetCount()
		'delete the previous test data sheet
		If IFDataSheetExist(Environment.Value("ENV_SCENARIO_NAME")) = False Then
			If intSheetCount = 4 Then
				DataTable.DeleteSheet(intSheetCount)
			End If
			DataTable.AddSheet(Environment.Value("ENV_SCENARIO_NAME"))
			DataTable.ImportSheet strFileName, Environment.Value("ENV_SCENARIO_NAME"), Environment.Value("ENV_SCENARIO_NAME")
		End If
		
		intTestDataRowCount = DataTable.GetSheet(Environment.Value("ENV_SCENARIO_NAME")).GetRowCount
		For intNoOfRowInTaCounter = 1 to intTestDataRowCount
			DataTable.GetSheet(Environment.Value("ENV_SCENARIO_NAME")).SetCurrentRow(intNoOfRowInTaCounter)
			
			'If the test data is mark as "Yes", it will be invoked during execution
			If DataTable.Value("DataSequence", Environment.Value("ENV_SCENARIO_NAME")) = "Yes"  Then
				Environment.Value("ALM_Case_Identifier") = DataTable.Value("ALMTestCaseId", Environment.Value("ENV_SCENARIO_NAME"))
				GloCurrentActiveRow = intNoOfRowInTaCounter
				
				Call ReturnBPCEndRow(Environment.Value("ENV_SCENARIO_NAME"))
				
				Call addTCNode(Environment.Value("ALM_Case_Identifier"))
				
				Call ExportXMLReportContent("---"& vbCrlf & Environment.Value("ALM_Case_Identifier") & ";")
				intReportCaseCount = intReportCaseCount + 1
				
				For intNoOfRowInTsCounter = GloTempRowCountStart to GloTempRowCountEnd
					
					'DataTable.SetCurrentRow(intNoOfRowInTsCounter)
					DataTable.GetSheet("BATCH_SUMMARY").SetCurrentRow(intNoOfRowInTsCounter)
					
					'get the BP element name in ENV variable
					Environment.Value("ENV_BP_ELEMENT_NAME") = DataTable.Value("BPC_Sequence", "BATCH_SUMMARY")
					Environment.Value("ENV_BP_ELEMENT_NAME_REPORT") = DataTable.Value("BPC_Sequence", "BATCH_SUMMARY")
					Environment.Value("ENV_BP_ACTION_ENABLED") = DataTable.Value("PRM_EXECUTE_SCRIPT", "BATCH_SUMMARY")
					Environment.Value("ENV_BP_TEST_CASE") = DataTable.Value("BPC_TestCase", "BATCH_SUMMARY")
					Environment.Value("ENV_BP_DESCRIPTION") = DataTable.Value("BPC_Description", "BATCH_SUMMARY")
					
					'step flag
					Environment.Value("ENV_BP_EXECUTE") = DataTable.Value("BPC_Execute", "BATCH_SUMMARY")
					
					If Environment.Value("ENV_Order") = 0 Then
						Call addBPNode(Environment.Value("ENV_BP_ELEMENT_NAME"))
					Else
						Rec = "Order :=" & Environment.Value("ENV_Order") & "::Record:=" & Environment.Value("ENV_Record") & "::BPC_Name:=" &Environment.Value("ENV_BP_ELEMENT_NAME")
						'print "Rec" &Rec
						Call addBPNode(Rec)
					End If
					
'					If Environment.Value("ENV_BP_EXECUTE") = "Yes" Then
'						Call addBPNode(Environment.Value("ENV_BP_ELEMENT_NAME"))
'					End If

					Glbl_Error_ScreenShotName = Environment.Value("ENV_BP_ELEMENT_NAME")
					
					Environment.Value("ENV_BP_ERR_FLAG") = 1
					
					'execute the keywords
					If Environment.Value("ENV_BP_EXECUTE") = "Yes" Then
						Reporter.ReportEvent micDone, "Start BPC Execution: '"& Environment.Value("ENV_BP_ELEMENT_NAME") & "'", "Start BPC '" & Environment.Value("ENV_BP_ELEMENT_NAME") & "'"
						
						Execute "Call " & Environment.Value("ENV_BP_ELEMENT_NAME") & "()"
						
						Call ExportXMLReportContent(Environment.Value("ENV_BP_ELEMENT_NAME") & "-" & glostatus & ";")
						gloStepName = gloStepName & ";" & Environment.Value("ENV_BP_ELEMENT_NAME")
					Else
						Reporter.ReportEvent micDone, "Skip BPC Execution: '"& Environment.Value("ENV_BP_ELEMENT_NAME") & "'", "Skip BPC Execution: '" & Environment.Value("ENV_BP_ELEMENT_NAME") & "'"	           
					End If
					
					
					If Environment.Value("ENV_BP_ERR_FLAG") = 0  Then
						Reporter.ReportEvent micDone, "TestcaseStepExecution", "The test case step at sequence" & Environment.Value("ENV_BP_ERR_SEQ") & " of the Business Process Component '" & Environment.Value("ENV_BP_ELEMENT_NAME")
					Else
						If strBPExecuteFlag = "Yes" Then
							Reporter.ReportEvent micDone, "Business Process Component '" & Environment.Value("ENV_BP_ELEMENT_NAME") & "' passed", "Business Process Componment '" & strBPName & "' passed" 
						End If					
						
					End If
					
					
				Next
				
				Else
					DataTable.GetSheet(Environment.Value("ENV_SCENARIO_NAME")).SetNextRow
				
			End If
			
		Next
		
		'DataTable.GetSheet("BATCH_SUMMARY").SetNextRow
		i = DataTable.GetSheet("BATCH_SUMMARY").GetCurrentRow
		i = i + 1
		
	Else
		'DataTable.GetSheet("BATCH_SUMMARY").SetNextRow
		i = DataTable.GetSheet("BATCH_SUMMARY").GetCurrentRow
		i = i + 1
	
	End If
	
	intRowCount = DataTable.GetSheet("BATCH_SUMMARY").GetRowCount
	
	
Loop While i < intRowCount



'*************************************************************************************************
'Function: ReturnBPCEndRow
'Type : 
'Name:	
'Owner:
'Objective:Getting the start and end row for the keyword
'Prerequisites:
'Parmeters:
'OutParmeters:Returning the start and the end row of given scenario
'*************************************************************************************************
Function ReturnBPCEndRow(strBPCValue)
	
	Dim i
	Dim j
	Dim rowCount
	Dim intTempRowCount
	Dim strTempBPCCase
	Dim ArrTempArray
	strTempBPCCase = ""
	
	rowCount = DataTable.GetSheet("BATCH_SUMMARY").GetRowCount
	For i = 1 to rowCount
		DataTable.SetCurrentRow(i)
		strTempBPCCase = strTempBPCCase & ";" & DataTable.Value("Auto_Test_Case_Name", "BATCH_SUMMARY")	
	Next
	
	ArrTempArray = Split(strTempBPCCase, ";")
	
	For j =0 to Ubound(ArrTempArray) - 1
		
		If ArrTempArray(j) = strBPCValue Then
			Exit For			
		End If
	Next
	
	GloTempRowCountStart = j
	Do
		j = j + 1
	Loop While ArrTempArray(j) = ""
	
	GloTempRowCountEnd = j - 1
	
End Function

'*************************************************************************************************
'Function: RetrieveTestData
'Type : 
'Name:	
'Owner:
'Objective:Retrieve the test data from the test data sheet
'Prerequisites:
'Parmeters:
'OutParmeters:Reteieve value for particular parameter
'*************************************************************************************************
Function RetrieveTestData(strBPCValue, strParameter)
	
	DataTable.GetSheet(strBPCValue).SetCurrentRow(GloCurrentActiveRow)
	If (DataTable.GetSheet(strBPCValue).GetParameter(strParameter).Name <> "") Then
		RetrieveTestData = DataTable.GetSheet(strBPCValue).GetParameter(strParameter).Value
	Else
	    RetrieveTestData = ReportError("RetrieveTestData", "The parameter does not exist, please kindly check the test data sheet.", True)
	End If
End Function


'*************************************************************************************************
'Function: RetrieveTestData
'Type : 
'Name:	
'Owner:
'Objective:Retrieve the test data from the test data sheet
'Prerequisites:
'Parmeters:
'OutParmeters:Reteieve value for particular parameter
'*************************************************************************************************
Function IFDataSheetExist(ByVal SheetName)
	
	IFDataSheetExist = True
	On Error Resume Next
	Dim oTest
	Set oTest = DataTable.GetSheet(SheetName)
	
	If Err.Number Then
		IFDataSheetExist = False
	Else
		IFDataSheetExist = True
	End If
End Function



