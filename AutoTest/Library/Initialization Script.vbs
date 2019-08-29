'*************************************************************************************************
'Type : Generic Function
'Name:	Initilization File
'Objective:To set the Test script for Initliztion
'Prerequisites:none
'Parmeters:none
'OutParmeters:none
'*************************************************************************************************
' Pre-check if browser and excel are opend

Err.Clear
On Error Resume Next

Call PreRun

strObjectRep = FN_RELATIVE_PATH(Environment.Value("ENV_OBJECT_REPOSITORY"))

If CDbl(Environment("ProductVer")) = 9.0 or CDbl(Environment("ProductVer")) = 9.1 Then
	
	If QCUTIL.IsConnected = False Then
		
		Set qtApp = CreateObject("QuickTest.Application")
		Set qtRepositories = qtApp.Test.Actions("MasterScript").ObjectRepositories
		
		strObjectRep = FN_RELATIVE_PATH(Environment.Value("ENV_OBJECT_REPOSITORY"))
		qtRepositories.RemoveAll
		
		If qtRepositories.Find(strObjectRep) = -1 Then
			qtRepositories.Add strObjectRep, 1
		End If
		
		' Set the new object repository configuration as the default for all new actions
		qtRepositories.SetAsDefault 'Set object repositories associated with the "Login" action as the default for all new actions
		Set qtRepositories = Nothing
		
	End If
	
End If

If CDbl(Environment("ProductVer")) >= 9.2 Then
	
	RepositoriesCollection.RemoveAll
	RepositoriesCollection.Add(strObjectRep)
	
End If


Environment("ProceedFlagBP") = 0
Environment.Value("ENV_CURRENT_SCREEN_NAME") = -1
Environment("StrBRDExecuteFlag") = -1
Environment.Value("ENV_BEFORE_VERIFICATION_FLAG") = False
Environment.Value("ENV_VERIFICATION_SUCCESS_FLAG") = True

'Loading the generic functions and keyword library functions
'ExecuteFile FN_RELATIVE_PATH(Environment.Value("ENV_GENERIC_FUNCTIONS"))

'Loading the common functions
ExecuteFile FN_RELATIVE_PATH(Environment.Value("ENV_COMMON_FUNCTIONS"))

strReportPath = FN_RELATIVE_PATH(Environment.Value("ENV_RESULTS_FILE_PATH"))

Set fso = CreateObject("Scripting.FileSystemObject")
If (fso.FolderExists(strReportPath)) = False Then
	
	Reporter.ReportEvent micFail, "Report Path not found", "Report Path not found. Please enter correct path in the config file against ENV_RESULTS_FILE_PATH and continue execution "
	ExitTest
End If


'Reporting File
'ExecuteFile FN_RELATIVE_PATH(Environment.Value("ENV_REPORTING_FILES"))

'Keyword Library
'ExecuteFile FN_RELATIVE_PATH(Environment.Value("ENV_KEYWORD_LIBRARY"))

'Loading business process library
ExecuteFile FN_RELATIVE_PATH(Environment.Value("ENV_BP_LIBRARY"))

'Loading support files
strSupportFilePath = Environment.Value("ENV_SUPPORT_FILES")

If strSupportFilePath <> "" Then
	
	Call LoadSupportFiles(Environment.Value("ENV_SUPPORT_FILES"))
	
End If

'Loading the required files
'Setting the data pointer variable to empty
Environment.Value("ENV_DATA_POINTER") = -1
Environment.Value("ENV_FOLDER_PATH") = FN_RELATIVE_PATH("..\..")

Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions") ' Create the Run Results Options object
qtResultsOpt.ResultsLocation = FN_RELATIVE_PATH(Environment.Value("ENV_RESULTS_FILE_PATH")) 'Set the results location


'Creating a folder to download the Bath XLS and Datashet from the system

If QCUtil.IsConnected = true Then
	
	Print "QC_Util_Current test" & QCUtil.CurrentTest.Name
	
	'Get the path
	arrCorrectPath = Split(Environment.Value("ENV_TEST_PATH"), "\")
	ReDim Preserve arrCorrectPath(Ubound(arrCorrectPath) - 1)
	strPathVal = Join(arrCorrectPath, "\")
	Set Library = FN_GETTEST(Environment.Value("TestName"), strPathVal ) 'Get the Test Object
	Call FN_DOWNLOAD(Library, "Batch") ' Download the Batch file to the Temp location
	
End If


'*************************************************************************************************
'Type : 
'Name:	PreRun
'Objective:To check if Browser ,excel or AUT is opend
'Prerequisites:none
'Parmeters:none
'OutParmeters:none
'*************************************************************************************************
Function PreRun
	
	Set ab = Description.Create
	ab("micclass").value = "Browser"
	Set obj = Desktop.ChildObjects(ab)
	
	If obj.Count <> 0 Then
		
		Msgbox "Broswer is open. Please close it and try again."
		ExitTest
	End If
	
	Set Shell = CreateObject("WScript.Shell")
	Set ShellResult = Shell.Exec("TaskList")
	While Not ShellResult.StdOut.AtEndOfStream
		If Instr(Ucase(ShellResult.StdOut.ReadLine), UCase("excel")) Then
			Msgbox "Excel is open, Please close it and try again"
			ExitTest	
		End If
			If Instr(Ucase(ShellResult.StdOut.ReadLine), UCase("Nexus")) Then
				Msgbox "SMI NEXUS is open, please close it and try again."
				ExitTest
			End If
	Wend
	
	Set cd = Description.Create
	cd("text").value = "Securitis Market Interface Desktop"
	Set oobj = Desktop.ChildObjects(cd)
	
	If oobj <> 0 Then
		Msgbox "SMI ESES is open, please close it and try again."
		ExitTest
	End If
End Function


