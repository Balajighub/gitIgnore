
'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:		fRptWriteReport
'	Objective							:		To write the test case results in html file
'	Input Parameters					:		strResult -  Pass/Fail , 'strStepName -  Sescription of the step,'strExpected -  Expected Results
'	Output Parameters					:		Nil
'	Date Created						:		
'	UFT Version							:		UFT 15
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************
Dim iSNO, iTestCaseNumber
Public Function fRptWriteReport(ByVal strResult, ByVal strStepName , ByVal strExpected)
	
	'Declaring variables
	'Dim iSNO
	Dim objFso
	Dim objFolder
	Dim gstrResultsFolder
	Dim gstrTestcasesPath
	Dim objFile
	Dim status 
	Dim html
	Dim link
	Dim arrPath
	Dim strResources
	
	
	If Instr(Environment("TestDir"),"TestScenarios")= 0 Then
		actionName=  Environment("ActionName")
		arrResource = Split(Environment("TestDir"),"T-Codes\"&actionName) 
		strResources = arrResource(0) & "Resources"
	Else
		arrResource = Split(Environment("TestDir"),"TestScenarios") 
		strResources = arrResource(0) & "Resources"
	End If
	
	'Verify if the results forlder exists -  if not create the same
	Set objFso = CreateObject("Scripting.FileSystemObject")
	If not objFso.FolderExists(gstrProjectResultPath) Then
		Set objFolder=objFso.CreateFolder(gstrProjectResultPath)
	End If	
	If gstrFolderName=Empty Then
	    fRptFoldername
	End If
	 
	gstrResultsFolder=gstrProjectResultPath&"\"& gstrFolderName 
	If not objFso.FolderExists(gstrResultsFolder) Then
	Set objFolder=objFso.CreateFolder(gstrResultsFolder)
	End If
	
	'Copying the logos
	strLogos=gstrResultsFolder&"\Logos"
	If not objFso.FolderExists(strLogos) Then
		Set objFolder=objFso.CreateFolder(strLogos)
		objFso.CopyFile strResources&"\Pass.png",strLogos&"\"
		objFso.CopyFile strResources&"\Fail.png",strLogos&"\"
		objFso.CopyFile strResources&"\Cigniti.png",strLogos&"\"
		objFso.CopyFile strResources&"\Client.png",strLogos&"\"
		objFso.CopyFile strResources&"\PassWithScr.png",strLogos&"\"
	End If   
	
	gstrTestcasesPath=gstrResultsFolder&"\Testcases"
	If not objFso.FolderExists(gstrTestcasesPath) Then
		Set objFolder=objFso.CreateFolder(gstrTestcasesPath)
	End If
	
	If not objFso.FileExists(gstrTestcasesPath&"\"&Environment.Value("TestName")&".html") Then
		iSNO = 1
		Set  objFile = objFso.CreateTextFile(gstrTestcasesPath&"\"&Environment.Value("TestName")&".html",true, false)  
		objFile.writeline  "<html>" & VBNewLine
		objFile.writeline  "<head> " & VBNewLine
		objFile.writeline  "<style type=""text/css"">.passed{display: table-row; background-color: #E1E1E1; border: 1px solid #4D7C7B; color: #000000; font-size: 0.75em; td, th { padding: 5px; border: 1px solid #4D7C7B; text-align: inherit /; } th.Logos { padding: 5px; border: 0px solid #4D7C7B; text-align: inherit /;} td.justified { text-align: Left; } td.pass { font-weight: bold; color: green; } </style>" & VBNewLine
		objFile.writeline  "<style type=""text/css"">.failed{display: table-row;background-color: #FFFFFF; color: #000000; font-size: 0.7em; display: table-row;} </style>" & VBNewLine&"<style type=""text/css"">.notvisible{display: None; </style>"& VBNewLine &"<meta charset='UTF-8'> "
		objFile.writeline  "<title>Detailed Results Report</title>"& VBNewLine
		objFile.writeline  "<style type='text/css'>"& VBNewLine
		objFile.writeline  "body { background-color: #FFFFFF; font-family: Verdana, Geneva, sans-serif; text-align: center; } small { font-size: 0.7em; } table { box-shadow: 9px 9px 10px 4px #BDBDBD;border: 0px solid #4D7C7B; border-collapse: collapse; border-spacing: 0px; width: 1000px; margin-left: auto; margin-right: auto; } tr.heading { background-color: #041944;color: #FFFFFF; font-size: 0.7em; font-weight: bold; background:-o-linear-gradient(bottom, #999999 5%, #000000 100%);background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #999999), color-stop(1, #000000));background:-moz-linear-gradient( center top, #999999 5%, #000000 100%);filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#999999, endColorstr=#000000); background: -o-linear-gradient(top,#999999,000000);} tr.subheading { background-color: #FFFFFF; color: #000000; font-weight: bold; font-size: 0.7em; text-align: justify; } tr.section { background-color: #A4A4A4; color: #333300; cursor: pointer; font-weight: bold; font-size: 0.7em; text-align: justify; background:-o-linear-gradient(bottom, #56aaff 5%, #e5e5e5 100%); background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #56aaff), color-stop(1, #e5e5e5));background:-moz-linear-gradient( center top, #56aaff 5%, #e5e5e5 100%);filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#56aaff, endColorstr=#e5e5e5); background: -o-linear-gradient(top,#56aaff,e5e5e5);} tr.subsection { cursor: pointer; } td, th { padding: 5px; border: 1px solid #4D7C7B; text-align: inherit /; } th.Logos { padding: 5px; border: 0px solid #4D7C7B; text-align: inherit /;} " & VBNewLine
		objFile.writeline  "</style>"& VBNewLine
		objFile.writeline  "</head>" & VBNewLine
		objFile.writeline  "<body>" & VBNewLine & "</br>"
		objFile.writeline  "<table id='Logos'> " & VBNewLine 
		objFile.writeline  "<colgroup>" & VBNewLine 
		objFile.writeline  "<col style='width: 25%' />" & VBNewLine 
		objFile.writeline  "<col style='width: 25%' />" & VBNewLine 
		objFile.writeline  "<col style='width: 25%' />" & VBNewLine 
		objFile.writeline  "<col style='width: 25%' />" & VBNewLine 
		objFile.writeline  "</colgroup>" & VBNewLine
		objFile.writeline  "<thead>"& VBNewLine
		objFile.writeline  "<tr class='content'>" & VBNewLine
		objFile.writeline "<th class ='Logos' colspan='2' > <img align ='left' src='..\Logos\Clientlogo.png ' height=60 width=140></img> </th>"
		objFile.writeline  "<th class = 'Logos' colspan='2' > <img align ='right' src= '..\Logos\Companylogo.png' height=60 width=140></img></th> </tr> " & VBNewLine
		objFile.writeline  "</thead>" & VBNewLine
		objFile.writeline  "</table><table id='header'> " & VBNewLine
		objFile.writeline  "<colgroup> <col style='width: 25%' /> " & VBNewLine
		objFile.writeline  "<col style='width: 25%' /> " & VBNewLine
		objFile.writeline  "<col style='width: 25%' /> " & VBNewLine
		objFile.writeline  "<col style='width: 25%' /> " & VBNewLine
		objFile.writeline  "</colgroup>" & VBNewLine
		objFile.writeline  "<thead>" & VBNewLine
		objFile.writeline  "<tr class='heading'> " & VBNewLine
		objFile.writeline  "<th colspan='4' style='font-family:Copperplate Gothic Bold; font-size:1.4em;'> **"& Environment.Value("TestName") & " **</th>" & VBNewLine
		objFile.writeline  "</tr> " & VBNewLine
		iCurrentTime = Now()
		objFile.writeline  "<tr class='subheading'>" & VBNewLine
		objFile.writeline  "<th>&nbsp;Date&nbsp;&&nbsp;Time&nbsp;&nbsp;</th> " & VBNewLine
		objFile.writeline  "<th>"& DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & "</th>" & VBNewLine
		objFile.writeline  "<th>&nbsp;Operating&nbsp;System&nbsp;:&nbsp;</th>" & VBNewLine
		objFile.writeline  "<th> "& Environment.Value("OS") & "</th> " & VBNewLine
		objFile.writeline  "</tr> " & VBNewLine
		objFile.writeline  "<tr class='subheading'>" & VBNewLine
		'objFile.writeline  "<th>&nbsp;SAP&nbsp;Server&nbsp;:&nbsp;</th>" & VBNewLine
		'objFile.writeline  "<th> "& Environment.Value("SAPSERVER") & "</th> " & VBNewLine
		'objFile.writeline  " <th>&nbsp;Executed&nbsp;on&nbsp;:&nbsp;</th>" & iCurrentTime & VBNewLine
		'objFile.writeline  "<th>" & Environment.Value("LocalHostName") & "</th> " & VBNewLine
		objFile.writeline  "</tr> " & VBNewLine
		objFile.writeline  "</thead>" & VBNewLine
		objFile.writeline  "</table>" & VBNewLine
		objFile.writeline  "<table id='main' cellpadding=""0"" cellspacing=""0""> " & VBNewLine
		objFile.writeline  "<Head>" & VBNewLine
		objFile.writeline  "<Body>" & VBNewLine
		objFile.writeline  "<colgroup>" & VBNewLine
		objFile.writeline  "<col style='width: 5%' /> <col style='width: 26%' /> <col style='width: 51%' /> " & VBNewLine
		objFile.writeline  "<col style='width: 8%' /> <col style='width: 10%' />" & VBNewLine
		objFile.writeline  "</colgroup>"
		objFile.writeline  "<thead>"
		objFile.writeline  "<tr class='heading'>"
		objFile.writeline  "<th>S.No</th> "
		objFile.writeline  "<th>Step"
		objFile.writeline  "<INPUT id=""txtStepValue""  onchange=""filterStatus()"">"
		objFile.writeline  "</th> "
		objFile.writeline  "<th>Details"
		objFile.writeline  "<INPUT id =""txtDetailsValue""  onchange=""filterDetails()"">"
		objFile.writeline  "</th> "
		objFile.writeline  "<th> Status"
		objFile.writeline  "<select id=""filter"" onchange=""filter()"">"
		objFile.writeline  "<option value=""all"">All</option> "
		objFile.writeline  "<option value=""passed"">Passed</option>"
		objFile.writeline  "<option value=""failed"">Failed</option>"
		objFile.writeline  "</select>"
		objFile.writeline  "</th>"
		objFile.writeline  "<th>Time</th>"
		objFile.writeline  "</tr> "
		objFile.WriteBlankLines(5)
		objFile.writeline  "<script type=""text/javascript"">" & VBNewLine
		objFile.writeline  "function filter()" & VBNewLine
		objFile.writeline  "{" & VBNewLine
		objFile.writeline  "if(document.getElementById(""filter"").value==""passed"")" & VBNewLine
		objFile.writeline  "{" & VBNewLine
		objFile.writeline  "document.getElementsByTagName(""style"")[0].textContent = "".passed{display: table-row;background-color: #E1E1E1; border: 1px solid #4D7C7B; color: #000000; font-size: 0.75em;}"";" & VBNewLine
		objFile.writeline  "document.getElementsByTagName(""style"")[1].textContent = "".failed{display: none;}"";" & VBNewLine
		objFile.writeline  "}" & VBNewLine
		objFile.writeline  "else if (document.getElementById(""filter"").value==""failed"")" & VBNewLine
		objFile.writeline  "{" & VBNewLine
		objFile.writeline  "document.getElementsByTagName(""style"")[1].textContent = "".failed{display: table-row;background-color: #FFFFFF;color: #000000; font-size: 0.7em; display: table-row;}"";" & VBNewLine
		objFile.writeline  "document.getElementsByTagName(""style"")[0].textContent = "".passed{display: none;}"";" & VBNewLine
		objFile.writeline  "}" & VBNewLine
		objFile.writeline  "else" & VBNewLine
		objFile.writeline  "{" & VBNewLine
		objFile.writeline  "document.getElementsByTagName(""style"")[0].textContent = "".passed{display: table-row;background-color: #E1E1E1; border: 1px solid #4D7C7B; color: #000000; font-size: 0.75em;}"";" & VBNewLine
		objFile.writeline  "document.getElementsByTagName(""style"")[1].textContent = "".failed{display: table-row;background-color: #FFFFFF;color: #000000; font-size: 0.7em; display: table-row;}"";" & VBNewLine
		objFile.writeline  "}" & VBNewLine
		objFile.writeline  "}" & VBNewLine
		objFile.writeline  "</script>" & VBNewLine
		objFile.writeline  "<script type=""text/javascript"">"
		objFile.writeline  "function filterStatus()"
		objFile.writeline  "{"
		objFile.writeline  "searchtext = (document.getElementById(""txtStepValue"").value).toLowerCase();"
		objFile.writeline  "if(searchtext!="""")"
		objFile.writeline  "{"
		objFile.writeline  "var rowIndex = 0; // rowindex, in this case the first row of your table"
		objFile.writeline  "var table = document.getElementById('main'); // table to perform search on"
		objFile.writeline  "var row = table.getElementsByTagName(""tr"");"
		objFile.writeline  "irowcount = row.length"
		objFile.writeline  "for (i = 1; i < row.length; i++) {"
		objFile.writeline  "status = (row[i].getElementsByTagName(""td"")[1].textContent).toLowerCase();"
		objFile.writeline  "if (status.indexOf(searchtext) == -1) "
		objFile.writeline  "{"
		objFile.writeline  "row[i].className = 'content notvisible'"
		objFile.writeline  "}}}"
		objFile.writeline  "else {"
		objFile.writeline  "window.location.reload()"
		objFile.writeline  "}}"
		objFile.writeline  "</script>"
		objFile.writeline  "<script type=""text/javascript"">"
		objFile.writeline  "function filterDetails()"
		objFile.writeline  "{"
		objFile.writeline  "searchtext = (document.getElementById(""txtDetailsValue"").value).toLowerCase();"
		objFile.writeline  "if(searchtext!="""")"
		objFile.writeline  "{"
		objFile.writeline  "var rowIndex = 0; // rowindex, in this case the first row of your table"
		objFile.writeline  "var table = document.getElementById('main'); // table to perform search on"
		objFile.writeline  "var row = table.getElementsByTagName(""tr"");"
		objFile.writeline  "for (i = 1; i < row.length; i++) {"
		objFile.writeline  "Details = (row[i].getElementsByTagName(""td"")[2].textContent).toLowerCase();;"
		objFile.writeline  "if (Details.indexOf(searchtext) == -1) "
		objFile.writeline  "{"
		objFile.writeline  "row[i].className = 'content notvisible'"
		objFile.writeline  "}}}"
		objFile.writeline  "else {"
		objFile.writeline  "window.location.reload()"
		objFile.writeline  "}}"
		objFile.writeline  "</script>"
	Else
		Set objFile=objFso.OpenTextFile(gstrTestcasesPath&"\"&Environment.Value("TestName")&".html", 8,TRUE)    
	End If
		
	Select Case ucase(strResult)
		Case "PASS" 
			Reporter.ReportEvent micPass , strStepName , strActual
			objFile.WriteLine "<tr class='content passed' ><td>" & iSNO & "</td> "
			objFile.WriteLine "<td class='justified'>" & strStepName &"</td>"
			objFile.WriteLine "<td class='justified'>" & strExpected & "</td>"
			objFile.WriteLine "<td class='Pass' align='center'><img  src='" & "..\Logos\Pass.png' width='18' height='18'/></td> "
			iCurrentTime = Now()
			objFile.WriteLine "<td><small>" & DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & ":" & Second(iCurrentTime)& "</small></td> </tr>"
			fRptReportLog strStepName,strExpected,"Pass"						
		
		Case "FAIL"	
			Reporter.ReportEvent micFail , strStepName , strActual
			objFile.WriteLine "<tr class='content failed' ><td>" & iSNO & "</td> "
			objFile.WriteLine "<td class='justified'>" & strStepName &"</td>"
			objFile.WriteLine "<td class='justified'>" & strExpected & "</td> "
			link = fRptScreenCapture()
			objFile.WriteLine "<td class='Fail' align='center'><a href="& link &">"
			iCurrentTime = Now
			objFile.WriteLine "<img  src='" & "..\Logos\Fail.png' width='18' height='18'/></td> <td><small>" & DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & ":" & Second(iCurrentTime)& "</small></td> </tr>"
			fRptReportLog strStepName,strExpected,"Fail"
			Environment("StepFailed") = "YES"
'			DataTable("StepStatus", "Global") = "Failed"
'			RunStatus = DataTable("StepStatus", "Global")
			Environment("fRptWriteReport")=strStepName&strExpected			
		
		Case "PASSWITHSCREENSHOT"
			Reporter.ReportEvent micPass , strStepName , strActual
			objFile.WriteLine "<tr class='content passed' ><td>" & iSNO & "</td> "
			objFile.WriteLine "<td class='justified'>" & strStepName &"</td>"
			objFile.WriteLine "<td class='justified'>" & strExpected & "</td> "
			link = fRptScreenCapture()
			objFile.WriteLine "<td class='Pass' align='center'><a href="& link &">"
			iCurrentTime = Now
			objFile.WriteLine "<img  src='" & "..\Logos\PassWithScr.png' width='30' height='30'/></td> <td><small>" & DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) & ":" & Second(iCurrentTime)& "</small></td> </tr>"
			fRptReportLog strStepName,strExpected,"Pass"
	End Select
	iSNO = iSNO+1
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:		fRptScreenCapture
'	Objective							:		To capture screen and send path of image fie to called function 
'	Input Parameters					:		Nil
'	Output Parameters					:		Nil
'	Date Created						:		
'	UFT Version							:		15
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************
Function fRptScreenCapture()	
	
	'Declaring the variables
	Dim objFso
	Dim strResultsPath
	Dim strScreenshotPath
	Dim objFolder
	Dim strImagePath
	Dim strFilePath
	Dim strImagelinkPath
	Dim objDesktop
	
	'Verify the screenshot folder existence and if not create one
	Set objFso = CreateObject("Scripting.FileSystemObject")
	strScreenshotPath=gstrProjectResultPath&"\"& gstrFolderName &"\Screenshot"
	If not objFso.FolderExists(strScreenshotPath) Then
		Set objFolder=objFso.CreateFolder(strScreenshotPath)
	End If
	
	strImagePath="\Screenshot"&Replace(Replace(Replace(now(),":","_"),"/","_")," ","_") &".png"
	strFilePath=strScreenshotPath&strImagePath
	strImagelinkPath="..\Screenshot"&strImagePath
	Set objDesktop = Desktop
	
	'Capture the Desktop
	objDesktop.capturebitmap strFilePath ,  true
	
	'Add the Captured Screen shot to the Results file
	fRptScreenCapture=strImagelinkPath
	
End Function 


'******************************************************************************************************************************************************************************************************************************************
'	Sub Name							:		fRptWriteResultsSummary
'	Objective							:		To create summary report of executed test cases
'	Input Parameters					:		Nil
'	Output Parameters					:		Nil
'	Date Created						:		
'	UFT Version							:		UFT 15
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************
Function fRptWriteResultsSummary()
	
	'Declaring Variables
	Dim strResultsPath
	Dim objSummary
	Dim objFilesummary
	Dim gstrResultsFolder
	Dim objFSO
	Dim objFolder
	Dim objFiles
	Dim intCount
	Dim intFailCount
	Dim intPassCount
	Dim objFile
	Dim SummaryChart
	Dim html
	Dim objWShell
	Dim intFailedScriptPercentage
	Dim intPassedSrciptPercentage
	
	arrResource = Split(Environment("TestDir"),"TestScenarios") 
	strResources = arrResource(0) & "Resources"
	SummaryChart =gstrProjectResultPath&"\"&gstrFolderName &"\SummaryChart.html"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(gstrProjectResultPath & "\" & gstrFolderName & "\Testcases")
	Set objFiles = objFolder.Files
	Set objFile = objFSO.CreateTextFile(SummaryChart, true, false)
		
	objFile.writeline "<table id='Logos'> <colgroup> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> </colgroup> "
	objFile.writeline "<thead>  <tr class='content'> <th class ='Logos' colspan='2' > <img align ='left' src='.\Logos\Clientlogo.png ' height=60 width=140></img> </th>"
	objFile.writeline "<th class = 'Logos' colspan='2' > <img align ='right' src='.\Logos\Companylogo.png ' height=60 width=140></img> </th> </tr> </thead> </table> "		
	objFile.writeline "<html> <head> <script src='http://ajax.googleapis.com/ajax/libs/jquery/1.7.1/jquery.min.js' type='text/javascript'></script>"
	objFile.writeline "<script src='/js/highcharts.js' type='text/javascript'></script><script src='http://code.highcharts.com/highcharts.js'></script>"
	objFile.writeline "<script src='http://code.highcharts.com/highcharts-3d.js'></script><script src='http://code.highcharts.com/modules/exporting.js'></script>"
	objFile.writeline "<meta charset='UTF-8'> <title> Execution Summary Report</title><style type='text/css'>body {background-color: #FFFFFF; "
	objFile.writeline "font-family: Verdana, Geneva, sans-serif; text-align: center; } small { font-size: 0.7em; } table { box-shadow: 9px 9px 10px 4px #BDBDBD;"
	objFile.writeline "border: 0px solid #4D7C7B;border-collapse: collapse; border-spacing: 0px; width: 1000px; margin-left: auto; margin-right: auto; } "
	objFile.writeline "tr.heading { background-color: #041944;color: #FFFFFF; font-size: 0.7em; font-weight: bold; "
	objFile.writeline "background:-o-linear-gradient(bottom, #999999 5%, #000000 100%); "
	objFile.writeline "background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #999999), color-stop(1, #000000) );"
	objFile.writeline "background:-moz-linear-gradient( center top, #999999 5%, #000000 100% );"
	objFile.writeline "filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#999999, endColorstr=#000000); "
	objFile.writeline "background:-o-linear-"
	objFile.writeline "gradient(top,#999999,000000);} tr.subheading { background-color: #6A90B6;color: #000000; font-weight: bold; font-size: 0.7em; "
	objFile.writeline "text-align:justify; } tr.section { background-color: #A4A4A4; color: #333300; cursor: pointer; font-weight: bold;font-size: 0.8em; "
	objFile.writeline "text-align: justify;"
	objFile.writeline "background:-o-linear-gradient(bottom, #56aaff 5%, #e5e5e5 100%); "
	objFile.writeline "background:-webkit-gradient( linear, left top, left bottom,color-stop(0.05, #56aaff), color-stop(1, #e5e5e5) );"
	objFile.writeline "background:-moz-linear-gradient( center top, #56aaff 5%, #e5e5e5 100% );"
	objFile.writeline "filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#56aaff, endColorstr=#e5e5e5);"
	objFile.writeline "background:-o-linear-gradient(top,#56aaff,e5e5e5);} tr.subsection { cursor: pointer; } "
	objFile.writeline "tr.content { background-color: #FFFFFF; color:#000000; font-size: 0.7em; display: table-row; } "
	objFile.writeline "tr.content2 { background-color:#;E1E1E1border: 1px solid #4D7C7B;color: #000000; "
	objFile.writeline "font-size: 0.7em; display: table-row; } td, th { padding: 5px; border: 1px solid #4D7C7B; text-align: inherit/; } th.Logos {" 
	objFile.writeline "padding: 5px; "
	objFile.writeline "border: 0px solid #4D7C7B; text-align: inherit /;} td.justified { text-align: justify; } td.pass {font-weight: bold; color: green;"
	objFile.writeline "}" 
	objFile.writeline "td.fail { font-weight: bold; color: red; } td.done, td.screenshot { font-weight: bold; color: black; } "
	objFile.writeline "td.debug { font-weight: bold;color: blue; } td.warning { font-weight: bold; color: orange; } </style> </head> "
	objFile.writeline "<body> </br><table id='header'> "
	objFile.writeline "<colgroup> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> " 
	objFile.writeline "<col style='width: 25%' /> </colgroup> <thead> <tr class='heading'> <th colspan='4' style='font-family:Copperplate Gothic Bold;" 
	objFile.writeline "font-size:1.4em;'> AUTODESK - Automation Execution Result Summary </th> </tr> <tr class='subheading'>   "
	'objFile.writeline "<th>&nbsp;Date&nbsp;&&nbsp;Time&nbsp;IST</th> <th>&nbsp;&nbsp;"& Now &"</th>"
	iCurrentTime = Now()
	objFile.writeline "<th>&nbsp;Date&nbsp;&&nbsp;Time&nbsp;IST</th> <th>&nbsp;&nbsp;"& DatePart("d", iCurrentTime) & "-" & MonthName(Month(iCurrentTime), True) & "-" & DatePart("yyyy", iCurrentTime) & Space(1) & Hour(iCurrentTime) & ":" & Minute(iCurrentTime) &"</th>"
	
    objFile.writeline "<th>&nbsp;&nbsp;"& "Environment : "& Environment.Value("Environment") & "</th>"
	objFile.writeline "<th>&nbsp;&nbsp;"& "Test System - "& Environment.Value("LocalHostName") & "</th><th></th></tr></thead></table>"
	objFile.writeline "<table id='main'> <colgroup> <col style='width: 10%' /> <col style='width: 40%' /> <col style='width: 20%' /> <col style='width:" 
	objFile.writeline "30%' /> </colgroup> "
	objFile.writeline "<thead> <tr class='heading'> <th>S.NO</th> <th>Test Case</th> <th>Status</th> <th>Time</th> </tr> </thead> <tbody>"
	Set objFile = Nothing
			
	intCount=0
	intFailCount=0
	intPassCount=0
	iTestCaseNumber = 0
	iTotalExecutionTime = 0		
	For Each Item In objFiles
	   If LCase(Right(Item.Name, 5)) = ".html" Or LCase(Right(Item.Name, 4)) = ".htm" Then
		  Set objFileDetailedReport = objFSO.OpenTextFile(Item.Path, 1, False)
			 strText = objFileDetailedReport.readAll()
			 Set objReg = New RegExp
			 objReg.Pattern = "[\d]+-[a-zA-Z]+-[\d]+ [\d]+:[\d]+:[\d]+"			 
			 objReg.Global = True
			 Set objMatches =  objReg.Execute(strText)
			 iStepCount = objMatches.Count
			 iStartTime = objMatches(0).Value
			 iEndTime = objMatches(iStepCount-1).Value
			 iExecutionTime = Round((CDbl(DateDiff("s",CDate(iStartTime),CDate(iEndTime)))/60), 2) &" Minutes"				  
			 iTotalExecutionTime = iTotalExecutionTime + Round((CDbl(DateDiff("s",CDate(iStartTime),CDate(iEndTime)))/60), 2)
			 If Instr(strText,"Fail.png") > 0 Then
				fRptAddTCsInSummary Item.Name, "FAIL", iExecutionTime
				intFailCount = intFailCount +1
			 Else
				fRptAddTCsInSummary Item.Name,"PASS" ,iExecutionTime
				intPassCount = intPassCount +1 
			 End If
	   End If
	   Set objFileDetailedReport = Nothing
	Next 
	
	intCount = intFailCount + intPassCount
	intFailedScriptPercentage = Cint(100*(intFailCount/intCount))
	intPassedSrciptPercentage = Cint(100*(intPassCount/intCount))	
			
	Set objFile=objFSO.openTextFile(SummaryChart, 8, True)
    strhtml="<h2 align=center><img src = https://chart.googleapis.com/chart?cht=p3&amp;chtt=ProductSuite&amp;chl=Pass-"&intPassedSrciptPercentage&"%--("&intPassCount&")|Fail-"&intFailedScriptPercentage&"%--("&intFailCount&")&amp;chs=500x250&amp;chd=t:"&intPassedSrciptPercentage&","&intFailedScriptPercentage&"&amp;chco=00FF00|FF0000&amp;height=700;width=900  /></h2>"	
	objFile.writeline strhtml	
	objFile.writeline "</table> <table id='footer'> <colgroup> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> <col style='width: 25%' /> </colgroup> "
	objFile.writeline "<tfoot> <tr class='heading'>	<th colspan='4'>Total Duration (Including Report Creation) : "&iTotalExecutionTime&" Minutes </th> </tr> <tr class='content'>"
	objFile.writeline "<td class='pass'>&nbsp;Tests passed</td>	<td class='pass'>&nbsp;"&intPassCount&"</td> <td class='fail'>&nbsp;Tests failed</td>	<td class='fail'>&nbsp; "&intFailCount&"</td> </tr> </tfoot> </table>"
	
	set objFile=nothing
	set objFSO=nothing
	set objFolder=nothing
	set objFiles=nothing 
		fUploadReport gstrProjectResultPath&"\"& gstrFolderName, gstrProjectResultPath&"\"& gstrFolderName
		Call fGetDataforTestRail(trRID,trTID,gstrProjectResultPath&"\"& gstrFolderName)
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Sub Name							:		fRptAddTCsInSummary
'	Objective							:		To add invidual TC's into summary report 
'	Input Parameters					:		Nil
'	Output Parameters					:		Nil
'	Date Created						:		
'	UFT Version							:		15.0
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 					
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************
Function fRptAddTCsInSummary(tname, tresult, iExecutionTime)
	
	'Declaring Variables
	Dim strResultsPath
	Dim gstrResultsFolder
	Dim objFileSummary
	Dim objSummary
	
	strResources=gTestDir &"Resources"
	SummaryChart =gstrProjectResultPath&"\"&gstrFolderName &"\SummaryChart.html"
	Set objFileSummary=CreateObject("scripting.filesystemobject")
	Set objSummary=objFileSummary.openTextFile(SummaryChart, 8, True)
	iTestCaseNumber = iTestCaseNumber+1
	
	If StrComp(tresult,0,1) = 0 or StrComp(tresult,"PASS",1) = 0  then		 
		objSummary.writeline "<tr class='content2'><td class='justified'><font color='#153e7e' size='1' face='arial'><b>"&iTestCaseNumber&"</b>"
		objSummary.writeline "</font></td><td class='justified'> <a href=.\TestCases\" & tname & ">"& Split(tName,".")(0) &"</a></td>"
		objSummary.writeline "<td class='justified'>Pass</td><td class='justified'>"&iExecutionTime&"</td></tr></tbody>" 
	
	Elseif  StrComp(tresult,1,1) = 0 or StrComp(tresult,"FAIL",1) = 0  then									
		objSummary.writeline "<tr class='content2'><td class='justified'><font color='#153e7e' size='1' face='arial'><b>"&iTestCaseNumber&"</b>"
		objSummary.writeline "</font></td><td class='justified'> <a href=.\TestCases\" & tname & ">"& Split(tName,".")(0) &"</a></td>"
		objSummary.writeline "<td class='justified'>Fail</td><td class='justified'>"&iExecutionTime&"</td></tr></tbody>" 
	End If
	
	Set objSummary = Nothing
	Set objFileSummary = Nothing
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Sub Name							:		fRptReportLog
'	Objective							:		To log the step details in log file 
'	Input Parameters					:		strStepName, ByVal strExpected,ByVal strStatus
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************

Public Function fRptReportLog(ByVal strStepName, ByVal strExpected,ByVal strStatus)
	
	'Declare Variables
	Dim objFilesys
	Dim gstrLogFloder
	Dim objFile
		
	set objFilesys = CreateObject("Scripting.FileSystemObject")
	gstrLogFloder =gstrProjectResultPath&"\"&gstrFolderName &"\Logs"

	If objFilesys.FolderExists(gstrLogFloder)= False Then
		objFilesys.CreateFolder(gstrLogFloder)
	End If
	
	If objFilesys.FileExists(gstrLogFloder&"\"&Environment.Value("TestName")&".txt")= false Then
		Set objFile=objFilesys.CreateTextFile(Trim(gstrLogFloder)&"\"&Environment.Value("TestName")&".txt")
		objFile.WriteLine "Test Name"&vbtab & "Expected" & vbtab & "Status" & vbtab & "Time"
		Set objFile=Nothing
	End if
	
	Set objFile = objFilesys.OpenTextFile(gstrLogFloder&"\"&Environment.Value("TestName")&".txt",8,True)
	objFile.WriteLine Environment.Value("TestName")& vbTab & strExpected &  vbTab & Ucase(strStatus)& vbtab& Now 
	Set objFile = Nothing
	Set objFilesys = Nothing
	gstrStatusPath=gstrLogFloder
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Sub Name							:		fRptFoldername
'	Objective							:		To make a result folder name with proper date formats
'	Input Parameters					:		Nil
'	Output Parameters					:		Nil
'	Date Created						:		
'	UFT Version							:		15.0
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************
sub fRptFoldername()

	Dim dDate
	Dim strdate
	Dim Filename
	
	dDate=Now()         
	Call fEnvironmentVarExists("FolderName")
	'Call fEnvironmentVarExists()
	If gEnvStatus = False  Then     	
		   	Foldername="Autodesk_Automation_Results_"&Month(dDate)&"-"&Day(dDate)&"-"&Year(dDate)&"-"&hour(dDate)&"-"&Minute(dDate)
	    gstrFolderName=Foldername 
	Else 
	  Set UftApplication = CreateObject("QuickTest.Application")
		  gstrFolderName = UftApplication.Test.Environment.Value("FolderName")
	End If
	     
End sub


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:		fCopyTestData
'	Objective							:		To Copy Test Data file from Test Data to Results folder
'	Input Parameters					:		Nil
'	Output Parameters					:		Nil
'	Date Created						:		
'	UFT Version							:		UFT 15.0
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************
Public Function fCopyTestData

	Set objFso = CreateObject("Scripting.FileSystemObject")
	gstrResultsFolder=gstrProjectResultPath&"\"& gstrFolderName 
	arrTestdata = Split(Environment("TestDir"),"TestScenarios") 
	strTestDataPath = arrTestdata(0) & "TestData"
	strResultsTestData=gstrResultsFolder & "\TestData"	
	
	If Not objFso.FolderExists(strResultsTestData) Then
		Set objFolder=objFso.CreateFolder(strResultsTestData)
	End If	
	If objFso.FileExists(strTestDataPath & "\" & gScenarioName &"_Testdata.xls") = True Then
		objFso.CopyFile strTestDataPath & "\" & gScenarioName &"_Testdata.xls",strResultsTestData &"\"
	End if
	
	Set objFile=Nothing
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:		fEnvironmentVarExists
'	Objective							:		Verify Env Varibale in UFT
'	Input Parameters					:		Nil
'	Output Parameters					:		Nil
'	Date Created						:		
'	UFT Version							:		UFT 15.0
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'***************************************************************************************************************************************************************	
Public Function fEnvironmentVarExists(strFolderName)
	
	
	On Error Resume Next
	Dim strVal
	Err.Clear()
	Set UftApplication = CreateObject("QuickTest.Application")   
    strVal = UftApplication.Test.Environment.Value(strFolderName)
	
	If Err.Number <> 0 Then
		fEnvironmentVarExists = False
	Else
		fEnvironmentVarExists = True
	End If
	
	gEnvStatus = fEnvironmentVarExists	
	On Error GoTo 0
	
	
'	On Error Resume Next
'	Dim strVal
'	Err.Clear()
'	Set UftApplication = CreateObject("QuickTest.Application")   
'    strVal = UftApplication.Test.Environment.Value("FolderName")
'	
'	If Err.Number <> 0 Then
'		fEnvironmentVarExists = False
'	Else
'		fEnvironmentVarExists = True
'	End If
'	
'	gEnvStatus = fEnvironmentVarExists	
'	On Error GoTo 0

End Function  


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fALMFailedStep
'	Objective							:					Used to failed ALM step
'	Input Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					
'	QTP Version							:					UFT 15.0
'	QC Version							:					
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'*****************************************************************************************************************************************************************************************************************************************		
Public Function fALMFailedStep()

	'Check the step status
	If Environment("StepFailed") = "YES" Then
        TDrun.Status = "Failed"
        TDrun.Field("RN_STATUS") = "Failed"
        TDrun.Post
        TDrun.Refresh
        Environment("StepFailed") = "NO"
	End If 
	Environment("StepFailed") = "YES"          
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fWriteExcelGraphicalReport
'	Objective							:					Updates the fields in the Graphical Report in Excel Format
'	Input Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					
'	QTP Version							:					UFT 15.0
'	QC Version							:					
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'*****************************************************************************************************************************************************************************************************************************************		
Public Function fWriteGraphicalReport(gstrFile)

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set res_file = fso.OpenTextFile(gstrFile,8)
	intPassPercentage  = (Environment("PASSCOUNT")/Environment("MASTERCOUNT"))*100
	intFailPercentage  = (Environment("FAILCOUNT")/Environment("MASTERCOUNT"))*100
	res_file.Write "<table width=1200 bgcolor=white border=50 bordercolor=white><td align=center><font color=White size=4 style=Verdhana><B><B></td ></tr></table>"
	res_file.Write "<table width=1200 bgcolor=Blue border=1 bordercolor=Red><td align=center><font color=White size=4 style=Verdhana><B>Graphical Summary Report<B></td ></tr></table>"
	res_file.Write "<TABLE width=1200 bgcolor=White BORDER=2 bordercolor=Gray CELLSPACING=4 CELLPADDING=3>"
	res_file.Write "<TR bgcolor=lightgrey><TH align=Center>Status</TH><TH>Count</TH><TH align=center>Percentage (%)</TH></TR><TR><TD align=center><font color=Green size=3 style=Verdhana>Pass</TD><TD align=center><font color=Green size=3 style=Verdhana>"&Environment("PASSCOUNT")&"</TD><TD valign=middle>"
	res_file.Write "<TABLE><TR><TD bgcolor=darkgreen><IMG SRC='/gifs/s.gif' width= "&intPassPercentage&" height=5></TD>"
	res_file.Write "<TD><FONT SIZE=1>"&intPassPercentage&"</FONT></TD></TR></TABLE></TD></TR></TD></TR>"
	
	res_file.Write "<TR><TD align=center><font color = Red size=3 style=Verdhana>Fail</TD><TD align=center><font color=Red size=3 style=Verdhana>"&Environment("FAILCOUNT")&"</TD><TD valign=middle>"
	res_file.Write "<TABLE><TR><TD bgcolor=darkred><IMG SRC='/gifs/s.gif' width="&intFailPercentage&" height=5></TD>"
	res_file.Write "<TD><FONT SIZE=1>"&intFailPercentage&"</FONT></TD></TR>"
	res_file.Write "</TABLE></TD></TR></TD></TR></TABLE>"
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fWriteSummaryStatus
'	Objective							:					Used to Write the Status whether Pass/Fail for each and even Scenario
'	Input Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					
'	QTP Version							:					UFT 15.0
'	QC Version							:					
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'*****************************************************************************************************************************************************************************************************************************************		
Public Function fWriteSummaryStatus()

	If Environment("ERRORSUMMARYCOUNT") <> 1Then
		Environment("TRANSACTIONSTATUS") = False
		Environment("FAILCOUNT") = Environment("FAILCOUNT") + 1
		Call fWriteSummaryResults()
	Else
		Environment("TRANSACTIONSTATUS") = True
		Environment("PASSCOUNT") = Environment("PASSCOUNT") + 1
		Call fWriteSummaryResults()
	End If
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fWriteGraphicalSummary
'	Objective							:					Used to Write Graphical Summary in HTML and Excel Format
'	Input Parameters					:					NIL
'	Output Parameters					:					NIL
'	Date Created						:					
'	QTP Version							:					UFT 15.0
'	QC Version							:					
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'*****************************************************************************************************************************************************************************************************************************************		

Public Function fWriteGraphicalSummary()

	Call fWriteGraphicalReport(Environment("SUMMARYREPORTHTML"))
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Sub Name		 					:					fQCGetResource
'	Objective							:					Used to get resources  names from ALM
'	Input Parameters					:					strResourceName, strSaveTo 
'	Output Parameters					:					NIL
'	Date Created						:					
'	QTP Version							:					UFT 15.0
'	QC Version							:					
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fQCGetResource(strResourceName,strSaveTo)
	
	Set objResource = QCUtil.QCConnection.QCResourceFactory
    Set objResourceList = objResource.NewList("")
	For each objResource in objResourceList
		If Trim(Ucase(objResource.Name)) = Trim(Ucase(strResourceName))  Then
			objResource.DownloadResource strSaveTo, True
			Exit For
		End If
	Next	
	Set objResource = Nothing
		
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Sub Name		 					:					fDownloadResourcesFromALM
'	Objective							:					Used to download Logos & OR
'	Input Parameters					:					
'	Output Parameters					:					NIL
'	Date Created						:					
'	QTP Version							:					UFT 15.0
'	QC Version							:					
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fDownloadResourcesFromALM()

	strResources =gstrFrameWorkFolder&"\Resources"
	Set objFso = CreateObject("Scripting.FileSystemObject")  

	If not objFso.FolderExists(gstrRootFolder) Then	
		Set objFolder=objFso.CreateFolder(gstrRootFolder)
	End If

	If not objFso.FolderExists(gstrFrameWorkFolder) Then
		Set objFolder=objFso.CreateFolder(gstrFrameWorkFolder)
	End If
	If not objFso.FolderExists(gstrProjectResultPath) Then
		Set objFolder=objFso.CreateFolder(gstrProjectResultPath)
	End If
	If not objFso.FolderExists(strResources) Then
		Set objFolder=objFso.CreateFolder(strResources)
	End If	
	If not objFso.FolderExists(gstrProjectTestdataPath) Then
		Set objFolder=objFso.CreateFolder(gstrProjectTestdataPath)
	End If	

	gstrFolderName = Empty
	If gstrFolderName=Empty Then
		fRptFoldername
	End If	
	gstrResultsFolder=gstrProjectResultPath&"\"& gstrFolderName 

	If not objFso.FolderExists(gstrResultsFolder) Then
	   Set objFolder=objFso.CreateFolder(gstrResultsFolder)
	End If
		
	If  Not objFso.FileExists(strResources& "\Clientlogo.png" )Then
	   Call fQCGetResource("Clientlogo",strResources )
	End If
	If  objFso.FileExists(gstrProjectResultPath & "\AutoDesk_TestData.xls")Then
		objFso.DeleteFile(gstrProjectResultPath & "\AutoDesk_TestData.xls")
		Call fQCGetResource("AutoDesk_TestData",gstrProjectTestdataPath)
	End If  
	If  objFso.FileExists(gstrProjectResultPath & Environment("TestName") & "_TD.xls") Then
		objFso.DeleteFile(gstrProjectResultPath & Environment("TestName") & "_TD.xls")
		Call fQCGetResource(Environment("TestName") & "_TD.xls",gstrProjectTestdataPath)
	End If  
	 
	If  Not objFso.FileExists(strResources& "\Companylogo.png" )Then
		Call fQCGetResource("Companylogo",strResources)
	End If
	If  Not objFso.FileExists(strResources& "\Fail.png" )Then
		Call fQCGetResource("Fail",strResources)
	End If
	If  Not objFso.FileExists(strResources& "\Pass.png" )Then
		Call fQCGetResource("Pass",strResources)
	End If
	If  Not objFso.FileExists(strResources& "\Warning.png" )Then
		Call fQCGetResource("Warning",strResources)
	End If

	If  Not objFso.FileExists(strResources& "\PassWithScr.png" )Then
		Call fQCGetResource("PassWithScr",strResources)
	End If

	If  objFso.FileExists(gstrProjectResultPath & "\eWEB.tsr")Then
	   objFso.DeleteFile(gstrProjectResultPath & "\eWEB.tsr")
	End If   

	If  objFso.FileExists(gstrProjectResultPath & "\dWEB.tsr")Then
	   objFso.DeleteFile(gstrProjectResultPath & "\dWEB.tsr")
	End If 
	If  Not objFso.FileExists(gstrProjectResultPath & "\dWEB.tsr" )Then
	End If

	strLogos=gstrResultsFolder&"\Logos"
	gstrTestData=gstrResultsFolder&"\TestData"

	If not objFso.FolderExists(gstrTestData) Then
	   Set objFolder=objFso.CreateFolder(gstrTestData)
	End If   		
	
	gstrTestcasesPath=gstrResultsFolder&"\Testcases"
	If not objFso.FolderExists(gstrTestcasesPath) Then
	   Set objFolder=objFso.CreateFolder(gstrTestcasesPath)
	End If
		
	If not objFso.FolderExists(strLogos) Then
	   Set objFolder=objFso.CreateFolder(strLogos)
	   objFso.CopyFile strResources&"\Pass.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Fail.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Companylogo.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Clientlogo.png",strLogos&"\"
	   objFso.CopyFile strResources&"\PassWithScr.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Warning.png",strLogos&"\" 	
	End If
	
	Set objFso = Nothing
	
End Function




'------------------------------------

'******************************************************************************************************************************************************************************************************************************************
'	Sub Name		 					:					fnGetDataforTestRail
'	Objective							:					Used to write TestRail mandatory details
'	Input Parameters					:					trRID,trTID,sPath
'	Output Parameters					:					NIL
'	Date Created						:					08/05/2020
'	UFT Version							:					UFT 15.0
'	ALM Version							:					
'	Module Name							:					
'	Pre-requisites						:					NILL  
'	Created By							:					Balaji Veeravalli
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fGetDataforTestRail(trRID,trTID,sPath)
	tcStatus=fGetStatus()
		If tcStatus="FAIL" Then
		   tcStatus=5
		ElseIf tcStatus="PASS" Then
		   tcStatus=1
		Else  
			tcStatus="Script Got failed"
		End If
			Set objFSO = CreateObject("Scripting.FileSystemObject")			
				If  objFSO.FileExists(gstrTestRailtFolder&"\"&"uftLog.txt") Then
					objFSO.DeleteFile(gstrTestRailtFolder&"\"&"uftLog.txt")		
				Set oFile=objFSO.CreateTextFile(gstrTestRailtFolder&"\"&"uftLog.txt",true)
					oFile.Write "TestRail RID:"&trRID &VBNEWLine
					oFile.Write "TestRail TCID:"&trTID &VBNEWLine		
					oFile.Write "Automation Status:"&tcStatus&VBNEWLine
					oFile.Write "sPath:"&sPath&".zip"&VBNEWLine
				End If	
		
		Call fCreatebatchFile()		
'		Call fAPIExecution()
	Set objFSO = nothing
	Set oFile = nothing
End Function
Public Function fCreatebatchFile()
	Set objFSO = CreateObject("Scripting.FileSystemObject")			
	If  objFSO.FileExists(gstrTestRailtFolder&"\"&"TRBatch.bat") Then
		objFSO.DeleteFile(gstrTestRailtFolder&"\"&"TRBatch.bat")		
	Set oFile=objFSO.CreateTextFile(gstrTestRailtFolder&"\"&"TRBatch.bat",true)
		oFile.Write "cd "&gstrTestRailtFolder &VBNEWLine
		oFile.Write "java -jar TRstatusUpdate.jar"&VBNEWLine							
	End If	
	Set objFSO = nothing
	Set oFile = nothing
End Function
'******************************************************************************************************************************************************************************************************************************************
'	Sub Name		 					:					fnAPIExecution
'	Objective							:					Used to run API .bat file
'	Input Parameters					:					
'	Output Parameters					:					NIL
'	Date Created						:					08/05/2020
'	UFT Version							:					UFT 15.0
'	AL Version							:					
'	Module Name							:					
'	Pre-requisites						:					NILL  
'	Created By							:					Balaji Veeravalli
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
'Public Function fAPIExecution()
'	Set objShell = CreateObject("WScript.Shell")
'		If CreateObject("Scripting.FileSystemObject").FileExists(gstrTestRailtFolder + "uftLog.txt") Then
'				objShell.run gstrTestRailtFolder +"TRBatch.bat"
'		End If
'	Set objShell=nothing	
'End Function

'******************************************************************************************************************************************************************************************************************************************
'	Sub Name		 					:					fGetStatus
'	Objective							:					Used to get test script status
'	Input Parameters					:					
'	Output Parameters					:					NIL
'	Date Created						:					08/05/2020
'	UFT Version							:					UFT 15.0
'	AL Version							:					
'	Module Name							:					
'	Pre-requisites						:					NILL  
'	Created By							:					Balaji Veeravalli
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fGetStatus()
	'strLogpath
	Dim strLogpath
	Dim objFSO
	Dim objFile
	'	strLogpath="C:\Users\e000218\Desktop\Autodesk_Automation\Results\Autodesk_Automation_Results_5-6-2020-18-14\Logs\"&Environment.Value("TestName")&".txt"
		strLogpath=gstrStatusPath&"\"&Environment.Value("TestName")&".txt"
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.OpenTextFile(strLogpath,1,False)
			strText = objFile.readAll()
				If Instr(strText,"FAIL") > 0 Then
					strStatus = "FAIL"
				Else
					strStatus = "PASS"
				End If	
				fGetStatus=strStatus
		Set objFSO=Nothing
		Set objFile=Nothing
End Function

''******************************************************************************************************************************************************************************************************************************************
''	Function Name							:						 		 fnUploadReport
''	Objective								:								 Used to post the local html report into ALM
''	Input Parameters						:								 sLinkName													 
''	Date Created							:								 11-Oct-2018
''	QTP Version								:								 12.0
''	QC Version								:								 QC 11 
''	Pre-requisites							:								 NIL  
''	Created By								:		   						 Balaji Veeravalli
''	Modification Date						:		   
''******************************************************************************************************************************************************************************************************************************************
Public Function fUploadReport(pathToZipFile, dirToZip)
Dim fso  
Set fso = CreateObject("Scripting.FileSystemObject")  
    pathToZipFile = fso.GetAbsolutePathName(pathToZipFile)
    dirToZip = fso.GetAbsolutePathName(dirToZip)    
                If fso.FileExists(pathToZipFile) Then       
                   fso.DeleteFile pathToZipFile
                End If 
                If Not fso.FolderExists(dirToZip) Then
                   Exit Function
                End If   
			Set oSubFldrs= fso.GetFolder(dirToZip).SubFolders
         For each subfldr in oSubFldrs
	           If fso.GetFolder(subfldr.Path).Files.Count=0 Then
	              fso.DeleteFolder(subfldr.Path)
	            End If
		   Next  
Set fso = CreateObject("Scripting.FileSystemObject")    
Set file = fso.CreateTextFile(pathToZipFile&".zip")    
	file.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)    
	file.Close
	Set fso = Nothing
	Set file = Nothing    
    Set sa = CreateObject("Shell.Application")    
    Dim zip
    Set zip = sa.NameSpace(pathToZipFile&".zip")
    zipResultFolder = pathToZipFile & ".zip"    
    Dim d
    Set d = sa.NameSpace(dirToZip)    
    zip.CopyHere d.items, 4    
	    Do Until d.Items.Count <= zip.Items.Count
	       Wait 1
	    Loop     
'If Environment("TestRailStatusUpdate")="YES" Then
'	Call fGetDataforTestRail(trRID,trTID,gstrProjectResultPath&"\"& gstrFolderName)	    	
'End If

End Function



'******************************************************************************************************************************************************************************************************************************************
'	Sub Name		 					:					fFrameworkFolderConfiguration
'	Objective							:					Used to download Logos & OR
'	Input Parameters					:					
'	Output Parameters					:					NIL
'	Date Created						:					
'	QTP Version							:					UFT 15.0
'	QC Version							:					
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fFrameworkFolderConfiguration()
	
	'Verify Test Plan Name is received from Batch file, else continue with test plan configured in Global Constants File
	If fEnvironmentVarExists("TestExecutionPlanName") Then
		gstrTestPlanName = Environment.Value("TestExecutionPlanName")
		gstrProjectTestPlanPath = gstrFrameWorkFolder&"\"&gstrTestPlanName
		gstrProjectConfigFilePath = gstrProjectTestPlanPath&"\TestExecutionConfig.xml"
		gstrProjectTestdataPath = gstrProjectTestPlanPath&"\TestData\"
		gstrProjectResultPath = gstrProjectTestPlanPath&"\TestResults"
		gstrProjectFilesPath = gstrProjectTestPlanPath&"\Files"
		gstrProjectPDFFilesPath = gstrProjectFilesPath&"\PDFFiles"
	End If
	
	'Load environement file in runtime
	Call fLoadEnvironment()
		
	strResources = gstrFrameWorkFolder&"\Resources"
	Set objFso = CreateObject("Scripting.FileSystemObject")  

	'Verify Root Folder is existing
	If not objFso.FolderExists(gstrRootFolder) Then	
		Set objFolder=objFso.CreateFolder(gstrRootFolder)
	End If
	
	'Verify Framework folder is existing 
	If not objFso.FolderExists(gstrFrameWorkFolder) Then
		Set objFolder=objFso.CreateFolder(gstrFrameWorkFolder)
	End If
	
	'Verify Project folder is existing 
	If not objFso.FolderExists(gstrProjectResultPath) Then
		Set objFolder=objFso.CreateFolder(gstrProjectResultPath)
	End If
	
	'Verify Resource folder is existing in Project Folder
	If not objFso.FolderExists(strResources) Then
		Set objFolder=objFso.CreateFolder(strResources)
	End If

	'Verify TestData folder is existing in Test Plan Folder
	If not objFso.FolderExists(gstrProjectTestdataPath) Then
		Set objFolder=objFso.CreateFolder(gstrProjectTestdataPath)
	End If	
	
	'Verify if Results folder is received from Batch file, else folder name is framed here
	gstrFolderName = Empty
	If gstrFolderName=Empty Then
		fRptFoldername
	End If
	
	gstrResultsRootFolder=gstrProjectResultPath&"\"& gstrRootFolderName 
	gstrResultsFolder=gstrResultsRootFolder&"\"& gstrFolderName 
	
	'Verify Results root folder is existing, else new folder is created
	If not objFso.FolderExists(gstrResultsRootFolder) Then
	   Set objFolder=objFso.CreateFolder(gstrResultsRootFolder)
	End If
	
	'Verify Results folder is existsing, else new folder is created
	If not objFso.FolderExists(gstrResultsFolder) Then
	   Set objFolder=objFso.CreateFolder(gstrResultsFolder)
	End If
		
	'Verify if TestData folder Exists in Results folder, else create a new folder
	gstrOutputTestDataFolderPath = gstrResultsFolder & "\TestData"
	If not objFso.FileExists(gstrOutputTestDataFolderPath)Then
		Set objFolder=objFso.CreateFolder(gstrOutputTestDataFolderPath)
	End If
		
	strLogos=gstrResultsFolder&"\Logos"
	gstrTestData=gstrResultsFolder&"\TestData\"

	If not objFso.FolderExists(gstrTestData) Then
	   Set objFolder=objFso.CreateFolder(gstrTestData)
	End If

	'Verify if test data sheet is alread existsing in Test Data folder in Results, else copy new file from project test data folder
	If not objFso.FileExists(gstrTestData & Environment("TestName") & "_TD.xls") Then
		objFso.CopyFile gstrProjectTestdataPath& Environment("TestName") & "_TD.xls", gstrTestData
	End If  
	
	'Verify if invidual test Case reports folder is existsing, else create a new folder
	gstrTestcasesPath=gstrResultsFolder&"\Testcases"
	If not objFso.FolderExists(gstrTestcasesPath) Then
	   Set objFolder=objFso.CreateFolder(gstrTestcasesPath)
	End If
		
	'Verify if project logo files exists on Logos folder in test results folder
	If not objFso.FolderExists(strLogos) Then
	   Set objFolder=objFso.CreateFolder(strLogos)
	   objFso.CopyFile strResources&"\Pass.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Fail.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Companylogo.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Clientlogo.png",strLogos&"\"
	   objFso.CopyFile strResources&"\PassWithScr.png",strLogos&"\"
	   objFso.CopyFile strResources&"\Warning.png",strLogos&"\" 	
	End If
	
	Set objFso = Nothing
	
End Function
