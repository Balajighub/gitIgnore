'*************************************************************************************************************************************************************************************
'	Script Name					 	 :						HRMS_TS001Sample
'	Objective						 :						Used to Fecth Employee details in HRMS 
'	Date Created				     :						
'	Module						     :						HRMS
'	UFT Version					     :						UFT 15.00
'	QC(ALM) Version				     :						
'	Pre-requisites				     :						
'	Created By					     :						Cigniti Technologies
'*************************************************************************************************************************************************************************************

'Clear temp files and kill unwanted procesess
Call fClearTempFilesAndKillUnwantedProcess()
If Environment("ERRORFLAG") = False Then
	ExitAction()
End If

'Create Folder structure'
Call fSetupFolderStructure()

'Get the Rowc ount from Excel
Set oReturnableOrder = fGetTestDataBySheet("HRMSample")'
Do while Environment("ERRORFLAG")<>False 
	iTCID =  oReturnableOrder.Item("TCID" & iRowCountRef)   
			For intRowStartCount = 1 To UBound(oReturnableOrder.Keys)
				If Environment("ERRORFLAG") = False Then
					Exit Do 
				End If  
				
				Set objDataDict=oReturnableOrder.Item(intRowStartCount)
				
				'Read Test data from Excel -  These input variables can be passed directly from here to the test case			
				'strEmpName=objDataDict.Item("FinmartReport" & iRowCountRef)
				'strJobTitle=objDataDict.Item("Last Name" & iRowCountRef)				
				'strEmpStatus=objDataDict.Item("AccountName" & iRowCountRef)
				
				'Execute Test Case
				'Call fHRMS_TS001Sample(strEmpName, strJobTitle, strEmpStatus)
				
				'Execute Test Case
				Call fHRMS_TS001Sample(objDataDict,intRowStartCount)
			Next
		If Environment("ERRORFLAG")=False Then
			Exit Do 
		End If        
	Exit Do          
Loop

'Writing to Summary Report
Call fRptWriteResultsSummary()  







'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 				:					fHRMS_TS001Sample
'	Objective							:					Used to Fecth Employee details in HRMS 
'	Input Parameters					:					
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti Technologies
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fHRMS_TS001Sample(objDataDict,intRowStartCount)

	'Read Test data from Excel				
	strEmpName=objDataDict.Item("EmployeeName" & iRowCountRef)
	strJobTitle=objDataDict.Item("JobTitle" & iRowCountRef)				
	strEmpStatus=objDataDict.Item("EmployeeStatus" & iRowCountRef)

	'Login to HRMS Application
	Call fHRMSLogin()
	
	'Fetching employee details on HRMS Application
	Call fHRMSEmployeeSearch(strEmpName, strJobTitle, strEmpStatus)
	
End  Function


