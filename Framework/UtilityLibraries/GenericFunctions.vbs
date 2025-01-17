'******************************************************************************************************************************************************************************
'	Function Name						:		fCloseAllOpenBrowsers
'	Objective							:		Used to close all the open browsers
'	Input Parameters					:		strBrName -  Browser Name 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fCloseAllOpenBrowsers(strBrName)
	
	On error resume next
	
	'Close the specific/all browsers
	Select Case ucase(strBrName)
		Case "IE"
		  'Close Internet Explorer
		   SystemUtil.CloseProcessByName "iexplore.exe"
		Case "FF"
		  'Close Firefox
		   SystemUtil.CloseProcessByName "firefox.exe"
		Case "CHROME"
		  'Close Chrome 
		   SystemUtil.CloseProcessByName "chrome.exe"
		Case "ALL"
		  'Close All Browsers
		   SystemUtil.CloseProcessByName "iexplore.exe"
		   SystemUtil.CloseProcessByName "firefox.exe"
		   SystemUtil.CloseProcessByName "chrome.exe"
	End Select
	
End Function

'****************************************************************************************************************************************************************************************
'	Function Name						:		fClearTempFilesAndKillUnwantedProcess
'	Objective							:		Used to clear temp files and kill unwanted procesess
'	Input Parameters					:		NIL
'	Output Parameters					:		NIL
'	Date Created						:		
'	UFT Version							:		
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'*****************************************************************************************************************************************************************************************
Public Function fClearTempFilesAndKillUnwantedProcess()

	On error resume next

	'Kill Excel Process
	SystemUtil.CloseProcessByName("EXCEL.EXE")
	
	'Kill Chrome Process
	SystemUtil.CloseProcessByName("chrome.exe")
	
	'Kill Internet Explorer 
	SystemUtil.CloseProcessByName("iexplore.exe")
	
	'Kill Teams Process
	'SystemUtil.CloseProcessByName("Teams.exe")
	
	'Kill Firefox Process
	SystemUtil.CloseProcessByName("firefox.exe")
	
	'Clearing Temp folder
	Dim objtemp
	Set objfso = CreateObject ("Scripting.FileSystemObject")
	Set objWinShell = CreateObject ("Wscript.Shell")
	Set objtemp = objfso.GetFolder(objWinShell.ExpandEnvironmentStrings("%TEMP%"))
	
	'Deleting the files
	For each strFile in objtemp.Files
		objfso.DeleteFile strFile
	Next
	
	'Deleting the Subfolders 
	For Each StrSubfldr in objtemp.subfolders
		objfso.DeleteFolder(StrSubfldr),true
	Next
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fEnterText
'	Objective						:				Used to Enter a Text/Value into Edit Box 
'	Input Parameters				:				Object Name - Object Reference, strValue -  Value to be entered into text field, strText -  Name of the text field
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 				:				
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fEnterText(ByRef sObjectName,ByRef strValue, ByRef strFieldName)
	
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objRefObject 
	Environment.Value("StepName") = "Enter a value"
	fEnterText=False
	Set objRefObject = sObjectName	
	objRefObject.RefreshObject
	
	'Verify the object exists
	If objRefObject.Exist(MIN_WAIT) Then
		
		'Get the property value and validate
        If objRefObject.GetRoProperty("enabled") = True OR objRefObject.GetRoProperty("disabled") = 0 Then
'            objRefObject.highlight
        	
        	Select Case objRefObject.GetRoProperty("micclass")					
				Case "WebEdit","SAPEdit"
					'Set value
					objRefObject.Set strValue
					fEnterText = True
					'Validation
					If Instr(1,objRefObject.GetRoProperty("name"),"password",1) > 0 or instr(1,objRefObject.GetRoProperty("type"),"password",1) > 0 Then							
				   		Call fRptWriteReport("Pass","Enter value "&Chr(34)&"********"&chr(34)&" in '"&strFieldName&"' field" ,"Successfully entered value as "&Chr(34)&"********"&chr(34)&" in '"&strFieldName)
					Else
						Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' field" ,"Successfully entered value as "&Chr(34)&strValue&chr(34)&" in '"&strFieldName)						
					End if
					
				Case "WebCheckBox"
					objRefObject.Set strValue
					fEnterText = True
					Call fRptWriteReport("Pass","Select value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' checkbox" ,"Successfully the value selected from '"&strFieldName&"' checkbox")

				Case "JavaEdit"
					objRefObject.Set strValue
					fEnterText = True
					Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' field" ,"Successfully the field value entered as '"&strFieldName&"'")

				Case "OracleTextField"
				    objRefObject.Enter strValue
					fEnterText = True'					
					If instr(1,objRefObject.GetROProperty("name"),"password",1) > 0 or instr(1,objRefObject.GetROProperty("type"),"password",1) > 0 Then
						Call fRptWriteReport("Pass","Enter value "&Chr(34)&"********"&chr(34)&" in '"&strFieldName&"' field" ,"Successfully the field value entered as '"&strFieldName&"'")
					Else
						Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' field" ,"Successfully the field value entered as '"&strFieldName&"'")
					End If										
				
				Case "SAPGuiEdit"
					 objRefObject.set strValue
					fEnterText = True					
					If instr(1,RefObject.GetROProperty("name"),"password",1) > 0 or instr(1,RefObject.GetROProperty("type"),"password",1) > 0 Then
						Call fRptWriteReport("Pass","Enter value "&Chr(34)&"********"&chr(34)&" in '"&strFieldName&"' field" ,"Successfully the field value entered as '"&strFieldName&"'")
					Else
						Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' field" ,"Successfully the field value entered as '"&strFieldName&"'")
					End If
				Case "SAPWebExtCalendar"
					 objRefObject.SetDate strValue
					fEnterText = True					
					If instr(1,objRefObject.GetROProperty("name"),"password",1) > 0 or instr(1,objRefObject.GetROProperty("type"),"password",1) > 0 Then
						Call fRptWriteReport("Pass","Enter value "&Chr(34)&"********"&chr(34)&" in '"&strFieldName&"' field" ,"Successfully the field value entered as '"&strFieldName&"'")
					Else
						Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' field" ,"Successfully the field value entered as '"&strFieldName&"'")
					End If
				Case "SAPGuiCheckBox"
					objRefObject.set strValue
					objRefObject.SetFocus
					fEnterText = True					
					Call fRptWriteReport("Pass","Select value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' checkbox" ,"Successfully the value selected in '"&strFieldName&"' checkbox")
				Case "SAPWebExtCheckBox"
					objRefObject.set strValue
					fEnterText = True					
					Call fRptWriteReport("Pass","Select value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' checkbox" ,"Successfully the value selected in '"&strFieldName&"' checkbox")					
				
				Case "SAPGuiOKCode"
					objRefObject.set strValue
					objRefObject.SendKey ENTER
					fEnterText = True					
					Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' field" ,"Successfully the field value entered as '"&strFieldName&"'")
				Case "SAPGuiTextArea"
					objRefObject.set strValue
					fEnterText = True					
					Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' field" ,"Successfully the field value entered as '"&strFieldName&"'")
				Case "SAPGuiRadioButton"
					objRefObject.set
					objRefObject.SetFocus
					fEnterText = True
					Call fRptWriteReport("Pass","Select value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' radio button" ,"Successfully the value selected in '"&strFieldName&"' radio button")					
				
				Case "SAPWebExtRadioButton"
					objRefObject.set
					fEnterText = True
					Call fRptWriteReport("Pass","Select value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' radio button" ,"Successfully the value selected in '"&strFieldName&"' radio button")			
				Case Else
					objRefObject.Set strValue
					fEnterText = True
					Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strFieldName&"' field" ,"Successfully the field value entered as '"&strFieldName&"'")
             
             End Select
		Else 
			Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"The value not entered in "&chr(34)&strFieldName&Chr(34))
			Environment("ERRORFLAG")=False
			fEnterText = False			
		    Exit Function
		
	    End If
	Else
		Call fRptWriteReport("Fail","Enter value in '"&strFieldName&"'" ,"Value is not entered in '"&strFieldName&"' field")
		Environment("ERRORFLAG")=False
		fEnterText = False		
		Exit Function

	End If
	
End Function  


''******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fClick
'	Objective						:				Used to click an object
'	Input Parameters				:				Object Name, strButtonName
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 				:				
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
''******************************************************************************************************************************************************************************************************************************************
Public Function fClick(ByRef sObjectName, ByRef strButtonName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objRefObject   
	Environment.Value("StepName") = "Click on "&strButtonName	   
	fClick=False
	Set objRefObject = sObjectName
	objRefObject.RefreshObject

	If objRefObject.Exist(MIN_WAIT) Then
		If objRefObject.GetRoProperty("enabled") = True OR objRefObject.GetRoProperty("disabled") = 0 Then
			objRefObject.Highlight
			objRefObject.Click
			fClick = True
				If objRefObject.getroproperty("micclass") = "link" Then
					Call fRptWriteReport("Pass","Click on '"&strButtonName&"' link", "Clicked on '"&strButtonName&"' link successfully")
				Else
					Call fRptWriteReport("Pass","Click on '"&strButtonName&"' button", "Clicked on '"&strButtonName&"' button successfully")
				End If
			
		Else
			Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" not clicked")
	       	Environment("ERRORFLAG") = False
	       	fClick = False
	    End If
	  
	  Else
   	   Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" not clicked")
       Environment("ERRORFLAG") = False
       fClick = False
   
   End If
	
End Function



''******************************************************************************************************************************************************************************************************************************************
'	Function Name				    :				    fSelect
'	Objective						:					This function is used to select an item from either a List, Navigation bar, Radio button or Tab
'	Input Parameters				:					sObjectName,strItem ,strItemName
'	Output Parameters			    :					Nil
'	Date Created					:					
'	UFT Version 					:					UFT 15.0				
'	Pre-requisites					:					NIL  
'	Created By						:					Cigniti
'	Modification Date		        :		   			
''******************************************************************************************************************************************************************************************************************************************
Function fSelect(sObjectName,strItem,strText)
'On Error Resume Next
'Verify if Step Failed, If yes, it will not run the function
If Environment("StepFailed") = "YES" Then
	Exit Function
End If	
Dim RefObject 
   Environment.Value("StepName") = "Verify Value Selected In DropDown"	
	fSelect=False
		Set RefObject = sObjectName				 
        RefObject.RefreshObject
		If RefObject.Exist(MIN_WAIT) Then
            If RefObject.GetRoProperty("visible") = True OR RefObject.GetRoProperty("disabled") = 1 OR RefObject.GetRoProperty("enabled") = True Then
				Select Case RefObject.GetRoProperty("micclass")  
					Case "SAPUIMenu"
						RefObject.Select strItem
						Wait(MIN_WAIT)
						fSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' menu" ,"Successfully the value selected from '"&strText&"' menu")
					Case "WebList"
						Environment.Value("StepName") = "Verify Value Selected In DropDown"	
						RefObject.Select strItem
						Wait(MIN_WAIT)
						fSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' list" ,"Successfully the value selected from '"&strText&"' list")
					Case "WebRadioGroup"
						Environment.Value("StepName") = "Select Radio Button"
					  	  Wait(MIN_WAIT)
						RefObject.Select strItem
						fSelect = True						
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' radio group" ,"Successfully the value selected from '"&strText&"' radio group")
					Case "WebTabStrip"	
						Environment.Value("StepName") = "Select Web Tab Strip"
					    Wait(MIN_WAIT)
						RefObject.Select strItem
						fSelect = True	
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' tab" ,"Successfully the value selected from '"&strText&"' tab")
					Case "SAPCheckBox"
			                        RefObject.Set strItem
			                        Environment.Value("StepName") = "Select Checkbox"    
			                        fSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' checkbox" ,"Successfully the value selected from '"&strText&"' checkbox")
					Case "JavaList"
						Wait(MIN_WAIT)
						RefObject.Select strItem
						Wait(MIN_WAIT)
						fSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' list" ,"Successfully the value selected from '"&strText&"' list")
					Case "OracleList"
						RefObject.Select strItem
						Wait(MIN_WAIT)
						fSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' list" ,"Successfully the value selected from '"&strText&"' list")
					Case "OracleListOfValues"
						RefObject.Select strItem
						Wait(MIN_WAIT)
						fSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' list" ,"Successfully the value selected from '"&strText&"' list")
					Case "OracleTabbedRegion"
						RefObject.Select
						Wait(MIN_WAIT)
						fSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' tab" ,"Successfully the value selected from '"&strText&"' tab")
					Case "OracleCheckbox"
						RefObject.Select
						Wait(MIN_WAIT)
						fSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' checkbox" ,"Successfully the value selected from '"&strText&"' checkbox")
					Case "OracleRadioGroup"
						RefObject.Select strItem
						Wait(MIN_WAIT)
						fSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' radio group" ,"Successfully the value selected from '"&strText&"' radio group")
					Case "SAPGuiComboBox"
						RefObject.Select strItem
						Wait(MIN_WAIT)
						fnSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' combo box" ,"Successfully the value selected from '"&strText&"' combo box")
					Case "SAPGuiTabStrip"
						RefObject.Select strItem
						Wait(MIN_WAIT)
						fnSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' tab" ,"Successfully the value selected from '"&strText&"' tab")
					Case "SAPGuiMenubar"
						RefObject.Select strItem
						Wait(MIN_WAIT)
						fnSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' menu" ,"Successfully the value selected from '"&strText&"' menu")
					Case "SAPWebExtMenu"
						RefObject.Select strItem
						Wait(MIN_WAIT)
						fnSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' menu" ,"Successfully the value selected from '"&strText&"' menu")
					Case "SAPWebExtList"
						Environment.Value("StepName") = "Verify Value Selected In DropDown"	
						RefObject.Select strItem
						Wait(MIN_WAIT)
						fnSelect = True
						Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' list" ,"Successfully the value selected from '"&strText&"' list")
					Case "SAPRadioGroup"
					RefObject.Select strItem
					Wait(MIN_WAIT)
					fnSelect = True
					Call fRptWriteReport("Pass","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"' radio group" ,"Successfully the value selected from '"&strText&"' radio group")
                  End Select 
                ''Return boolean Value True to the Called block
				fSelect = True
			Else
				Call fRptWriteReport("Fail","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"'" ,"The value not selected from '"&strText&"'")
				Environment("ERRORFLAG")=False
				fSelect = False
			  	Exit Function
			End If
		Else
			Call fRptWriteReport("Fail","Select the value as '"&Chr(34)&strItem&chr(34)&"' from '"&strText&"'","The value not selected from '"&strText&"'")
			Environment("ERRORFLAG")=False
			fSelect = False
			Exit Function
		End If	
	''Return boolean Value True to the Called block
'	On Error GOTO 0
End Function  
 


''******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				    fWebtableGetCelldata
'	Objective						:					Used to get data from a cell in webtable
'	Input Parameters				:					sObjectName,iRow,iCol,strText
'	Output Parameters			    :					Nil
'	Date Created					:					
'	UFT Version 					:					UFT 15.0					
'	Pre-requisites					:					
'	Created By						:					Cigniti
'	Modification Date		        :		   			
''******************************************************************************************************************************************************************************************************************************************
Public Function fWebtableGetCelldata(sObjectName,iRow,iCol,strTableName)
'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objRefObject 
 	Environment.Value("StepName") = "Retrieve Value In a Table "
	Set objRefObject = sObjectName			 	
    
    'Initially Assigning block to False
	fWebtableGetCelldata=False
    objRefObject.RefreshObject
    
    'Get the data from specified cell in web table
	If objRefObject.Exist(MIN_WAIT) Then
        If objRefObject.GetRoProperty("enabled") = True OR objRefObject.GetRoProperty("visible") = True Then
			Select Case objRefObject.GetRoProperty("micclass")			
				
				Case "WebTable"
					iGetCelldata = objRefObject.GetCellData(Cint(iRow),Cint(iCol))
					Call fRptWriteReport("PASSWITHSCREENSHOT",Environment.Value("StepName"),chr(34) & "'"&strTableName&"'" & chr(34) & " retrieved value is : " & chr(34) & strTableName& chr(34))															
					fWebtableGetCelldata = iGetCelldata
				
				Case "JavaTable"
					fGetCelldata = objRefObject.GetCellData(Cint(iRow),Cint(iCol))
					Call fRptWriteReport("PASSWITHSCREENSHOT", sObjectName.ToString,sObjectName.ToString&" retrieved value is "&"'"&fGetCelldata&"'")
					fWebtableGetCelldata = True
				
				Case "OracleTable"
					fGetCelldata = objRefObject.GetFieldValue(Cint(iRow),Cint(iCol))
					Call fRptWriteReport("PASSWITHSCREENSHOT", sObjectName.ToString,sObjectName.ToString&" retrieved value is "&"'"&fGetCelldata&"'")
					fWebtableGetCelldata = True
			End select
		Else
            Call  fRptWriteReport("Fail", Environment.Value("StepName"),objRefObject.ToString&" retrieved value is "&fGetCelldata&"The value could not be Fetched as the object exit but not enabled")
            Environment("ERRORFLAG") = False
            fWebtableGetCelldata = False
		    Exit Function
		End If
	Else
	    Call  fRptWriteReport("Fail", Environment.Value("StepName"),objRefObject.ToString&" retrieved value is "&fGetCelldata&"The value could not be Fetched as the object does not exit")
	    Environment("ERRORFLAG") = False
	    fWebtableGetCelldata = False
	    Exit Function
	End If
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fHighlight
'	Objective						:				Used to highlight the mentioned object 
'	Input Parameters				:				Object Name
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fHighlight(sObjectName)
	
	'Initially Assigning to False
	On Error Resume Next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	fHighlight=False
	If Not IsObject(sObjectName) Then
		Set objRefObject=Eval(fGetObjectHierarchy(sObjectName)) 		 
	Else
		Set objRefObject = sObjectName
	End If
	objRefObject.RefreshObject
	
	'Highlight on specifed object
	If objRefObject.Exist(MID_WAIT) Then
		If objRefObject.GetRoProperty("enabled") = True OR objRefObject.GetRoProperty("disabled") = 0 Then	
			objRefObject.highlight
			fHighlight = True
			Exit Function
	    End If
	Else
	   Call fRptWriteReport("Fail", sObjectName, "Highlight operation is not performed on" &sObjectName &" object is disabled")
	   Environment("ERRORFLAG") = False
	   fHighlight = False
	   Exit Function
	End If
	On Error Goto 0
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fWebCheckBox
'	Objective						:				Used to Check and uncheck checkboxes in any environment 
'	Input Parameters				:				objCheckBoxName  (Name of the checkbox object during object spy)
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fWebCheckBox(objCheckBoxName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set objDesc = Description.Create()	
		objDesc("micclass").Value = "WebCheckBox"
		objDesc("html tag").Value= "INPUT"
		objDesc("type").value = "checkbox"
	Set objChkBoxCount = oFSObj.ChildObjects(objDesc)	
	
	For intloop = 0 to objChkBoxCount.Count -1
		appVal = Trim(objChkBoxCount(intloop).GetROProperty("name"))
		if instr(appVal,objCheckBoxName) then
			objChkBoxCount(intloop).set "ON"
		End If
	Next
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fWebCheckBoxoff
'	Objective						:				Used to Check and uncheck checkboxes in any environment 
'	Input Parameters				:				objCheckBoxName  (Name of the checkbox object during object spy)
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fWebCheckBoxoff(objCheckBoxName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set objDesc = Description.Create()	
		objDesc("micclass").Value = "WebCheckBox"
		objDesc("html tag").Value= "INPUT"
		objDesc("type").value = "checkbox"
	Set objChkBoxCount = oFSObj.ChildObjects(objDesc)	
	
	'Condition to uncheck checkboxes in any environment 
	For intCount = 0 to objChkBoxCount.Count -1
		appVal = Trim(objChkBoxCount(intCount).GetROProperty("name"))
		if instr(appVal,objCheckBoxName) then
			objChkBoxCount(intCount).set "OFF"
		End If
	Next
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fVerifyProperty
'	Objective						:				Used to verify the property value of the given object 
'	Input Parameters				:				sObjectName,sPropertyName, strValue
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Function fVerifyProperty(strObjectName,strPropertyName,strValue)
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    Dim objRefObject
  	Set objRefObject = strObjectName		
  	
  	'Get the property name from object
	If objRefObject.Exist(MID_WAIT) Then
		strText=objRefObject.GetRoProperty(strPropertyName)
	End If
		
	If Trim(strValue)=Trim(strText) Then
'         Call fRptWriteReport("PASSWITHSCREENSHOT",strObjectName,"Value is Retrieved and the Value is " &strValue) 
         Call fRptWriteReport("PASSWITHSCREENSHOT", "Caompare values","Values are matched and the compared values are '"&strValue&"'")
	     fVerifyProperty = True         
   	Else   	  	 
         Call fRptWriteReport("Fail",strObjectName,"Compare values","Values are not matched and the compared values are '"&strValue&"' and '"&strText&"'" )
         Environment("ERRORFLAG") = False
         fVerifyProperty = False
    End If
    
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fSynUntilFieldExists
'	Objective						:				Used to wait until field exists on System
'	Input Parameters				:				gfString, intWait
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fSynUntilFieldExists(gfString,intWait)

	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set objFSO = createobject("Scripting.filesystemobject")
	set objFile = objFSO.OpenTextFile(StrFrameWorkFolder&"\Resources\Log.txt")
	strData = objFile.ReadAll	
	gWait=0
	Do 
		wait 1
		gWait=gWait+1
	Loop Until(Instr(strData,gfString)>0 or gWait=intWait)
	wait 2
	On error goto 0
	set objFile =Nothing
	Set objFSO= Nothing
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fSynUntilObjExists
'	Objective						:				Used to wait until object exists on application
'	Input Parameters				:				objControl,intWait
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fSynUntilObjExists(objControl,intWait)
    Dim blnFound,RefObject,gWait
    On error resume next
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    blnFound=False
    gWait=0   
    If Not IsObject(objControl) Then        
    Else
        Set RefObject = objControl   
    End If   
    Do
        wait 1
        gWait=gWait+1
        If RefObject.Exist(4) Then blnFound=True
    Loop Until(RefObject.Exist(4) or gWait=intWait)   
        If blnFound=False Then
            rptWriteReport "WARNING", "Dynamic Wait-fnSynUntilObjExists" , RefObject.toString&" --object not found"
        End If   
    fSynUntilObjExists=blnFound
    Set RefObject=Nothing
    On error goto 0   
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetFutureDateAdd()
'	Objective						:				Add days to current date
'	Input Parameters				:				NoOfDays
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fGetFutureDateAdd(NoOfDays)
    
    currentDate = Date+NoOfDays
    dtDate = Day(currentDate)
    If Len(dtDate) = 1 Then
        dtDate = 0&dtDate
    End If
    
    strMonthName=Month(currentDate)
    dtMonth= left(strMonthName,3) 
	If Len(dtMonth) = 1 Then
	    dtMonth = 0&dtMonth
	End If
	
    gfGetCurrentCalendarMonthName = Ucase(dtMonth)		
    dtYear = Year(currentDate)
    strFutureDate = gfGetCurrentCalendarMonthName &"/"&dtDate&"/"&dtYear
    fGetFutureDateAdd = strFutureDate
    
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetCurrentDate()
'	Objective						:				Used to get current date
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fGetCurrentDate()
	
	currentDate = Date
    dtDate = Day(currentDate)
	    If Len(dtDate) = 1 Then
	        dtDate = 0&dtDate
	    End If
    strMonthName=MonthName(Month(currentDate))	
    dtMonth= left(strMonthName,3) 
    gfGetCurrentCalendarMonthName = Ucase(dtMonth)		
    dtYear = Year(currentDate)
    strCurDate = dtDate &" "&gfGetCurrentCalendarMonthName&" "&dtYear
    fGetCurrentDate = strCurDate
    
End Function



'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fDateConversion()
'	Objective						:				Used to Convert Date to DD-Mon-YYYY Format
'	Input Parameters				:				sDate
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fDateConversion(sDate)

	strDateFormat  = FormatDateTime(sDate, 1)
	strDateFormat1  = Split (Trim(strDateFormat),",",-1,1)
	strDt  = Right(Trim(strDateFormat1(1)),2)
	strMonth  = Left(Trim(strDateFormat1(1)),3)
	strDateformat  = strDt&"-"&strMonth&"-"&Trim(strDateFormat1(2))
	fDateConversion = strDateFormat
			
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetNumericValues()
'	Objective						:				Get numeric value from string
'	Input Parameters				:				strVal
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fGetNumericValues(strVal)
	
	On Error resume next
	For iCount = 1 to len(strVal)
		strValues = mid(strVal,iCount,1)
		If isnumeric(strValues) Then
			intNumbers = intNumbers&strValues
		Else
			alphabet = alphabet&strValues
		End If
	Next
	fGetNumericValues = intNumbers
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetEnvFilePath()
'	Objective						:				Get the Path of the Enviromental File based on path that is passed (From QC or Local Path)
'	Input Parameters				:				sProjectPath
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fGetEnvFilePath(strProjectPath)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	If Instr(1,strProjectPath,"Subject") <> 0 Then
		Environment("QCENVPATH") = strProjectPath & "\" & "EnvironmentalVariables"
		strFile = fGetFolderAttachmentPath(Environment("QCENVPATH"),"Environment.xls")
	Else
		strFile = strProjectPath & "\" & "EnvironmentalVariables\Environment.xls"
	End If
	
	'Return the values to the function
	fGetEnvFilePath = strFile
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fSplitFor()
'	Objective						:				Used to split two values in the String using a delimetet
'	Input Parameters				:				intItemNumber,strSplitChar,strString
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fSplitFor(intItemNumber,strSplitChar,strString)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	If Len(Trim(strString))=0 Then
		fSplitFor=""
		Exit Function
	End If
	
	arrValue = Split(strString,strSplitChar)
	
	If IsNumeric(intItemNumber) Then intItemNumber=intItemNumber-1

	If intItemNumber>UBound(arrValue) Then
		fSplitFor=arrValue(UBound(arrValue))
	ElseIf intItemNumber<LBound(arrValue) Then
		fSplitFor=arrValue(LBound(arrValue))
	Else
		fSplitFor = arrValue( intItemNumber)
	End If

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fRandomNumber()
'	Objective						:				Used to generate Random Number
'	Input Parameters				:				intLowBound,intUpperBound,strText
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fRandomNumber(intLowBound,intUpperBound,strText)
'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim intRand
	Randomize
	intRand = Int((intUpperBound - Cint(intLowBound) + 1) * Rnd + Cint(intLowBound))
	If strText<>"" Then
		intRand=strText&intRand
	End If
	
	'Return value
	fRandomNumber=intRand
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetRowCount()
'	Objective						:				Used to get the rowcount with Run 'Y' from Excel or MS Access
'	Input Parameters				:				sSheetName
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fGetRowCount(sSheetName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	strFile = gstrProjectTestdataPath&Environment("ActionName")&"_Testdata.xls"
	strItemName = strSheetName

	Set DB_CONNECTION=CreateObject("ADODB.Connection")
	DB_CONNECTION.Open "DBQ="&strFile&";DefaultDir=C:\;Driver={Driver do Microsoft Excel(*.xls)};DriverId=790;FIL=excel 8.0;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\matdsn2.dsn;MaxScanRows=8;PageTimeout=5;ReadOnly=0;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
	
	intCheck = Instr(1,strItemName,"$")
	If intCheck = 0 Then
		sItemName = strItemName&"$"
	End If
	
	strQuery =  "SELECT Count(*) FROM ["&strItemName&"] WHERE Run = 'Y'"
	set Record_Set1=DB_CONNECTION.Execute(strQuery)
	intRowCountValue = 0

	Do While Not Record_Set1.EOF
		For Each Element In Record_Set1.Fields
				intRowCount = Record_Set1(intRowCountValue)
		Next
		Record_Set1.MoveNext
	Loop
	Record_Set1.Close
	Set Record_Set1=Nothing
	DB_CONNECTION.Close
	Set DB_CONNECTION=Nothing
    fGetRowCount = iRowCount

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetIterations
'	Objective						:				Used to get the rowcount with Run 'Y' from Excel or MS Access
'	Input Parameters				:				sSheetName
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fGetIterations(strSheetName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	strFile = gstrProjectTestdataPath&Environment("ActionName")&"_Testdata.xls"
	strItemName = strSheetName
	Set DB_CONNECTION=CreateObject("ADODB.Connection")
	DB_CONNECTION.Open "DBQ="&gstrFile&";DefaultDir=C:\;Driver={Driver do Microsoft Excel(*.xls)};DriverId=790;FIL=excel 8.0;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\matdsn2.dsn;MaxScanRows=8;PageTimeout=5;ReadOnly=0;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
	intCheck = Instr(1,sItemName,"$")
	If intCheck = 0 Then
		strItemName = strItemName&"$"
	End If
	strQuery =  "SELECT Count(*) FROM ["&strItemName&"] WHERE Run = 'Y' and TCID >= 1"
	set Record_Set =DB_CONNECTION.Execute(strQuery)
	iRowCountValue = 0

	Do While Not Record_Set.EOF
		For Each Element In Record_Set.Fields
			iRowCount = Record_Set(iRowCountValue)
		Next
		Record_Set.MoveNext
	Loop
	Record_Set1.Close
	Set Record_Set=Nothing
	DB_CONNECTION.Close
	Set DB_CONNECTION=Nothing
    fGetIterations = intRowCount
    
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetTestData
'	Objective						:				Used to get the testdata with Run ' 'Y' from Excel or MS Access into Dictonary Object
'	Input Parameters				:				strItemName
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fGetTestData(strItemName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If		
	strFile = gstrProjectTestdataPath&Environment("TestName")&"_Testdata.xls"

	Set Data = CreateObject("Scripting.Dictionary")
	Data.RemoveAll		
	intCheck = Instr(1,strItemName,"$")
	If intCheck = 0 Then
		sItemName = strItemName&"$"
	End If

	strQuery =  "SELECT * FROM ["&strItemName&"] Where Run = 'Y'"
	Set DB_CONNECTION=CreateObject("ADODB.Connection")
	
	DB_CONNECTION.Open "DBQ="&strFile&";DefaultDir=C:\;Driver={Driver do Microsoft Excel(*.xls)};DriverId=790;FIL=excel 8.0;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\matdsn2.dsn;MaxScanRows=8;PageTimeout=5;ReadOnly=0;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"	

	Set Record_Set1=DB_CONNECTION.Execute(strQuery)
	Set Record_Set2=DB_CONNECTION.Execute(strQuery)

	intRowCount = 0

	Do While Not Record_Set2.EOF
		intColumnCount = 0
		For Each Field In Record_Set1.Fields
			strColumnName = Field.Name& (intRowCount + 1)
			intRowValue = Record_Set2(intColumnCount)
			If IsNull(intRowValue) Then
				intRowValue = ""
			End If
			Data.Add strColumnName,intRowValue
		intColumnCount = intColumnCount + 1
		Next
		Record_Set2.MoveNext
		intRowCount = intRowCount + 1
	Loop

	Record_Set1.Close
	Set Record_Set1=Nothing
	Record_Set2.Close
	Set Record_Set2=Nothing
	DB_CONNECTION.Close
	Set DB_CONNECTION=Nothing
	Set fGetTestData = Data	

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetSingleValue
'	Objective						:				Used to get the Single Test  from Prameter Sheet based on the Field Name from Env Excel
'	Input Parameters				:				strFieldName,strTcode,strTestName
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fGetSingleValue(strFieldName,strTcode,strTestName) 


   	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
   	'gstrFile = gstrOutputTestDataFile      
    strItemName = strTcode    
	intCheck = Instr(1,strItemName,"$")
	If intCheck = 0 Then
		strItemName = strItemName&"$"			
	End If

    'strQuery="Select "&strFieldName&" from ["&strItemName&"] where TCID = "&Environment("ROWCOUNT") 	'added single quote to tcid date: 15-Mar-2017
    strQuery="Select "&strFieldName&" from ["&strItemName&"] where TCID = '"&iTCID&"'"
    
	strSingleValue = ""
	Dim DB_CONNECTION
	Set DB_CONNECTION=CreateObject("ADODB.Connection")
	DB_CONNECTION.Open "DBQ="&gstrFile&";DefaultDir=C:\;Driver={Driver do Microsoft Excel(*.xls)};DriverId=790;FIL=excel 8.0;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\matdsn2.dsn;MaxScanRows=8;PageTimeout=5;ReadOnly=0;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
	Set Record_Set=DB_CONNECTION.Execute(strQuery)
	Dim intRowSingleValue
	intRowSingleValue = 0
	Do While Not Record_Set.EOF
		    strSingleValue = Record_Set(0)
			If IsNull(strSingleValue) or Len(Trim(strSingleValue))=0 Then
				strSingleValue="Empty"
			End If
				intRowSingleValue = 1
			Exit Do
	Loop
	If intRowSingleValue = 0 Then strSingleValue="Empty"
	fGetSingleValue = strSingleValue
	Record_Set.Close
	Set Record_Set=Nothing
	DB_CONNECTION.Close
	Set DB_CONNECTION=Nothing

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetRowByTCID
'	Objective						:				Used to get row data by using TCID
'	Input Parameters				:				sSheetName,iTCID
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fGetRowByTCID(sSheetName,iTCID)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	gstrFile = gstrProjectTestdataPath&strActionName&"_Testdata.xls"
	sItemName = sSheetName
	Set oExl = CreateObject("Excel.Application")
	Set oWbk = oExl.Workbooks.Open(gstrFile)
	Set oWkSht = oWbk.Worksheets(sItemName)		
	iRowCou = oWkSht.UsedRange.Rows.Count		
	For iCount = 2 to iRowCou		
		If  CStr(trim(oWkSht.Cells(iCount,2).Value)) = CStr(iTCID) Then
            Exit for
		End If
	Next		
	fGetRowByTCID = ite - 1
	Set oWkSht = Nothing
	Set oWbk = Nothing
	oExl.Quit
	Set oExl = Nothing

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fCreateResultsSummaryFile
'	Objective						:				Used to Create the Summary File in HTML
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fCreateResultsSummaryFile()
	
	Environment("TOTALFAILEDCASES") = 0
	Environment("TOTALSUCCESSCASES") = 0

	'Dynamically  creating Environment Variable for Result Template File
	sResultsTemplateFileName =  "Result_Summary_"&fTimeStamping()
	Environment("SUMMARYCOUNT") = 1
	
	If Environment("QC") = "Yes" Then
		Environment("QCTEMPREPORTPATH") = "C:\QC Temp Reports"
		Set fso = CreateObject("Scripting.FileSystemObject")
		If Not (fso.FolderExists(Environment("QCTEMPREPORTPATH"))) Then
			fso.CreateFolder(Environment("QCTEMPREPORTPATH"))
		End If
		Set fso = Nothing
		Environment("SUMMARYREPORTHTML") = Environment("QCTEMPREPORTPATH") & "\" & sResultsTemplateFileName & ".html"
		sMasterFilePath = fGetFolderAttachmentPath(Environment("QCMASTERPATH"),"Master.xls")
		sMasterItemName = "SCENARIO_LIST"
	Else
		Environment("SUMMARYREPORTHTML") = gstrProjectResultPath & "\" & sResultsTemplateFileName & ".html"
	End If
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.CreateTextFile(Environment("SUMMARYREPORTHTML"))
	Call fReportSummaryHeader()
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fTimeStamping
'	Objective						:				Used to Create the Current Time Stamping
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fTimeStamping()

	fTimeStamping = Replace(Date(),"/","-") & "_" & Replace(Time(),":",".")

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fCreateFolderStructure
'	Objective						:				Used to Create the Detailed Results File in Excel, Access or HTML based on the Testdata Type and Secnario ID Folder
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fCreateFolderStructure()
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	sDetailedResultsTemplateFileName =  "Detailed_Results_"& Environment("ActionName") &"_"& fTimeStamping() 
	Environment.Value("DETAILEDRESULTSFILENAME") = sDetailedResultsTemplateFileName	
	If Environment("QC") = "Yes" Then
		Environment("DETAILEDREPORTHTML") = Environment("QCTEMPREPORTPATH") & "\" & sDetailedResultsTemplateFileName & ".html"
		Call fAddScenarioFolderInToQC()
	Else
		Set fso = CreateObject("Scripting.FileSystemObject")
	   	'Creating Folders for Detailed Results
		If Not (fso.FolderExists(gstrProjectResultPath)) Then
			fso.CreateFolder(gstrProjectResultPath)
		End If
				
		If Not (fso.FolderExists(gstrProjectResultPath & "\" & Environment("ActionName"))) Then
			fso.CreateFolder(gstrProjectResultPath &  "\" & Environment("ActionName"))
		End If
		
		Environment("DETAILEDREPORTHTML") = gstrProjectResultPath & "\" & Environment("ActionName") & "\" & sDetailedResultsTemplateFileName & ".html"
	End If
	  
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.CreateTextFile(Environment("DETAILEDREPORTHTML"))
	Environment("TRANSACTIONSTARTTIME") = Timer
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fExecuteScript
'	Objective						:				Used to Navigate to the Scenarion and any Common code will be written if Exists
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fExecuteScript()
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Call fCreateFolderStructure()
	Call fStartTransactionExecution()

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fErrorCheck
'	Objective						:				Used to Checks for the Existence of the Error Message for evey Button and Link Click
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fErrorCheck()
	
	strErrorType = fGetStatusBarMessage("messagetype")
	If sErrorType = "E" Then
		sErrorMsg = fGetStatusBarMessage("text")
		Call fRptWriteReport("Fail","Error check for every Button/Link click",strErrorType&" shouldn't be shown in the application")
		Environment("ERRORFLAG") = False
		ExitAction
	End If
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fCaptureScreenShot()
'	Objective						:				Used to Capture the Snap shot and place it in the Result Folder
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function fCaptureScreenShot()

	Dim sPNG,sSSPath,oDesktop
	If Environment("QC") = "Yes" Then
		sSSPath = Environment("QCTEMPREPORTPATH") 
		strFileName = Environment("SCENARIOID") & "_"& fTimeStamping()&".png"
		sPNG = sSSPath &"\"& strFileName
		sQCPath = Environment("QCSNAPSHOTPATH") &"\"& strFileName
		fCaptureScreenShot = sQCPath
	Else
		sSSPath = gstrProjectResultPath & "\" & Environment("TestName")
		sPNG = sSSPath&"\"& Environment("TestName") & "_"& fTimeStamping()&".png"
		fCaptureScreenShot = sPNG
	End If		
    Set oDesktop = Desktop
	oDesktop.CaptureBitmap sPNG, True
	If  Environment("QC") = "Yes" Then
		Call fExportErrorSnapShotToQC(sPNG)
	End If
	
	Set fso = Nothing
	Set oDesktop = Nothing

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fTabOutSync()
'	Objective						:				Used to tab from one field to another field
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************

Public Function  fTabOutSync()

	On Error Resume Next
	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.SendKeys "{TAB}"
    
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fMinimizeUFT()
'	Objective						:				Used to Minimize the UFT while execution to take Application Snap Shots
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fMinimizeUFT()
	
	Dim objUFT
	Set objUFT = GetObject("","QuickTest.Application")
	objUFT.WindowState = "Minimized"
	Set objUFT = Nothing
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fKillProcess()
'	Objective						:				Used to Kill the Process based on the Process Name that is Passed as parameter
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fKillProcess(strProcessName)
    
    Dim intTerminationCode
    Dim objService
    Dim objInstance
    Dim strProcess
    Dim strProPath
    Dim intStatus
    
    intTerminationCode = 0
    fKillProcess = True
    Set objService = GetObject("winmgmts:")
    For Each strProcess In objService.InstancesOf("Win32_process")
        If UCase(strProcess.Name) = UCase(strProcessName) Then
            strProPath = "Win32_Process.Handle=" & strProcess.ProcessID
            Set objInstance = objService.Get(strProPath)
            intStatus = objInstance.Terminate(intTerminationCode)
            If intStatus = 0 Then
                Finc_KillProcess = False
            End If
        End If
    Next
    
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetTestDataFilePath()
'	Objective						:				Used to Get the Full Path with TestData file Name based on the Testdata Type (Excel or Access)
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fGetTestDataFilePath()
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	If Environment("QC") = "Yes" Then
		sTestDataFileName = Environment("SCENARIOID") & "_TestData" & ".xls"	
		sTestDataFile = Environment("QCTESTDATAPATH")
		sTestDataFilePath = fGetFolderAttachmentPath(sTestDataFile,sTestDataFileName)
	Else
		sTestDataFilePath = Environment("TESTDATAPATH") & "\" & Environment("SCENARIOID") & "_TestData" & ".xls"
	End If	
	fGetTestDataFilePath = sTestDataFilePath

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetEnvFilePath()
'	Objective						:				Get the Path of the Enviromental File based on path that is passed (From QC or Local Path)
'	Input Parameters				:				sProjectPath
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fGetEnvFilePath(sProjectPath)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	If Instr(1,sProjectPath,"Subject") <> 0 Then
			Environment("QCENVPATH") = sProjectPath & "\" & "EnvironmentalVariables"
			gstrFile = fGetFolderAttachmentPath(Environment("QCENVPATH"),"Environment.xls")
	Else
			gstrFile = sProjectPath & "\" & "EnvironmentalVariables\Environment.xls"
	End If
	'Return the values to the function
	fGetEnvFilePath = gstrFile
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fCreateQuery()
'	Objective						:				Used to Create a Condition based on the Value used in the Transaction Range
'	Input Parameters				:				sTransactionRange
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fCreateQuery(sTransactionRange)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	If fDoesExist(sTransactionRange,";") Then  
		ExitRun
	End If

    If IsNull(sTransactionRange) or Len(Trim(sTransactionRange))=0 Then
		MsgBox "No transactions were allocated for this machine"	
		fCreateQuery=-1
		ExitRun
	End If
	sTransactionRange = UCase(sTransactionRange)
	sTransactionRange =Trim(sTransactionRange)
	If fDoesExist (sTransactionRange,";")Then
		arr1=Split(sTransactionRange,";")
		sIndividualTIDs=""
		sRangeTIDs=""
		RangesExist=False
		For i=0 to UBound(arr1)
			If fDoesExist(arr1(i),"-")Then
			RangesExist=True
			sRangeTIDs=sRangeTIDs&";"&Trim(arr1(i))
			Else
			sIndividualTIDs=sIndividualTIDs&";"&Trim(arr1(i))
			End If
			If i=UBound(arr1) Then 
			If Trim(Len(sRangeTIDs))>0 then sRangeTIDs=Right(sRangeTIDs,Len(sRangeTIDs)-1)
			If Trim(Len(sIndividualTIDs))>0 then sIndividualTIDs=Right(sIndividualTIDs,Len(sIndividualTIDs)-1)
			End If
		Next
		arrRanges=Split(sRangeTIDs,";")
		For j=0 to UBound(arrRanges)
			CurrentRange=Trim(arrRanges(j))
			nStart=fSplitFor(1,"-",CurrentRange)
			nEnd=fSplitFor(2,"-",CurrentRange)
			If IsNumeric(nStart) and IsNumeric(nEnd) Then
				nStart=CDbl(fSplitFor(1,"-",CurrentRange))
				nEnd=CDbl(fSplitFor(2,"-",CurrentRange))
				If nStart>=nEnd Then
					ExitRun
				Else
					If j=0 Then
						sSubQuery=" TCID between "&nStart&" and "&nEnd
					Else
						sSubQuery=" or TCID between "&nStart&" and "&nEnd
					End If
					sQuery=sQuery&sSubQuery
				End If
			Else
			   ExitRun  
			End If
		Next
		
		bIndividualTIDsExist=False
		arrIndividualTIDs=Split(sIndividualTIDs,";")
		If  UBound(arrIndividualTIDs)<>-1 Then bIndividualTIDsExist=True
		If RangesExist and bIndividualTIDsExist Then
			sQuery=sQuery&" or TCID in("
		ElseIf RangesExist=False and bIndividualTIDsExist=True Then
			sQuery=sQuery&" TCID in("
		End If
		For k=0 to UBound(arrIndividualTIDs)
			If IsNumeric(arrIndividualTIDs(k)) Then
				If k=0 and k<>UBound(arrIndividualTIDs)Then
					sQuery=sQuery&Trim(arrIndividualTIDs(k))&","
				ElseIf k=0 and k=UBound(arrIndividualTIDs) Then
					sQuery=sQuery&Trim(arrIndividualTIDs(k))&")"
				ElseIf k<>0 and k=UBound(arrIndividualTIDs) Then
					sQuery=sQuery&Trim(arrIndividualTIDs(k))&")"
				ElseIf k<>0 and k<>UBound(arrIndividualTIDs)Then
					sQuery=sQuery&Trim(arrIndividualTIDs(k))&","         
				End If
			Else
				ExitRun
			End If
		Next
		sQuery=sQuery&" order by TCID"
		fCreateQuery=sQuery
	ElseIf fDoesExist(sTransactionRange,"-") Then
		CurrentRange=Trim(sTransactionRange)
		nStart=fSplitFor(1,"-",CurrentRange)
		nEnd=fSplitFor(2,"-",CurrentRange) 
		If IsNumeric(nStart) and IsNumeric(nEnd) Then
			nStart=CDbl(fSplitFor(1,"-",CurrentRange))
			nEnd=CDbl(fSplitFor(2,"-",CurrentRange))
			If nStart>=nEnd Then
				ExitRun
			Else
				sSubQuery=" TCID between "&nStart&" and "&nEnd
				sQuery=sQuery&sSubQuery
			End If
			sQuery=sQuery&" order by TCID"
			fCreateQuery=sQuery
		Else
			ExitRun  
		End If    
	ElseIf UCase(Trim(sTransactionRange))="ALL" Then
		arrsQuery=Split(sQuery,"and")
		sQuery=Join(arrsQuery)
		sQuery=sQuery&" order by TCID"
		fCreateQuery=sQuery
	ElseIf IsNumeric(sTransactionRange) Then
		sQuery=sQuery&" TCID="&sTransactionRange
		sQuery=sQuery&" order by TCID"
		fCreateQuery=sQuery
	Else
		ExitRun
	End If
	
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fDoesExist()
'	Objective						:				Used to check whether the values are the in a string or not
'	Input Parameters				:				sString1,sString2
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fDoesExist(strString1,strString2)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	If IsNull(strString1) or IsNull(strString2) or Len(Trim(strString1))=0 or Len(Trim(strString2))=0 Then
		fDoesExist=False
		Exit Function
	End If
	a=Trim(UCase(strString1))
	b=Trim(UCase(strString2))
	arr=Split(a,b)
	If UBound(arr)>=1 Then
		fDoesExist=True
	Else
		fDoesExist=False
	End If
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetRowCountFromTestData()
'	Objective						:				Used to retrive the RowCount from the Test Data File
'	Input Parameters				:				gstrFile,sItemName
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fGetRowCountFromTestData(gstrFile,sItemName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	sQuery =  "SELECT * FROM["&sItemName&"] WHERE Run = 'Y'"
	Set DB_CONNECTION=CreateObject("ADODB.Connection")

	iCheck = Instr(1,sItemName,"$")
	If iCheck = 0 Then
		sItemName = sItemName&"$"
	End If
	
	If Environment("TRANSACTIONRANGE") = "" Then
		sQuery =  "SELECT * FROM ["&sItemName&"] WHERE Run = 'Y'"
	Else 
		sQueryCondition= fCreateQuery(Environment("TRANSACTIONRANGE"))
		sQuery =  "SELECT * FROM ["&sItemName&"] WHERE Run = 'Y' and "&sQueryCondition
	End If

	DB_CONNECTION.Open "DBQ="&gstrFile&";DefaultDir=C:\;Driver={Driver do Microsoft Excel(*.xls)};DriverId=790;FIL=excel 8.0;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\matdsn2.dsn;MaxScanRows=8;PageTimeout=5;ReadOnly=0;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
                
	Set Record_Set1=DB_CONNECTION.Execute(sQuery)
	Set Record_Set2=DB_CONNECTION.Execute(sQuery)
	iRowCount = 0

	Do While Not Record_Set2.EOF
		For Each Field In Record_Set1.Fields
			If IsNull(iRowValue) Then
				iRowValue = ""
			End If
		Next
		Record_Set2.MoveNext
		iRowCount = iRowCount + 1
	Loop

	Record_Set1.Close
	Set Record_Set1=Nothing
	Record_Set2.Close
	Set Record_Set2=Nothing
	DB_CONNECTION.Close
	Set DB_CONNECTION=Nothing
	fGetRowCountFromTestData = iRowCount
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetFormFilePath()
'	Objective						:				Used to fetch the forms file path
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fGetFormFilePath()
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	sTestDataFilePath = Environment("TESTDATAPATH") & "\" & Environment("TRANSACTIONTYPE") & "_" & Environment("PRODUCTNAME") & "_Forms" & ".xls"
	fGetFormFilePath = sTestDataFilePath
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fObjectCheck()
'	Objective						:				Used to check for the Object Existece, Enable/Disable
'	Input Parameters				:				ObjControl
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fObjectCheck(objControl)   
   	On Error Resume Next
   	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	sFlag = True
	If(IsEmpty(objControl)) Then
		sFlag = False
		Exit Function 
	ElseIf objControl Is Nothing Then
		sFlag = False
		Exit Function 
	ElseIf(Not(IsObject(objControl)))Then
		sFlag = False
		Exit Function
	ElseIf(objControl.GetROProperty("disabled")) Then
		sFlag = False
		Exit Function
	End If
	fObjectCheck = sFlag
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fIfToString()
'	Objective						:				used to display ObjectName
'	Input Parameters				:				ObjControl
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fIfToString(objControl)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim Properties,intPropertiesCount,arrProperties(),i
	If objControl Is Nothing Then
		Exit Function
	End If
	
	If(Trim(Split(objControl.ToString," ")(0))="[")Then
		Set Properties = objControl.GetTOProperties
		intPropertiesCount =  Properties.Count
		ReDim  arrProperties(intPropertiesCount - 1)
		For i = 0 To intPropertiesCount - 1
			arrProperties(i)=Properties(i).Name & ":" &  Properties(i).Value
		Next
		fIfToString = Join(arrProperties,",")
	Else
		'If Object is in OR
		fIfToString = Trim(objControl.GetTOProperty("TestObjName"))
	End If
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fButClk()
'	Objective						:				Used to Click the  "SwfButton","SwfLabel" and Check for the Error Message (Exception Handling)
'	Input Parameters				:				sObjControl,sErrorName
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fButClk(strObjType,strObjName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set sObjObject= SAPGuiSession("Session").SAPGuiWindow("SAP")
	Select Case strObjType
		Case "SAPGuiButton"
			If  sObjObject.SAPGuiButton(strObjName).Exist Then
				If sObjObject.SAPGuiButton(strObjName).GetROProperty("enabled")<> False Then
					sObjObject.SAPGuiButton(strObjName).Click
			    Else
					Call fRptWriteReport("Fail", "Clicking on SAPGuiButton", strObjName&" should be shown in enabled state")
				End If
		    Else
			  Call fRptWriteReport("Fail", "Clicking on SAPGuiButton", strObjName&" is not being shown in the application")
			End If
		
		Case "SAPGuiLabel"
			If  sObjObject.SAPGuiLabel(strObjName).Exist Then
				sObjObject.SAPGuiLabel(strObjName).SetFocus	
			Else
				Call fRptWriteReport("Fail", "Clicking on SAPGuiButton", strObjName&" is not being shown in the application")
			End If
			
		Case "SAPGuiTextArea"
				If sObjObject.SAPGuiTextArea(strObjName).Exist   Then
					sObjObject.SAPGuiTextArea(strObjName).DoubleClick
				Else
					Call fRptWriteReport("Fail", "Clicking on SAPGuiTextArea", strObjName&" is not being shown in the application")
				End If
			
		Case "Enter"
				SAPGuiSession("Session").SAPGuiWindow("SAP").SendKey ENTER									
		
		Case "F4"
				SAPGuiSession("Session").SAPGuiWindow("SAP").SendKey F4
		
		Case "F2"
				SAPGuiSession("Session").SAPGuiWindow("SAP").SendKey F2
		
		Case "F3"
				SAPGuiSession("Session").SAPGuiWindow("SAP").SendKey F3
	
	End Select
   
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fOpenHTMLReport()
'	Objective						:				Used to Open the HTML Report that is generated
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fOpenHTMLReport()

	Dim oReport
	Set oReport = CreateObject("WScript.Shell")
	TestVal = "iexplore.exe "&Environment("SUMMARYREPORTHTML")
	oReport.Run TestVal, 3, False
	Set oReport = Nothing
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fWindowSync()
'	Objective						:				Used to Wait for the Window Process to complete
'	Input Parameters				:				ObjControl,ObjType
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fWindowSync(ObjControl,ObjType)
	
	On Error Resume Next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	bWaitFlag = True
	iWait = 0
	iMaxWait = 100000
	Do While (iMaxWait > iWait)
		Select Case ObjType
			Case "SwfEdit"
				If ObjControl.WaitProperty("Visible","True") Then
					bWaitFlag = False
					Exit Do
				Else
					iWait = iWait + 1
					Wait 1
				End If

			Case "SwfButton"
				If ObjControl.WaitProperty("Visible","True") Then
					bWaitFlag = False
					Exit Do
				Else
					iWait = iWait + 1
					Wait 1
				End If

			Case "WinButton"
				If ObjControl.WaitProperty("Visible","True") Then
					bWaitFlag = False
					Exit Do
				Else
					iWait = iWait + 1
					Wait 1
				End If

			Case "SwfLabel"
				If ObjControl.Exist Then
					iWait = iWait + 1
					Wait 1
				Else
					bWaitFlag = False
					Exit Do
				End If

			End Select
		
		Loop
		Err.Clear

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fCtrl()
'	Objective						:				Used to Click CTRL + Any Key
'	Input Parameters				:				ObjControl,strValue
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fCtrl(ObjControl,strValue)
	
	ObjControl.Type micCtrlDwn + strValue + micCtrlUp
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetAvailableDiskSpace()
'	Objective						:				Used to check disk space on the Specified drive
'	Input Parameters				:				sDriveName
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fGetAvailableDiskSpace(sDriveName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set oDrive = objFso.GetDrive(sDriveName)
	sSpaceInBytes = oDrive.AvailableSpace 
	sSpaceInGB = sSpaceInBytes/1073741824
	Set objFso = Nothing
	Set oDrive = Nothing
	fGetAvailableDiskSpace = sSpaceInGB
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fRetriveValue()
'	Objective						:				Used to Retrvie Values from the Application
'	Input Parameters				:				sObjObject,sProperty
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fRetriveValue(sObjObject,sProperty)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
   Select Case sObjObject.GetROProperty("micClass")
		Case "SwfLabel","SwfEditor","SwfEdit","SwfComboBox"
			strRetrivedValue = sObjObject.GetROProperty(sProperty)
		
		Case "SwfListView"
			If sProperty = "" Then
				strRetrivedValue = sObjObject.GetItemsCount
			Else
				strRetrivedValue = sObjObject.GetROProperty(sProperty)
			End If
		
		Case "SwfList"
			strRetrivedValue = sObjObject.GetContent
	End select

	fRetriveValue = strRetrivedValue

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fSetDefaultPrinter()
'	Objective						:				Used to do functionality  set the default printer
'	Input Parameters				:				sPrinterName
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fSetDefaultPrinter(sPrinterName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set WshNetwork =CreateObject("WScript.Network")
	WshNetwork.SetDefaultPrinter(sPrinterName)
	Set WshNetwork = Nothing
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fDeleteFolder()
'	Objective						:				Used to delete the folder based on the Input Path
'	Input Parameters				:				sPath
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fDeleteFolder(sPath)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set deletefolder = fso.GetFolder(sPath)
	deletefolder.Delete(True) 
	set fso = Nothing
	Call fRptWriteReport("Pass","Deleting folder","Folder: "&sPath&" should be delected successfully")

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fCalcRunTime()
'	Objective						:				Used to Calculate the Time of Execution using Timer Method
'	Input Parameters				:				gSTime
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fCalcRunTime(gSTime)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	gEndTime = Timer
	intDuration = gEndTime - gSTime

	If intDuration >= 3600 Then
		intHR = Int (intDuration/3600)
		intHRTime = Int(intDuration Mod 3600)
		strValue = intHR&" hrs "
		
		If intHRTime >= 60 Then
		  intMIN = Int(intHRTime/60)
		  intSEC = Int(intHRTime Mod 60)
		  strValue = strValue&intMIN&" min "&intSEC&" sec."
		Else
		  intSEC = intHRTime
		  strValue = strValue&intSEC&" sec."
		End If

	Else If intDuration >= 60 Then
		intMIN = Int (intDuration/60)
		intSEC = Int(intDuration mod 60)
		strValue = strValue&intMIN&" min "&intSEC&" sec."
	
	Else
		intSEC = Int(intDuration)
		strValue = strValue&intSEC&" sec."	
		End If
	End If
	
	fCalcRunTime = strValue
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fCalcEventTime()
'	Objective						:				Used to Calculate the Time of Execution
'	Input Parameters				:				gSTime,gEndTime
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fCalcEventTime(gSTime,gEndTime)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    intDuration = DateDiff("s", gSTime, gEndTime)
	intHours = intDuration\3600
    intMinutes = (intDuration Mod 3600) \ 60
    intSeconds = intDuration Mod 60
    fCalcEventTime = intHours&":"&intMinutes&":"&intSeconds
    
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fRightClick()
'	Objective						:				Used to right click on Specific Object
'	Input Parameters				:				sObjObject,sValue
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fRightClick(sObjObject,sValue)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Select Case sObjObject.GetROProperty("micClass")
		Case "SwfListView","SwfTreeView","SwfTab"
			sObjObject.Select sValue,micRightBtn
	End select

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetFolderName()
'	Objective						:				Used to get the folder name from the TreeView
'	Input Parameters				:				sTestCaseName
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fGetFolderName(sTestCaseName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	strCount = SwfWindow("Stroz Review").SwfTreeView("Document Set Tree").GetItemsCount
	For i = 1 to strCount-1
		sValue = SwfWindow("Stroz Review").SwfTreeView("Document Set Tree").GetItem(i)
		sValueArray = Split(sValue,"(")	
		If Trim(sValueArray(0)) = Trim(sTestCaseName) Then
			fGetFolderName = sValue
			Exit For
		End If
	Next
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fSetTreeView()
'	Objective						:				Used to Select the Tree Node based on the Method Type
'	Input Parameters				:				sObjObject,sMethod,sValue
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fSetTreeView(sObjObject,sMethod,sValue)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Select Case sMethod
		Case "Expand"
			sObjObject.Expand sValue
		Case "Select"
			sObjObject.Select sValue
		Case "Collapse"
			sObjObject.Collapse sValue				
	End Select
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fDateConversion()
'	Objective						:				Used to Convert Date to DD-Mon-YYYY Format
'	Input Parameters				:				sDate
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fDateConversion(sDate)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	sDateFormat  = FormatDateTime(sDate, 1)
	sDateFormat1  = Split (Trim(sDateFormat),",",-1,1)
	sDt  = Right(Trim(sDateFormat1(1)),2)
	sMonth  = Left(Trim(sDateFormat1(1)),3)
	sDateformat  = sDt&"-"&sMonth&"-"&Trim(sDateFormat1(2))
	fDateConversion = sDateFormat
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fCreateNotepad()
'	Objective						:				Used to Create Notepad to check Pass / Fail
'	Input Parameters				:				sParameter,sStatus
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fCreateNotepad(sParameter,sStatus)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	If Environment("QC") = "No" Then
		strFileName = gstrProjectResultPath&"\"&sParameter&" -- "&sStatus&""&".txt"
		Const ForReading = 1, ForWriting = 2
		Set sObjFSO = CreateObject("Scripting.FileSystemObject")
		Set sObjNotepad = sObjFSO.CreateTextFile(strFileName,True)
		sObjNotepad.WriteLine Environment("DETAILEDREPORTHTML")
		sObjNotepad.Close
		Set sObjNotepad = Nothing
		Set sObjFSO = Nothing
	End If
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fGetLatestTraceFilePath()
'	Objective						:				Used to get the Latest Trace File path that is created by the application
'	Input Parameters				:				
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fGetLatestTraceFilePath()
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	sTracePath = "C:\Documents and Settings\"&Environment("UserName")&"\Local Settings\Application Data\Docuity\DocuityTemp\trace\"
	Dim LatestFile
	Set oFolder=CreateObject("scripting.FileSystemObject").GetFolder(sTracePath)
	For Each eachFile In oFolder.Files
		If LatestFile = "" Then
			Set LatestFile = eachFile
		Else If LatestFile.DateCreated < eachFile.DateCreated Then
			Set LatestFile = eachFile
			End If
		End If
	Next
	sTraceFilePath = sTracePath& LatestFile.Name
	fGetLatestTraceFilePath = sTraceFilePath
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fContextMenuCheck()
'	Objective						:				Used to Check for the Item in the ContextMenu
'	Input Parameters				:				sObjObject,sItemName
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fContextMenuCheck(sObjObject,sItemName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	fContextMenuCheck = sObjObject.GetItemProperty(sItemName,"Exists")
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				fRandomNumber()
'	Objective						:				This function is used to get random number
'	Input Parameters				:				intLowBound,intUpperBound,strText
'	Output Parameters			    :				Nil
'	Date Created					:				
'	UFT Version 			    	:				UFT 15.0
'	Pre-requisites					:				NIL  
'	Created By						:				Cigniti
'	Modification Date		        :		   		
'******************************************************************************************************************************************************************************************************************************************
Public Function fRandomNumber(intLowBound,intUpperBound,strText)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim intRand
	Randomize
	intRand = Int((intUpperBound - Cint(intLowBound) + 1) * Rnd + Cint(intLowBound))
	If strText<>"" Then
		intRand=strText&intRand
	End If
	'Return value
	fRandomNumber=intRand
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 				:					fInsertOutputValueToNewExcel
'	Objective							:					Used to insert previous file output value into new file input value
'	Input Parameters					:					intRow,strFieldName,strTcode,strTestName,strNewFieldName,sSheetName
'	Output Parameters					:					Nil
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   			
'******************************************************************************************************************************************************************************************************************************************		
Public Function fInsertOutputValueToNewExcel(intRow,strFieldName,strTcode,strTestName,strNewFieldName,sSheetName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	sRequiredValue = fGetSingleValue(strFieldName,strTcode,strTestName)
	Set fso = CreateObject("Scripting.FileSystemObject")
	gstrFile = gstrProjectTestdataPath&Environment("TestName")&"_Testdata.xls"
	If (fso.FileExists(gstrFile)) Then								
		Set  objExcel = CreateObject("Excel.Application")
		objExcel.UserControl = True
		objExcel.Application.DisplayAlerts = False
		objExcel.visible =   False
		objExcel.Workbooks.Open(gstrFile)
		objExcel.Sheets(sSheetName).Select
		intLastCol = objExcel.ActiveWorkbook.ActiveSheet.UsedRange.Columns.Count
		For iValue = 1 to intLastCol
				sColumnName = objExcel.ActiveWorkbook.ActiveSheet.Cells(1,iValue)
				If Trim(sColumnName) = strNewFieldName Then
						objExcel.ActiveWorkbook.ActiveSheet.Cells(intRow+1,iValue) = sRequiredValue
					Exit For
				End If
		Next
		objExcel.Selection.Columns.Autofit
		objExcel.Range("A1:J200").Select
		objExcel.Selection.Columns.Autofit
		objExcel.Range("A1").Select
		objExcel.ActiveWorkbook.Save
		objExcel.ActiveWorkbook.Close
		objExcel.Quit
		Set objExcel=Nothing
	End If
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 				:					fDateConversionMMDDYYY
'	Objective							:					Used to Convert Date to MM/DD/YYYY Format
'	Input Parameters					:					sSheetName
'	Output Parameters					:					Nil
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   			
'******************************************************************************************************************************************************************************************************************************************		
Public Function fDateConversionMMDDYYY(sDate)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	sDateformat  = Right("0" & Month(sDate), 2) & "/" & Right("0" & Day(sDate), 2) & "/" & Right(Year(sDate), 4)
	fDateConversionMMDDYYY = sDateFormat
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 				:					fCheckSheetExistance
'	Objective							:					Used to check sheet Existance
'	Input Parameters					:					sSheetName
'	Output Parameters					:					Nil
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   			
'******************************************************************************************************************************************************************************************************************************************		
Public Function fCheckSheetExistance(sSheetName)
 	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set fso = CreateObject("Scripting.FileSystemObject")
	gstrFile = gstrProjectTestdataPath&Environment("TestName")&"_Testdata.xls"
	If (fso.FileExists(gstrFile)) Then								
        Set  objExcel = CreateObject("Excel.Application")
		objExcel.UserControl = True
		objExcel.Application.DisplayAlerts = False
		objExcel.visible =   False
		Set objWB =objExcel.Workbooks.Open(gstrFile)
		bSheetExist = False
		For Each objWS in objWB.Worksheets
			If trim(objWS.Name)=trim(sSheetName) Then
				bSheetExist = True
				Exit For	
			End If
		Next
		objWB.close
		Set objWB = Nothing
		objExcel.Quit
		Set objExcel = Nothing
 	End If
 	
 	fCheckSheetExistance = bSheetExist
				
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 				:					fGetInputValue
'	Objective							:					Used to Take input from previous sheet
'	Input Parameters					:					sSheetName
'	Output Parameters					:					Nil
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   			
'******************************************************************************************************************************************************************************************************************************************		
Public Function fGetInputValue(strSheetName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	bSheetExist = fCheckSheetExistance(strSheetName)
	IF bSheetExist Then
		sDocumentNumber=fGetSingleValue("OutputValue",strSheetName,Environment("TestName"))
	Else
		Call fRptWriteReport("Fail", "Finding Excel Sheet for Input maaping", strSheetName&" Sheet is not Found in the"&Environment("TestName")&"_Testdata workbook")
		Environment("ERRORFLAG") = False
		ExitAction
	End If
	fGetInputValue = sDocumentNumber
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 				:					fCheckROExist
'	Objective							:					Used to check RunTime Object Existance
'	Input Parameters					:					strObjName,strObjType
'	Output Parameters					:					Nil
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   			
'******************************************************************************************************************************************************************************************************************************************		
Public Function fCheckROExist(strObjName,strObjType)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set sObjObject=SAPGuiSession("Session").SAPGuiWindow("SAP")
	Select Case strObjType	
		Case "SAPGuiEdit"
		   bExist = sObjObject.SAPGuiEdit(strObjName).Exist
		Case "SAPGuiButton"
		   bExist = sObjObject.SAPGuiButton(strObjName).Exist
		Case "SAPGuiCheckBox"
		   bExist = sObjObject.SAPGuiCheckBox(strObjName).Exist
		Case "SAPGuiTable"
		   bExist = sObjObject.SAPGuiTable(strObjName).Exist
		Case "SAPGuiTabStrip"
		   bExist = sObjObject.SAPGuiTabStrip(strObjName).Exist
		Case "SAPGuiRadioButton"
		   bExist = sObjObject.SAPGuiRadioButton(strObjName).Exist
		Case "SAPGuiTextArea"
		   bExist = sObjObject.SAPGuiTextArea(strObjName).Exist
		Case "SAPGuiComboBox"
		   bExist = sObjObject.SAPGuiComboBox(strObjName).Exist
	End Select
	fCheckROExist = bExist
  
 End Function
    	  
    	  
'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 					:					fWriteOutputValueInExcel
'	Objective							:					Used to Write Output Value In Excel Sheet
'	Input Parameters					:					ScritName,intRow,StrValue,sSheetName,strColumnName
'	Output Parameters					:					Nil
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   			
'******************************************************************************************************************************************************************************************************************************************		
Public Function fWriteOutputValueInExcel(ScritName,intTCID,StrValue,sSheetName,strColumnName)

	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	intRow = 1
	Set fso = CreateObject("Scripting.FileSystemObject")
'	sFile = strProjectTestdataPath&ScritName&"_Testdata.xls"
	gstrFile = gstrProjectTestdataPath&ScritName&"_TD.xls"
	
	If (fso.FileExists(gstrFile)) Then								
		Set  objExcel = CreateObject("Excel.Application")
		objExcel.UserControl = True
		objExcel.Application.DisplayAlerts = False
		objExcel.visible =   False
		objExcel.Workbooks.Open(gstrFile)
		objExcel.Sheets(sSheetName).Select
		intLastCol = objExcel.ActiveWorkbook.ActiveSheet.UsedRange.Columns.Count
		intLastrow = objExcel.ActiveWorkbook.ActiveSheet.UsedRange.rows.Count
		
		For irowValue = 1 to intLastrow
			srowValue = objExcel.ActiveWorkbook.ActiveSheet.Cells(irowValue,2)
			If srowValue = intTCID  Then
				intRow = irowValue
				Exit For
			End If
		Next		
		For iValue = 1 to intLastCol
				sColumnName = objExcel.ActiveWorkbook.ActiveSheet.Cells(1,iValue)
				
				If Trim(sColumnName) = strColumnName Then
						
							If IsNumeric(StrValue) Then
							 	objExcel.ActiveWorkbook.ActiveSheet.Cells(intRow,iValue).NumberFormat = "General"
							 	
							End If
							objExcel.ActiveWorkbook.ActiveSheet.Cells(intRow,iValue) = StrValue
							

						Call fRptWriteReport("Pass", "Output value",StrValue&" is enterd into " &ScritName&" test data sheet")
						fWriteOutputValueInExcel = True
					Exit For
				End If
		Next
		objExcel.Selection.Columns.Autofit
		objExcel.Range("A1:J200").Select
		objExcel.Selection.Columns.Autofit
		objExcel.Range("A1").Select
		objExcel.ActiveWorkbook.Save
		objExcel.ActiveWorkbook.Close
		objExcel.Quit
		Set objExcel=Nothing
	else
		Call fRptWriteReport("Fail", StrValue,"Output value"&" is not enterd into " &ScritName&" test data sheet")
		Environment("ERRORFLAG") = False
		fWriteOutputValueInExcel = False
	End If ' End If for Check for the existience of the File	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 					:					fExcelStringCompare
'	Objective							:					Used to verifying values at excel
'	Input Parameters					:					sPath,sString
'	Output Parameters					:					Nil
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   			
'******************************************************************************************************************************************************************************************************************************************		
Public Function fExcelStringCompare(strFileName,sString)
	
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim sFSO
	Dim objExcel
	Dim oData
	Dim iFoundCell
	Dim sFind
	Dim gstrFile
	
	sFind = sString
	gstrFile=strFileName
	Set sFSO = CreateObject("Scripting.FileSystemObject")
	Set objExcel = CreateObject("Excel.Application")
	Set oData = objExcel.Workbooks.Open(gstrResourcesPathForData& "\"&gstrFile&".xls")	
	Set iFoundCell = oData.ActiveSheet.Range("A4:AI100").Find(sFind)
	If Not iFoundCell Is Nothing Then				
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Request field verification","Request field is found--"&sFind & "in row No:"&iFoundCell.Row)
		fExcelStringCompare = True
	Else				
		Call fRptWriteReport("Fail", "Request field verification" ,"Request field is NOT found")	
		Environment("ERRORFLAG") = False
		fExcelStringCompare = False
	End If				
	Set gstrFile = nothing
	Set sFind = nothing
	Set iFoundCell = nothing
	Set oData = Nothing
	Set objExcel = Nothing
	Set sFSO = Nothing
	On error goto 0
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 					:					fSaveExcelFile
'	Objective							:					Used to Save the excel log file at resource
'	Input Parameters					:					
'	Output Parameters					:					NIL
'	Date Created						:					16-May-2017
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   			
'******************************************************************************************************************************************************************************************************************************************		
Public Function fSaveExcelFile(strFileName)

	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim obj, gstrFile
	Set obj = CreateObject("Scripting.FileSystemObject")			
	gstrFile=trim(strFileName)
	If  obj.FileExists(gstrResourcesPathForData& "\"&gstrFile&".xls") Then
		obj.DeleteFile(gstrResourcesPathForData& "\"&gstrFile&".xls")	
	End If
	Set obj= Nothing
		Dialog("WindowsInternetExplorer").WinButton("SaveAs").Click
		wait 3
		Dialog("dlgSave As").WinEdit("edtFileName").Set gstrResourcesPathForData& "\"&gstrFile&".xls"
		Dialog("dlgSave As").WinButton("btnSave").Click		
	On error goto 0
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name							:						  		fGetTestDataBySheet
'	Objective								:							 	Used to read data from multiple rows in data sheet
'	Input Parameters						:							  	strSheetName
'	Output Parameters						:							   	Data
'	Date Created							:								
'	UFT Version								:								UFT 15.0
'	Pre-requisites							:								NIL  
'	Created By								:								Cigniti
'	Modification Date						:		   
'******************************************************************************************************************************************************************************************************************************************
 Public Function fGetTestDataBySheet(strSheetName)
 	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	sSheetName = strSheetName
	'gstrFile = gstrProjectTestdataPath & Environment("TestName") & "_TD.xls"
	Set Data = CreateObject("Scripting.Dictionary")
	Data.RemoveAll

	sQuery =  "SELECT * FROM [" & sSheetName & "$] Where Run='YES'"	                    
'	sQuery =  "SELECT * FROM [" & sSheetName & "$] Where Run = 'YES' or Run is Null"	                    
	Set DB_CONNECTION=CreateObject("ADODB.Connection")
	
	 DB_CONNECTION.CursorLocation = 3
'  DB_CONNECTION.CursorType = adOpenDynamic
  'DB_CONNECTION.LockType = adLockOptimistic
	
	DB_CONNECTION.Open("DBQ=" & gstrFile & ";DefaultDir=C:\;Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DriverId=790;FIL=excel 8.0;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\matdsn2.dsn;MaxScanRows=8;PageTimeout=5;ReadOnly=0;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;")
	Set Record_Set1=DB_CONNECTION.Execute(sQuery)
	Data.Add "TCID", Record_Set1.Fields.Item("TCID").Value	
	iRowCount=1
	Do While Not Record_Set1.EOF
		Set TempData=CreateObject("Scripting.Dictionary")
		TempData.RemoveAll
		iColumnCount = 0
		For Each Field In Record_Set1.Fields
			sColumnName = Field.Name
			iRowValue = Record_Set1(iColumnCount)
			If IsNull(iRowValue) Then
				iRowValue = ""
			End If
			TempData.Add sColumnName,iRowValue
			iColumnCount = iColumnCount + 1
		Next
		Data.Add iRowCount, TempData
		Record_Set1.MoveNext
		iRowCount=iRowCount+1
	Loop
	Record_Set1.Close
	Set Record_Set1=Nothing
	DB_CONNECTION.Close
	Set DB_CONNECTION=Nothing
	Set fGetTestDataBySheet = Data
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name							:						  		fGetRowCountFromTestData
'	Objective								:							 	Used to retrive the RowCount from the Test Data File
'	Input Parameters						:							  	gstrFile,sItemName
'	Output Parameters						:							   	NIL
'	Date Created							:								
'	UFT Version								:								UFT 15.0
'	Pre-requisites							:								NIL  
'	Created By								:								Cigniti
'	Modification Date						:		   
'******************************************************************************************************************************************************************************************************************************************
Public Function fGetRowCountFromTestData(gstrFile,sItemName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	sQuery =  "SELECT * FROM["&sItemName&"] WHERE Run = 'Y'"
	Set DB_CONNECTION=CreateObject("ADODB.Connection")

	iCheck = Instr(1,sItemName,"$")
	If iCheck = 0 Then
		sItemName = sItemName&"$"
	End If
	
	If Environment("TRANSACTIONRANGE") = "" Then
		sQuery =  "SELECT * FROM ["&sItemName&"] WHERE Run = 'Y'"
	Else 
		sQueryCondition= fCreateQuery(Environment("TRANSACTIONRANGE"))
		sQuery =  "SELECT * FROM ["&sItemName&"] WHERE Run = 'Y' and "&sQueryCondition
	End If

	DB_CONNECTION.Open "DBQ="&gstrFile&";DefaultDir=C:\;Driver={Driver do Microsoft Excel(*.xls)};DriverId=790;FIL=excel 8.0;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\matdsn2.dsn;MaxScanRows=8;PageTimeout=5;ReadOnly=0;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
                
	Set Record_Set1=DB_CONNECTION.Execute(sQuery)
	Set Record_Set2=DB_CONNECTION.Execute(sQuery)
	iRowCount = 0

	Do While Not Record_Set2.EOF
		For Each Field In Record_Set1.Fields
			If IsNull(iRowValue) Then
				iRowValue = ""
			End If
		Next
		Record_Set2.MoveNext
		iRowCount = iRowCount + 1
	Loop

	Record_Set1.Close
	Set Record_Set1=Nothing
	Record_Set2.Close
	Set Record_Set2=Nothing
	DB_CONNECTION.Close
	Set DB_CONNECTION=Nothing
	fGetRowCountFromTestData = iRowCount
	
End Function


'***********************************************************************************************************************************
'	Function Name						:					fExcelComparison
'	Objective							:					Compare two excel sheets and Highlight differences
'	Input Parameters					:					strSrcPath, sExcelFileName,sSheetName,sCompareSrcPath, sCompareExcelFileName, sCompareSheetName
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					Files Should be avilable in specified path  
'	Created By							:					Cigniti
'	Modification Date					:		   
'**********************************************************************************************************************************	  
Public Function fExcelComparison(Byval strSrcPath,Byval  sExcelFileName,Byval sSheetName,Byval sCompareSrcPath,Byval sCompareExcelFileName,Byval sCompareSheetName)
	
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objExcel
	Dim objWorkbook1
	Dim objWorkbook2
	Dim objWorksheet1
	Dim objWorksheet2
	Dim sExcelPath1
	Dim sExcelPath2
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
	sExcelPath1=strSrcPath+"\"+sExcelFileName+".xlsx"
	Set objWorkbook1= objExcel.Workbooks.Open(sExcelPath1)
	sExcelPath2=sCompareSrcPath+"\"+sCompareExcelFileName+".xlsx"
	Set objWorkbook2= objExcel.Workbooks.Open(sExcelPath2)
	Set objWorksheet1= objWorkbook1.Worksheets(sSheetName)
	Set objWorksheet2= objWorkbook2.Worksheets(sCompareSheetName)

	For Each cell In objWorksheet1.UsedRange
	   If Trim(cell.Value) <> Trim(objWorksheet2.Range(cell.Address).Value) Then
	       cell.Interior.ColorIndex = 3 
	   Else
	       cell.Interior.ColorIndex = 0
	   End If
	Next
		
	objWorkbook1.Save
	objWorkbook2.Save
	objWorkbook1.Close
	objWorkbook2.Close
	Set objWorkbook1=Nothing
	Set objWorkbook2=Nothing
	objExcel.Quit
	Set objExcel= Nothing
		
	On error goto 0
	
End Function


'********************************************************************************************************************************************************
'	Function Name						:					fGetExcelSheetCount
'	Objective							:					Count Sheets in an Excel file
'	Input Parameters					:					strExcelPath("C:\Desktop\ExcelSheet1")
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'*******************************************************************************************************************************************************
Public Function fGetExcelSheetCount(Byval strExcelPath )
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objExcel
	Dim objWorkbook
    
    Set objExcel = CreateObject("Excel.Application")
    Set objWorkbook = objExcel.Workbooks.Open(strExcelPath+".xlsx")
    fGetExcelSheetCount = objWorkbook.Worksheets.Count
   	objWorkbook.Close
   	objExcel.Quit
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    
End Function


'***********************************************************************************************************************************
'	Function Name						:					fCreateExcelFile
'	Objective							:					Creating an Excel File and Saved in specified location
'	Input Parameters					:					strPath,strExcelName ("C:\Desktop\" ; "ExcelFileName")
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'**********************************************************************************************************************************
Public Function fCreateExcelFile(Byval strPath,Byval strExcelName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim obj
	Dim obj1
	Set obj = createobject("Excel.Application")  		
	obj.visible=True                                    
	Set obj1 = obj.Workbooks.Add()       				
	obj1.SaveAs strPath  +strExcelName+".xlsx" 			
	obj1.Close                                          
	obj.Quit                                            
	Set obj1=Nothing                                 	
	Set obj=Nothing 								 	

End Function


'*********************************************************************************************************************************************************
'	Function Name						:					fGetCellDataInExcelSheet
'	Objective							:					Reading data from a specific cell in Excel
'	Input Parameters					:					strExcelPath, strExcelSheetName, intRow, intCol ("C:\Desktop\", "ExcelFileName",1,1)
'	Output Parameters					:					gGetExcelCellData  -Perticular cell value 
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'*********************************************************************************************************************************************************
Public Function fGetCellDataInExcelSheet(Byval strExcelPath,Byval strExcelSheetName,Byval intRow,Byval intCol)
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    Dim objExcel
    Dim objWorkbook
    Dim objSheet
    Dim strValue
    
    Set objExcel = CreateObject("Excel.Application")
    Set objWorkbook = objExcel.Workbooks.Open(strExcelPath+".xlsx")
    Set objSheet  = objWorkbook.Worksheets(strExcelSheetName) 
    strValue = objSheet.Cells(intRow, intCol)
    fGetCellDataInExcelSheet = strValue
    objWorkbook.Close
    objExcel.Quit
    Set objSheet  = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    
End Function


'*******************************************************************************************************************************************************
'	Function Name						:					fReadEachCellDataInExcel
'	Objective							:					Reading Cell Values One by One for all Rows of an Excel Sheet
'	Input Parameters					:					strExcelPath, strExcelSheetName("C:\Desktop\", "ExcelFileName")
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'*******************************************************************************************************************************************************
Public Function fReadEachCellDataInExcel(Byval strExcelPath,Byval strExcelSheetName)
   'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objExcel
	Dim objWorkbook
	Dim objSheet
	Dim ColCount
	Dim RowCount
	Dim iCount
	Dim jCount
	Dim fieldvalue
	
	Set objExcel = CreateObject("Excel.Application")
    Set objWorkbook = objExcel.Workbooks.Open(strExcelPath+".xlsx")
    Set objSheet = objWorkbook.Worksheets(strExcelSheetName)
    ColCount = objSheet.UsedRange.Columns.Count
    RowCount = objSheet.UsedRange.Rows.Count
    For iCount = 1 To RowCount
        For jCount = 1 To ColCount
            fieldvalue = objSheet.Cells(iCount,jCount)
            MsgBox fieldvalue
        Next
   Next
   objWorkbook.Close
   objExcel.Quit
   Set objSheet  = Nothing
   Set objWorkbook = Nothing
   Set objExcel = Nothing
   
End Function


'********************************************************************************************************************************************************
'	Function Name						:					fSetCellDataInExcel
'	Objective							:					Writing data to a specific cell in Excel
'	Input Parameters					:					strExcelPath, strExcelSheetName,intRow, intCol, strValue("C:\Desktop\", "ExcelFileName",2,2,"Test Data")
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'********************************************************************************************************************************************************
Public Function fSetCellDataInExcel(Byval strExcelPath,Byval strExcelSheetName,Byval intRow,Byval intCol,Byval strValue)	
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objExcel
	Dim objWorkbook
	Dim objSheet			
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False
    Set objWorkbook = objExcel.Workbooks.Open(strExcelPath+".xlsx")
    Set objSheet  = objWorkbook.Worksheets(strExcelSheetName)   'Or pass sheet number integer value 1,2,etc
    objSheet.Cells(intRow, intCol).value = strValue
    objWorkbook.Save
    objWorkbook.Close
    objExcel.Quit
    Set objSheet  = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
	    
End Function


'***********************************************************************************************************************************
'	Function Name						:					fSaveExcelFile
'	Objective							:					Save excel file in specified path along with Date and Time stamp
'	Input Parameters					:					sSrcExcelPath, sExcelFileName, sDestResultsPath 
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					Nil
'	Created By							:					Cigniti
'	Modification Date					:		   
'**********************************************************************************************************************************	  
Public Function fSaveExcelFile(Byval sSrcExcelPath,Byval sExcelFileName,Byval sDestResultsPath )	
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objExcel
	Dim objWorkbook1
	Set objExcel=CreateObject("Excel.Application")
	objExcel.Visible=TRUE
	strCurrentDateAndTime= Replace(Replace(Replace(now(),":",""),"/","")," ","")
	strSrcPath=sSrcExcelPath+"\"+sExcelFileName+".xlsx"
	sDestResultsPath=sDestResultsPath+"\"+sExcelFileName+"_"+strCurrentDateAndTime+".xlsx"
	Set objWorkbook1= objExcel.Workbooks.Open(strSrcPath)
	objWorkbook1.SaveAs(sDestResultsPath)
	objWorkbook1.Close
	Set objWorkbook1=Nothing
	objExcel.Quit
	Set objExcel= Nothing
	
End Function


'***********************************************************************************************************************************
'	Function Name						:					fGetRowCounts
'	Objective							:					Get Row count in Specified Sheet in an Excel file
'	Input Parameters					:					strExcelPath("C:\Desktop\")
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'**********************************************************************************************************************************
Public Function fGetRowCounts(Byval strExcelPath,Byval strExcelSheetName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If		
	Dim objExcel
	Dim objWorkbook
	Dim objSheet
	
	Set objExcel=CreateObject("Excel.Application")
	Set objWorkbook=objExcel.WorkBooks.Open(strExcelPath+".xlsx")
	Set objSheet=objWorkbook.WorkSheets(strExcelSheetName)
	fGetRowCounts=objSheet.usedrange.rows.count 'Get Row count
	objWorkbook.Close
	objExcel.Quit
	Set objExcel=Nothing

End Function 


'***********************************************************************************************************************************
'	Function Name						:					fAddSheet
'	Objective							:					Add a sheet in an Excel file
'	Input Parameters					:					strExcelPath , strNewSheetName ("C:\Desktop\","Sheet1")
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					Nil
'	Created By							:					Cigniti
'	Modification Date					:		   
'**********************************************************************************************************************************
Public Function fAddSheet (Byval strExcelPath,Byval strNewSheetName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim obj
	Dim obj1
	Dim obj2
	
	Set obj = createobject("Excel.Application")   		
	obj.visible=True                                    
	Set obj1 = obj.Workbooks.open(strExcelPath+".xlsx")    
	Set obj2=obj1.sheets.Add  							   
	obj2.name=strNewSheetName 							   
	obj1.Save												
	obj1.Close                                              
	obj.Quit                                                  
	Set obj1=Nothing                                 
	Set obj2 = Nothing                                                        
	Set obj=Nothing                                 

End Function


'***********************************************************************************************************************************
'	Function Name						:					fDeleteSheet
'	Objective							:					Delete a sheet in an Excel file
'	Input Parameters					:					strExcelPath , strDeleteSheetName ("C:\Desktop\","Sheet1")
'	Output Parameters					:					NIL
'	Date Created						:				
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'**********************************************************************************************************************************
Public Function fDeleteSheet (Byval strExcelPath,Byval strDeleteSheetName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim obj
	Dim obj1
	Dim obj2
	
	Set obj = createobject("Excel.Application")   		
	obj.visible=True                                    
	Set obj1 = obj.Workbooks.open(strExcelPath+".xlsx")    
	Set obj2= obj1.Sheets(strDeleteSheetName)  				
	obj2.Delete       										
	obj1.Save
	obj1.Close                                              
	obj.Quit                                                  
	Set obj1=Nothing                                 
	Set obj2 = Nothing                                                   
	Set obj=Nothing                             

End Function


'***********************************************************************************************************************************
'	Function Name						:					fCopyAndPastDatafromOneExcelToAnotherExcel
'	Objective							:					Copying and Pasting Data from one Excel file to Another Excel File
'	Input Parameters					:					sSourceExcelPath, sSourceSheetName, sDestPath, sDestSheetName ("C:\Desktop\WorkBook1","Sheet1","C:\Desktop\TestFolder/WorkBook2","Sheet1")
'	Output Parameters					:					Nil
'	Date Created						:					
'	UFT Version							:					
'	Pre-requisites						:					UFT 15.0
'	Created By							:					Cigniti
'	Modification Date					:		   
'**********************************************************************************************************************************
Public Function fCopyAndPastDatafromOneExcelToAnotherExcel(ByVal sSourceExcelPath,ByVal sSourceSheetName,ByVal sDestPath,ByVal sDestSheetName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim obj
	Dim obj1
	Dim obj2
	
	Set obj = createobject("Excel.Application")   
	obj.visible=True                                    
	Set obj1 = obj.Workbooks.open(sSourceExcelPath+".xlsx")    
	Set obj2 = obj.Workbooks.open(sDestPath+".xlsx")    
	obj1.Worksheets(sSourceSheetName).usedrange.copy  
	obj2.Worksheets(sDestSheetName).usedrange.pastespecial  
	obj1.Save                                              
	obj2.Save                                              
	obj1.Close                                             
	obj2.Close 
	obj.Quit                                                 
	Set obj1=Nothing                                
	Set obj2 = Nothing                              
	Set obj=Nothing          

End Function 
	

'***********************************************************************************************************************************
'	Function Name						:					fFindSheetExistence
'	Objective							:					Verify Excel sheet Existence   
'	Input Parameters					:					strExcelPath,strExcelSheetName ("C:\Desktop\WorkBook1","Sheet1")
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFt 15.0
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'**********************************************************************************************************************************	

Public Function fFindSheetExistence(Byval strExcelPath,Byval strExcelSheetName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If		
	Dim ObjExcel
	Dim ObjWorkBook
	Dim objWorksheet
	
	Set ObjExcel =  CreateObject("Excel.Application")
	ObjExcel.DisplayAlerts = False
	Set ObjWorkBook = ObjExcel.WorkBooks.Open(strExcelPath+".xlsx")
	On Error Resume Next
	Set objWorksheet = ObjWorkBook.Sheets.Item(strExcelSheetName)
	boolRC = Err.Number <> 0
	On Error GoTo 0
	If boolRC Then
		Reporter.ReportEvent micFail, "Get sheet", "Failed to retrieve worksheet"
		fFindSheetExistence=1
	Else
		Reporter.ReportEvent micDone, "Get sheet", "Successfully retrieved object's instance"
		fFindSheetExistence=0
	End If			
	ObjWorkBook.Save
	ObjWorkBook.Close
	ObjExcel.Quit
	Set ObjWorkBook = Nothing
	Set ObjExcel = Nothing

End Function
	

'***********************************************************************************************************************************
'	Function Name						:					fReadAllDataInExcelSheet
'	Objective							:					Read all data in specified excel sheet  
'	Input Parameters					:					strExcelPath,strExcelSheetName ("C:\Desktop\WorkBook1","Sheet1")
'	Output Parameters					:					RowData
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'**********************************************************************************************************************************	

Public Function fReadAllDataInExcelSheet(Byval strExcelPath,Byval strExcelSheetName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objexcel
	Dim objWorkbook
	Dim objWorksheet
	Dim intRowCount
	Dim intColumnCount
	Dim intRow
	Dim intColumn
	Dim RowData
	
	'Create Excel Object 
	Set objexcel=createobject("Excel.application")
	'Make it Visible  
	objexcel.Visible=True  
	'Open Excel File
	Set objWorkbook=objexcel.Workbooks.open(strExcelPath+".xlsx")
	'Get Control on Sheet
	Set objWorksheet=objexcel.Worksheets.Item(strExcelSheetName)
	'Get Used Row and Column Count
	intRowCount=objWorksheet.usedrange.rows.count
	intColumnCount=objWorksheet.usedrange.columns.count
	'Loop through Rows
	For intRow=1 to intRowCount
		'Loop through Columns
		For intColumn=1 to intColumnCount
			'Get Cell Data
			RowData=RowData&objWorksheet.cells(intRow,intColumn)&vbtab
		Next
		RowData=RowData&vbcrlf
	Next
	fReadAllDataInExcelSheet=RowData
	
	'Close Workbook  
	objWorkbook.Close  
	
	'Quit from Excel Application  
	objexcel.Quit  
	
	'Release Variables  
	Set objWorksheet=Nothing
	Set objWorkbook=Nothing
	Set objexcel=Nothing
	
End Function
	

'***********************************************************************************************************************************
'	Function Name						:					fxlsDeleteColumnRange
'	Objective							:					Delete Columns based on user input
'	Input Parameters					:					sSrcExcelPath,sExcelFileName, sSheetName, sStartCol, sEndCol
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NIL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'**********************************************************************************************************************************
Public Function fxlsDeleteColumnRange (ByVal sSrcExcelPath,Byval sExcelFileName, ByVal sSheetName, ByVal sStartCol,ByVal sEndCol) 
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim strCurrentDateAndTime
	Dim objExcel
	Dim strSrcPath
	Dim objWorkbook
	Dim objWorksheet
	
	strCurrentDateAndTime= Replace(Replace(Replace(now(),":",""),"/","")," ","") 
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	strSrcPath=sSrcExcelPath+"\"+sExcelFileName+".xlsx"
	Set objWorkbook= objExcel.Workbooks.Open(strSrcPath)
	Set objWorksheet= objWorkbook.Worksheets(sSheetName)
	
	'Delete row range
	objWorksheet.Columns(sStartCol + ":" + sEndCol).Delete
	'Save new book to Excel file
	objWorkbook.Save
	'Close the xls file
	objWorkbook.Close()
	objExcel.Quit
	Set objWorkbook = Nothing
	set objExcel = Nothing

End Function


'******************************************************************************************************************************************************************************
'	Function Name						:		fxlsDeleteRowRange()
'	Objective							:		Delete Rows based on user input
'	Input Parameters					:		sSrcExcelPath,sExcelFileName, sSheetName, sStartCol, sEndCol
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fxlsDeleteRowRange (ByVal sSrcExcelPath,Byval sExcelFileName,ByVal sSheetName,ByVal sStartRow,ByVal sEndRow) 'Create Excel object
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim strCurrentDateAndTime
	Dim objExcel
	Dim strSrcPath
	Dim objWorkbook1
	Dim objWorksheet1
	
	strCurrentDateAndTime= Replace(Replace(Replace(now(),":",""),"/","")," ","") 
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	strSrcPath=sSrcExcelPath+"\"+sExcelFileName+".xlsx"
	
	Set objWorkbook1= objExcel.Workbooks.Open(strSrcPath)
	Set objWorksheet1= objWorkbook1.Worksheets(sSheetName)
	
	'Delete row range
	objWorksheet1.Rows(sStartRow +":"+ sEndRow).EntireRow.Delete
	
	'Save new book to Excel file
	objWorkbook1.Save
	
	'Close the xls file
	objWorkbook1.Close()
	objExcel.Quit
	Set objWorkbook1 = Nothing
	set objExcel = Nothing

End Function 


'******************************************************************************************************************************************************************************
'	Function Name						:		fGetOutputData()
'	Objective							:		Used to retrive Data from excel
'	Input Parameters					:		strTestReport, strSheet, strColName 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fGetOutputData(Byval strTestReport, Byval strSheet , Byval strColName)
	
	On Error Resume Next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.DisplayAlerts = False
	Set objWorkbook= objExcel.Workbooks.Open(strTestReport, False, False)
	Set objSheet = objExcel.ActiveWorkbook.Worksheets(strSheet)
	
	'Get the number of used rows
	nUsedRows = objSheet.UsedRange.Rows.Count
	
	'Get the number of used columns
	nUsedCols = objSheet.UsedRange.Columns.Count
	
	'Get the topmost row that has data
	nTop = objSheet.UsedRange.Row
	
	'Get leftmost column that has data
	nLeft = objSheet.UsedRange.Column
	
	intRow = 3
	Set objCells = objSheet.Cells
		
	'Getting the col number of Upgrade From
	For inti = 0 To nUsedCols-1 Step 1
		If Instr(1, objCells(4, inti + nLeft).Value,strColName)> 0 then
			fGetOutputData = objCells(5, inti + nLeft)
			Exit For
		End If	
	Next
	
	objWorkbook.Save
	objWorkbook.Close
	objExcel.Quit
	
End  Function


'******************************************************************************************************************************************************************************
'	Function Name						:		fVerifyingFINMARTReport()
'	Objective							:		Used to Verify Fimart Reports
'	Input Parameters					:		strTestReport, strSheet  
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fVerifyingFINMARTReport(Byval strTestReport, Byval strSheet)
	
	On Error Resume Next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.DisplayAlerts = False
	Set objWorkbook= objExcel.Workbooks.Open(strTestReport, False, False)
	Set objSheet = objExcel.ActiveWorkbook.Worksheets(strSheet)
	
	'Get the number of used rows
	nUsedRows = objSheet.UsedRange.Rows.Count
	
	'Get the number of used columns
	nUsedCols = objSheet.UsedRange.Columns.Count
	
	'Get the topmost row that has data
	nTop = objSheet.UsedRange.Row
	
	'Get leftmost column that has data
	nLeft = objSheet.UsedRange.Column
	
	intRow = 3
	Set objCells = objSheet.Cells
	
	'Getting the col number of Upgrade From
	For inti = 0 To nUsedCols-1 Step 1
		If(objCells(intRow + nTop, inti + nLeft).Value="Upgrade From") then
			intUpgradeFromCol=inti
		End If	
	Next
		
	'Getting the col number of Switch Direction
	For inti = 0 To nUsedCols-1 Step 1
		If(objCells(intRow + nTop, inti + nLeft).Value="Switch Direction") then
			intSwitchDirectionCol=inti
		End If	
	Next
		
	'Getting the col number of Offer Detail
	For inti = 0 To nUsedCols-1 Step 1
		If(objCells(intRow + nTop, inti + nLeft).Value="Offer Detail") then
			intOfferDetailCol=inti
		End If	
	Next
		
	'Getting the col number of Salesbook Cust Seg and Lic Type
	For inti = 0 To nUsedCols-1 Step 1
		If(objCells(intRow + nTop, inti + nLeft).Value="Salesbook Cust Seg and Lic Type") then
			intSalesbookCSLTCol=inti
		End If	
	Next

	'Getting the col number of Switch Type
	For inti = 0 To nUsedCols-1 Step 1
		If(objCells(intRow + nTop, inti + nLeft).Value="Switch Type") then
			intSwitchTypeCol=inti
		End If	
	Next
		
	'Getting the col number of M2S Indicator
	For inti = 0 To nUsedCols-1 Step 1
		If(objCells(intRow + nTop, inti + nLeft).Value="M2S Indicator") then
			intM2SCol=inti
		End If	
	Next
		
	'Getting the Settlement Start Date
	For inti = 0 To nUsedCols-1 Step 1
		If(objCells(intRow + nTop, inti + nLeft).Value="Settlement Start Date") then
			intSettlementStartDate=inti
		End If	
	Next
		
	'Getting the Original Order Date
	For inti = 0 To nUsedCols-1 Step 1
		If(objCells(intRow + nTop, inti + nLeft).Value="Original Order Date") then
			intOriginalOrderDate=inti
		End If	
	Next

	'Looping through all the rows
	For introws = 4 To nUsedRows-2 Step 1
	
		strSwitchDirectionValue= objCells(intRows + nTop, intSwitchDirectionCol + nLeft).Value				
		'Verifying Switch Type
		If NOT (objCells(intRows + nTop, intSwitchTypeCol + nLeft).Value="Switch At Renewal") Then
			objCells(intRows + nTop, intSwitchTypeCol + nLeft).Interior.ColorIndex = 3	
			Call fRptWriteReport("Fail","Verify "& chr(34) &"Switch Type"& chr(34) &" column value as "& chr(34) & "Switch At Renewal" & chr(34) &" in row no: "& intRows + nTop, "Did not meet the condition and Acutal Value " & objCells(intRows + nTop, intSwitchTypeCol + nLeft).Value)						
		End If			
				
		'Verifying M2S Indicator
		If NOT (objCells(intRows + nTop, intM2SCol + nLeft).Value="M2S Migration") Then
			objCells(intRows + nTop, intM2SCol + nLeft).Interior.ColorIndex = 3	
			Call fRptWriteReport("Fail","Verify "& chr(34) &"M2S Indicator"& chr(34) &" column value as "& chr(34) & "M2S Migration" & chr(34) &" in row no_"& intRows + nTop ,  "Did not meet the condition and Acutal Value " & objCells(intRows + nTop, intM2SCol + nLeft).Value)
		End If

		If strSwitchDirectionValue="Switch From" Then						
			
			'Verifying the Upgrade From
			cellValue=objCells(intRows + nTop, intUpgradeFromCol + nLeft).Value
			If NOT ((cellValue="*") OR (IsEmpty(cellValue))) Then
				objCells(intRows + nTop, intUpgradeFromCol + nLeft).Interior.ColorIndex = 3	
				Call fRptWriteReport("Fail","Verify "& chr(34) & "Upgrade From" & chr(34) &" column should value as either "& chr(34) &"*"& chr(34) &" or "& chr(34) &"blank"& chr(34) &" When Switch Direction is "& chr(34) &"Switch From"& chr(34) &" in row no_"& intRows + nTop, "There is a mismatch found as the actual value is " & objCells(intRows + nTop, intUpgradeFromCol + nLeft).Value &". It should be as either "& chr(34) &"*"& chr(34) &" or "& chr(34) &"blank"& chr(34))
			End If
			
			'Verifying the Offer Detail
			cellValue1=objCells(intRows + nTop, intOfferDetailCol + nLeft).Value
			If NOT (Instr(1,cellValue1,"Maintenance")>0) Then
				objCells(intRows + nTop, intOfferDetailCol + nLeft).Interior.ColorIndex = 3	
				Call fRptWriteReport("Fail","Verify "& chr(34) &"Offer Detail" & chr(34) &" column should contain value as "& chr(34) &"Maintenance" & chr(34) & " When Switch Direction is "& chr(34) &"Switch From"& chr(34) &" in row no_"& intRows + nTop, "There is a mismatch found as the actual value is " & objCells(intRows + nTop, intOfferDetailCol + nLeft).Value &". It should be "& chr(34) &"Maintenance"& chr(34) &" as per expectation.")
			End If
						
		ElseIf strSwitchDirectionValue="Switch To" Then 
			'Verifying the Offer Detail
			cellValue2=objCells(intRows + nTop, intOfferDetailCol + nLeft).Value
			If NOT (Instr(1,cellValue2,"Desktop Subscription")>0 OR Instr(1,cellValue2,"Network Subscription")>0) Then
				objCells(intRows + nTop, intOfferDetailCol + nLeft).Interior.ColorIndex = 3	
				Call fRptWriteReport("Fail","Verify "& chr(34) &"Offer Detail" & chr(34) &" column should contain value as "& chr(34) &"Desktop Subscription or Network Subscription" & chr(34) & " When Switch Direction is "& chr(34) &"Switch To"& chr(34) &" in row no_"& intRows + nTop, "There is a mismatch found as the actual value is "& objCells(intRows + nTop, intOfferDetailCol + nLeft).Value &". It should be "& chr(34)&"Desktop Subscription"& chr(34) &"or "& chr(34)&"Network Subscription"& chr(34) &" as per expectation")
			End If
			
			If NOT Instr(1,cellValue2,"Renewal")>0 Then
				objCells(intRows + nTop, intOfferDetailCol + nLeft).Interior.ColorIndex = 3	
				Call fRptWriteReport("Fail","Verify "& chr(34) &"Offer Detail" & chr(34) &" column should contain value as "& chr(34) &"Renewal" & chr(34) & " When Switch Direction is "& chr(34) &"Switch To"& chr(34) &" in row no_"& intRows + nTop, "There is a mismatch found as the actual value is " & objCells(intRows + nTop, intOfferDetailCol + nLeft).Value &". It should be "&chr(34)&"Renewal"&chr(34)&" as per expectation")
			End If 

			'Verifying the Salesbook CSLT
			cellValue3=objCells(intRows + nTop, intSalesbookCSLTCol + nLeft).Value
			If NOT (Instr(1,cellValue3,"Renewal")>0) Then
				objCells(intRows + nTop, intSalesbookCSLTCol + nLeft).Interior.ColorIndex = 3
				Call fRptWriteReport("Fail","Verify "&chr(34)&"Salesbook CSLT"& chr(34) &" column should contain value as "& chr(34) &"Renewal"& chr(34) &" When Switch Direction is "& chr(34) &"Switch To"& chr(34) &" in row no_"& intRows + nTop, "There is a mismatch found as the actual value is " & objCells(intRows + nTop, intSalesbookCSLTCol + nLeft).Value &".It should be "&chr(34)&"Renewal"&chr(34)&" as per expectation")									
			End If
			
			'Verifying the Settlement Start Date is higher than the Original Start Date
			CellValueSSDate= objCells(intRows + nTop, intSettlementStartDate + nLeft).Value
			CellValueOSDate = objCells(intRows + nTop, intOriginalOrderDate + nLeft).Value
			If NOT (CDATE(CellValueSSDate) >= CDATE(CellValueOSDate)) Then
				objCells(intRows + nTop, intSettlementStartDate + nLeft).Interior.ColorIndex = 3
				Call fRptWriteReport("Fail","Verify "& chr(34) &"Settlement Start Date"& chr(34) &" is higher than the "& chr(34) &"Original Start Date"& chr(34) &" When Switch Direction is "& chr(34) &"Switch To"& chr(34) &" in row no_"& intRows + nTop, "Did not meet the condition," & chr(34) & "Settlement Start Date"& chr(34) &" value is "& chr(34) & CellValueSSDate & chr(34) &" and "& Chr(34)& "Original Order Date" & chr(34)&" value is "& chr(34) & CellValueOSDate & chr(34))
			End If
		End If		
	Next
	
	objWorkbook.Save
	objWorkbook.Close
	objExcel.Quit
	
End Function


'******************************************************************************************************************************************************************************
'	Function Name						:		fVerifyingFINMARTReportWithExpected
'	Objective							:		Used to Verify Fimart Reports with expected data
'	Input Parameters					:		strExpectedReport, strGeneratedReport  
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fVerifyingFINMARTReportWithExpected(Byval strExpectedReport,Byval strGeneratedReport)
	
	On Error Resume Next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.DisplayAlerts = False
	Set objExpWorkbook= objExcel.Workbooks.Open(strExpectedReport, False, False)
	Set objGenWorkbook= objExcel.Workbooks.Open(strGeneratedReport, False, False)
			
	Set objGeneratedSheet = objGenWorkbook.Worksheets("Report 1")
	Set objExpectedSheet = objExpWorkbook.Worksheets("Report 1")
	
	'Get the number of used rows
  	nUsedRowsGeneratedSheet = objGeneratedSheet.UsedRange.Rows.Count
  	nUsedRowsExpectedSheet = objExpectedSheet.UsedRange.Rows.Count
  	
  	'Get the number of used columns
  	nUsedColsGeneratedSheet = objGeneratedSheet.UsedRange.Columns.Count
  	nUsedColsExpectedSheet = objExpectedSheet.UsedRange.Columns.Count
  	
	'Get the top most row that has data
	nTopGeneratedSheet = objGeneratedSheet.UsedRange.Row
	nTopExpectedSheet = objExpectedSheet.UsedRange.Row
	
	'Get leftmost column that has data
	nLeftGeneratedSheet = objGeneratedSheet.UsedRange.Column
	nLeftExpectedSheet = objExpectedSheet.UsedRange.Column

	intRow = 3
	Set objCellsGenerated = objGeneratedSheet.Cells
	Set objCellsExpected = objExpectedSheet.Cells
	
	For intcolumns =2 to nUsedColsExpectedSheet+1 Step 1
		strColName= objCellsExpected(4,intcolumns).value
		intColumnGenerated= fGetColumninGenerated(objGeneratedSheet,strColName)
	    For introws = 5 To nUsedRowsGeneratedSheet Step 1
			If NOT (Trim(objCellsExpected(introws,intcolumns).value) = Trim(objCellsGenerated(introws, intColumnGenerated).value)) Then
				If IsEmpty(objCellsExpected(introws,intcolumns).value) AND NOT ISEmpty(objCellsGenerated(introws, intColumnGenerated).value) Then
					objCellsGenerated(introws, intColumnGenerated).Interior.ColorIndex = 3	
					Call fRptWriteReport("Fail","Compare "& chr(34) & strColName & chr(34) &" value between FINMART Output sheet and Expected sheet in row no "& introws ,chr(34) & objCellsGenerated(introws, intColumnGenerated).value & chr(34) &" is displayed as actual value but it should be blank as per expectation")				
				Else							
					objCellsGenerated(introws, intColumnGenerated).Interior.ColorIndex = 3	
					Call fRptWriteReport("Fail","Compare "& chr(34) & strColName & chr(34) &" value between FINMART Output sheet and Expected sheet in row no "& introws , "There is a mismatch found between expected value " & chr(34) & objCellsExpected(introws,intcolumns).value & chr(34) &" and Acutal Value " & chr(34) & objCellsGenerated(introws, intColumnGenerated).value & chr(34) )				
				End If
			End If
	    Next
	Next
	
	'Save and close the excel
	objGenWorkbook.Save
	objGenWorkbook.Close
	objExpWorkbook.Close
	objExcel.Quit
		
End Function	


'******************************************************************************************************************************************************************************
'	Function Name						:		fGetColumninGenerated
'	Objective							:		Used to get column names
'	Input Parameters					:		strFolderPath, strFolderName 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fGetColumninGenerated(Byval objSheet,Byval strName)
	
	On Error Resume Next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If		
	nUsedCols = objSheet.UsedRange.Columns.Count
	Set objCells = objSheet.Cells	
	For inti = 2 To nUsedCols+1 Step 1
		If Trim(Lcase(objCells(4,inti).value)) = Trim(Lcase(strName)) Then
			fGetColumninGenerated=inti
			Exit For
		End If
	Next
	
End Function


'******************************************************************************************************************************************************************************
'	Function Name						:		fCreateFolder
'	Objective							:		Create a Folder in Specified path
'	Input Parameters					:		strFolderPath, strFolderName 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fCreateFolder(Byval strFolderPath,ByVal strFolderName)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim FolderName
	Dim fso
	strFolderName=strFolderPath+"\"+strFolderName
    Set fso = createobject("Scripting.FilesystemObject")
    If Not fso.FolderExists(strFolderName) Then
        fso.CreateFolder (strFolderName)
    End If
    
  End Function
   
   
'******************************************************************************************************************************************************************************
'	Function Name						:		fDeleteFolder
'	Objective							:		Delete a Folder in Specified path
'	Input Parameters					:		strFolderPath, strDeleteFolderName 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fDeleteFolder(Byval strFolderPath,ByVal strDeleteFolderName)   
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim strFolderName
	Dim fso
	strFolderName=strFolderPath+"\"+strDeleteFolderName
   	Set fso = CreateObject("Scripting.FileSystemObject")
   	If fso.FolderExists(strFolderName) Then
	 	fso.DeleteFolder(strFolderName)
	End  IF
	Set fso = Nothing
	
   End Function
   

'******************************************************************************************************************************************************************************
'	Function Name						:		fReadingDataInTextFile
'	Objective							:		Read Data in text file
'	Input Parameters					:		strFolderPath, strTXTFileName 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fReadingDataInTextFile(Byval strFolderPath,Byval strTXTFileName)
   	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If			
	Const ForReading = 1
	Dim objFso, objFile, FileName, TextLine
	Dim txtFileName
	txtFileName=strFolderPath+"\"+strTXTFileName+".txt"			
	FileName = txtFileName'Provide your file path
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.OpenTextFile(FileName, ForReading)
	'Read from the file
	Do While objFile.AtEndOfStream <> True
	    TextLine = objFile.ReadLine
	Loop
	objFile.Close
	Set objFile = Nothing
	Set objFso = Nothing
	
   End Function


'******************************************************************************************************************************************************************************
'	Function Name						:		fReadingAllDataInTextFile
'	Objective							:		Read all data in text file
'	Input Parameters					:		strFolderPath, strTXTFileName 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fReadingAllDataInTextFile(Byval strFolderPath,Byval strTXTFileName)
   'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Const ForReading = 1
	Dim objFso, objFile, FileName, strText
	Dim txtFileName
	txtFileName=strFolderPath+"\"+strTXTFileName+".txt"
	
	FileName = txtFileName
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.OpenTextFile(FileName, ForReading)
	fReadingAllDataInTextFile = objFile.ReadAll
	objFile.Close
	Set objFile = Nothing
	Set objFso = Nothing
	
End Function
   
 
'******************************************************************************************************************************************************************************
'	Function Name						:		fCreateTextFile
'	Objective							:		Create text file
'	Input Parameters					:		strFolderPath, strTXTFileName 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fCreateTextFile(Byval strFolderPath,Byval strTXTFileName)   
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If		
	Dim objFSO
	Dim objFolder
	Dim objFile
	
	strTXTFileName=strTXTFileName+".txt"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFolderPath +"\"+strTXTFileName) Then
		Set objFolder = objFSO.GetFolder(strFolderPath)
	Else
		Set objFile = objFSO.CreateTextFile(strFolderPath +"\"+strTXTFileName)
	End if		
	Set objFSO = Nothing
	Set objFolder = Nothing
	Set objFile = Nothing
		
End Function


'******************************************************************************************************************************************************************************
'	Function Name						:		fDeleteTextFile
'	Objective							:		Delete text file
'	Input Parameters					:		strFolderPath, strTXTFileName 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fDeleteTextFile(Byval strFolderPath,Byval strTXTFileName)   		
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objFSO
	Dim strTxtFilepath
	
	strTXTFileName=strTXTFileName+".txt"
	Set objFSO=createobject("Scripting.filesystemobject")
	
	If objFSO.FileExists(strFolderPath +"\"+strTXTFileName) Then
		Set strTxtFilepath = objFSO.GetFile(strFolderPath +"\"+strTXTFileName)
		strTxtFilepath.Delete()
	End If			
	Set strTxtFilepath = Nothing 
	Set objFSO = Nothing
	
End Function


'******************************************************************************************************************************************************************************
'	Function Name						:		fDBGetFieldValue
'	Objective							:		Used to get retrieve from Database
'	Input Parameters					:		strQuery,strCol 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fDBGetFieldValue(strQuery, strCol)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Const adUseClient = 3
	Const adOpenStatic = 3
	Const adLockOptimistic = 3
	Set UdfDBConnect1 = CreateObject("ADODB.Connection") 
	UdfDBConnect1.CommandTimeout = 800
	strDBConnString = "Driver= {Microsoft ODBC for Oracle}; " &_
                       "ConnectString=(DESCRIPTION=" &_
                       "(ADDRESS=(PROTOCOL=TCP)" &_
                       "(HOST=racd1901e.dc.ricohonline.org) (PORT=1510))" &_
 					   "(CONNECT_DATA=(SERVICE_NAME=UATE)));uid=BVEERAVALLI; pwd=BVEERAVALLI_11122018;"    	
	UdfDBConnect1.Open strDBConnString
	If Err.Number = 0 Then
		blnRtnVal = True	
	Else
		blnRtnVal = False
	End If
	
	If UdfDBConnect1.State="1" Then
		Reporter.ReportEvent micPass,"Database Connection" ,"successfully Connected to UATE Database."
	Else
		Reporter.ReportEvent micFail,"Database Connection" ,"Connection to UATE Database failed."
		Environment("ERRORFLAG") = False
		Exit Function
	End If
	
	Set objRecSet = CreateObject("ADODB.Recordset")
	objRecSet.CursorLocation = adUseClient
	objRecSet.Open strQuery, UdfDBConnect1, adOpenStatic, adLockOptimistic
	
	If Err.Number = 0 Then
		blnRtnVal = True
	Else
		blnRtnVal = False
	End If
	
	strDBFieldValue=Trim(objRecSet.Fields.Item(strCol))		
	fDBGetFieldValue=strDBFieldValue
	Set objRecSet = Nothing
	
	If Err.Number <> 0 	Then 
		blnRetStatus = False		
	Else
		Reporter.ReportEvent micPass,"Verify Returned record value","Verify record retrieved:'"&strDBFieldValue&"' as expected"
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Returned record value" , "Record is retrieved sucessfully :'"&strDBFieldValue)
	End If
	
End Function


'******************************************************************************************************************************************************************************
'	Function Name						:		fSetupFolderStructure
'	Objective							:		Used to setup the folder structure and download resources from ALM
'	Input Parameters					:		strBrName -  Browser Name 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
'Public Function fSetupFolderStructure()
'	'Verify if Step Failed, If yes, it will not run the function
'    If Environment("StepFailed") = "YES" Then
'		Exit Function
'	End If	
'	Call fFrameworkFolderConfiguration()
'	If fEnvironmentVarExists("strTestDataPath") Then
'		strFileName = Environment("TestName") & "_TD.xls"
'		gstrFile = Environment.Value("strTestDataPath")
'	Else
'		strFileName = Environment("TestName") & "_TD.xls"
'		gstrFile = gstrTestData & Environment("TestName") & "_TD.xls"
'	End If	
'	gstrOutputTestDataFile = gstrTestData & strFileName
'	Call ReadExternalConfigFile(gstrGlobalConfigPath)
'	'Msgbox "Input Folder - " & gstrFile
'	'Msgbox "Output Folder - " &gstrOutputTestDataFile
'	
'	'Old Code
''	'Call fLoadEnvironment()
''	Call fDownloadResourcesFromALM 
''	strFileName = Environment("TestName") & "_TD.xls"			
''	gstrFile = gstrProjectTestdataPath & Environment("TestName") & "_TD.xls"
''    Datatable.GetSheet("Global").AddParameter "StepStatus",""
'    
'End Function

Public Function fSetupFolderStructure()

	Call fFrameworkFolderConfiguration()
	'Load Object Repositories in runtime
	
	If RepositoriesCollection.Find(gstrAcceleratorsSAPAribaORPath&"\SAPAriba.tsr") = -1 Then
		RepositoriesCollection.Add(gstrAcceleratorsSAPAribaORPath&"\SAPAriba.tsr")
	End If
	
	If RepositoriesCollection.Find(gstrAcceleratorsSAPFioriORPath&"\SAPFiori.tsr") = -1 Then
		RepositoriesCollection.Add(gstrAcceleratorsSAPFioriORPath&"\SAPFiori.tsr")
	End If
	
'	If RepositoriesCollection.Find(gstrAcceleratorsSAPConcurORPath&"\SAPConcur.tsr") = -1 Then
'		RepositoriesCollection.Add(gstrAcceleratorsSAPConcurORPath&"\SAPConcur.tsr")
'	End If
	
'	If fEnvironmentVarExists("strTestDataPath") Then
'		strFileName = Environment("TestName") & "_TD.xls"
'		gstrFile = Environment.Value("strTestDataPath")
'	Else
'		strFileName = Environment("TestName") & "_TD.xls"
'		gstrFile = gstrTestData & Environment("TestName") & "_TD.xls"
'	End If	

'	gstrOutputTestDataFile = gstrTestData & strFileName

	strFileName = Environment("TestName") & "_TD.xls"           
    gstrFile = gstrProjectTestdataPath & Environment("TestName") & "_TD.xls"
	
	'Reading Config Files
	Call fReadGlobalConfigFile()
	Call fReadFrameworkConfig()
		    
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Sub Name						    :					fLoadEnvironment
'	Objective							:					Used to load environment 
'	Input Parameter 					: 					Nil
'	Date Created						:					
'	UFT Version							:					UFT 15.0										
'	Pre-requisites						:					Nil
'	Created By							:					Cigniti
'	Modification Date					:
'*****************************************************************************************************************************************************************************************************************************************
Sub fLoadEnvironment
	'Load environment file in runtime
	Environment.LoadFromFile gstrProjectConfigFilePath
     
End Sub


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fLoadFunctionLibraries
'	Objective							:					Used to load Libraries in runtime
'	Input Parameter 					: 					
'	Date Created						:					
'	UFT Version							:					UFT 15.0					
'	Pre-requisites						:					Nil
'	Created By							:					Cigniti
'	Modification Date					:		   
'*****************************************************************************************************************************************************************************************************************************************
Public Function fLoadFunctionLibraries()
	
	'Load functional libraries in runtime
	LoadFunctionLibrary gGlobalLibraryFilePath    
	LoadFunctionLibrary gstrCommonLibraryFilePath     
	LoadFunctionLibrary gstrBusinessLibraryFilePath     	
	LoadFunctionLibrary gstrQCutilLibraryFilePath    
	LoadFunctionLibrary gRegisterLibraryFilePath     
	LoadFunctionLibrary gstrReportsLibraryFilePath
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fEnterText_SetSecureMode
'	Objective							:					Used to Enter a Text/Value into Edit Box in SecureMode
'	Input Parameter 					: 					Object Name,strValue
'	Date Created						:					
'	UFT Version							:					UFT 15.0					
'	Pre-requisites						:					Nil
'	Created By							:					Cigniti
'	Modification Date					:		   
'*****************************************************************************************************************************************************************************************************************************************
Public Function fEnterText_SetSecureMode(ByRef sObjectName,ByRef strValue, ByRef strText)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim RefObject 
	Environment.Value("StepName") = "Enter a value"
	fEnterText_SetSecureMode=False
    Set RefObject = sObjectName    
    RefObject.RefreshObject
    If RefObject.Exist(MIN_WAIT) Then
        If RefObject.GetRoProperty("enabled") = True OR RefObject.GetRoProperty("disabled") = 0 Then
            RefObject.highlight
            Select Case RefObject.GetRoProperty("micclass")    
                Case "WebEdit"
                    RefObject.SetSecure strValue
                    fEnterText_SetSecureMode = True
                    If Instr(1,RefObject.GetRoProperty("name"),"password",1) > 0 or instr(1,RefObject.GetRoProperty("type"),"password",1) > 0 Then                            
                       Call fRptWriteReport("Pass","Enter value "&Chr(34)&"********"&chr(34)&" in '"&strText&"' field" ,"Value successfully entered in '"&strText&"' field")
					Else
						Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strText&"' field" ,"Value successfully entered in '"&strText&"' field")						     
                    End  IF
                Case "WebCheckBox"
                    RefObject.SetSecure strValue
                    fEnterText_SetSecureMode = True
                     Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strText&"' CheckBox" ,"Value successfully entered in '"&strText&"' CheckBox")						     

                Case "JavaEdit"
                    RefObject.SetSecure strValue
                    fEnterText_SetSecureMode = True
                    Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strText&"' field" ,"Value successfully entered in '"&strText&"' field")						     

                Case "OracleTextField"
                    RefObject.Enter strValue
                    fEnterText_SetSecureMode = True'                    
                    If instr(1,RefObject.GetROProperty("name"),"password",1) > 0 or instr(1,RefObject.GetROProperty("type"),"password",1) > 0 Then
                        Call fRptWriteReport("Pass","Enter value "&Chr(34)&"********"&chr(34)&" in '"&strText&"' field" ,"Value successfully entered in '"&strText&"' field")
					Else
						Call fRptWriteReport("Pass","Enter value "&Chr(34)&strValue&chr(34)&" in '"&strText&"' field" ,"Value successfully entered in '"&strText&"' field")						     
                    End If                                        
                Case Else
                    RefObject.SetSecure strValue
                    fEnterText_SetSecureMode = True
                    Call fRptWriteReport("Pass","Enter value "&Chr(34)&"********"&chr(34)&" in '"&strText&"' field" ,"Value successfully entered in '"&strText&"' field")
             End Select
        Else 
            Call fRptWriteReport("Fail","Enter value in '"&strFieldName&"'" ,"Value is not entered in '"&strText&"' field")
            Environment("ERRORFLAG")=False
            fEnterText_SetSecureMode = False            
            Exit Function
        End If
    Else
        Call fRptWriteReport("Fail","Enter value in '"&strFieldName&"'" ,"Value is not entered in '"&strText&"' field")
        Environment("ERRORFLAG")=False
        fEnterText_SetSecureMode = False        
        Exit Function
    End If
	On Error GOTO 0

End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fEnterTextInCell
'	Objective							:					Used to Enter a Text/Value into Edit Box in AUT
'	Input Parameter 					: 					sObjectName,strRow,strCol,strMicClas,intIndex,strText
'	Date Created						:					
'	UFT Version							:					UFT 15.0					
'	Pre-requisites						:					Nil
'	Created By							:					Cigniti
'	Modification Date					:		   
'*****************************************************************************************************************************************************************************************************************************************
Function fEnterTextInCell(sObjectName,strRow,strCol,strMicClas,intIndex,strText)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	On Error Resume Next
	If Len(Trim(strText))<>0 Then
		If Not IsObject(sObjectName) Then
			Set RefObject=Eval(fnGetObjectHierarchy(sObjectName)) 
		Else
			Set RefObject = sObjectName    
		End If
		RefObject.RefreshObject
		If RefObject.Exist(MID_WAIT) Then
			If RefObject.GetRoProperty("enabled") = True or GetRoProperty("visible") = True Then
				Select Case RefObject.GetRoProperty("micclass")
				Case "OracleTable"                                                                                                          
					RefObject.EnterField Cint(strRow),strCol, Cstr(strText)
					Call fRptWriteReport("Pass","Enter value "&strText,"Value successfully entered in table")
				Case "WebTable"                                                                                                             
					RefObject.childitem(strRow,strCol,strMicClas,intIndex).set strText
					Call fRptWriteReport("Pass","Enter value "&strText,"Value successfully entered in table")
				End Select   
				'Return boolean Value True to the Called block
				fEnterTextInCell= True
			Else
				Call fRptWriteReport("Fail","Enter value "&strText,"Value is not entered in table")
				Exit Function
			End If
		Else
			Call fRptWriteReport("Fail","Enter value "&strText,"Value is not entered in table")
			Exit Function
		End If
	Else	
		fEnterTextInCell=False
	End If
	On Error GOTO 0
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fnGetRoProperty
'	Objective							:					Used to retrive the property value of the given object
'	Input Parameter 					: 					sObjectName,sPropertyName
'	Date Created						:					
'	UFT Version							:					UFT 15.0					
'	Pre-requisites						:					Nil
'	Created By							:					Cigniti
'	Modification Date					:		   
'*****************************************************************************************************************************************************************************************************************************************
Function fGetRoProperty(sObjectName,sPropertyName,strText)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim RefObject 
	Dim sValue
	Environment.Value("StepName") = "Retrieve Object Value during runtime"
	Set RefObject=sObjectName
	
	'Return boolean Value Flase to the Called block
	fGetRoProperty=False
	If RefObject.Exist(MIN_WAIT) Then
		fGetRoProperty=RefObject.GetRoProperty(sPropertyName)
		sValue = RefObject.GetRoProperty(sPropertyName) '' Retrvied for report PASSWITHSCREENSHOT
		'Call fRptWriteReport("PASSWITHSCREENSHOT", "Capture "&strText&" value","value has been captured "&fGetRoProperty)
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Capture '"&strText&"' value","The property value captured as '"&fGetRoProperty&"'")
	else
		'Call fRptWriteReport("Fail", "Capture '"&strText&"' value","value has not been captured")
		Call fRptWriteReport("Fail", "Capture '"&strText&"' value","The property value not captured")
		Environment("ERRORFLAG") = False
		fGetRoProperty = False
	End If
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fVerifyObjectExist
'	Objective							:					Used to Verify the Object Exist in the AUT
'	Input Parameter 					: 					Object Name,InputValue
'	Date Created						:					
'	UFT Version							:					UFT 15.0					
'	Pre-requisites						:					Nil
'	Created By							:					Cigniti
'	Modification Date					:		   
'*****************************************************************************************************************************************************************************************************************************************
Function fVerifyObjectExist(sObjectName)
   	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
   	'Initially Assigning block to False
	fVerifyObjectExist=False
	If Not IsObject(sObjectName) Then
		Set RefObject=Eval(fnGetObjectHierarchy(sObjectName))
	Else
		Set RefObject = sObjectName		 
	End If
	
	If  RefObject.Exist(1)Then
        'Return boolean Value True to the Called block
		fVerifyObjectExist = True
	Else
		fVerifyObjectExist = False
	End If
	Set RefObject=Nothing
	
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name						:					fGetCelldata
'	Objective							:					Used to Enter a Text/Value into Edit Box in AUT
'	Input Parameter 					: 					Object Name,InputValue
'	Date Created						:					
'	UFT Version							:					UFT 15.0					
'	Pre-requisites						:					Nil
'	Created By							:					Cigniti
'	Modification Date					:		   
'*****************************************************************************************************************************************************************************************************************************************
Public Function fGetCelldata(sObjectName,iRow,iCol,strText)

	On Error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim RefObject 
 	Environment.Value("StepName") = "Retrieve Value In a Table "
	Set RefObject = sObjectName			 	
    
    'Initially Assigning block to False
	fGetCelldata=False
    RefObject.RefreshObject
	If RefObject.Exist(MIN_WAIT) Then
		If RefObject.GetRoProperty("enabled") = True OR RefObject.GetRoProperty("visible") = True Then
			Select Case RefObject.GetRoProperty("micclass")			
				Case "WebTable"
					fGetCelldata = RefObject.GetCellData(Cint(iRow),Cint(iCol))
					'Call fRptWriteReport("PASSWITHSCREENSHOT",Environment.Value("StepName"),chr(34) & "'"&strText&"'" & chr(34) & " retrieved value is : " & chr(34) & strText& chr(34))															
				Case "JavaTable"
					fGetCelldata = RefObject.GetCellData(Cint(iRow),Cint(iCol))
					'Call fRptWriteReport("PASSWITHSCREENSHOT", sObjectName.ToString,sObjectName.ToString&" retrieved value is "&"'"&fnGetCelldata&"'")
				Case "SAPUITable"
					fGetCelldata = RefObject.GetCellData(Cint(iRow),Cint(iCol))
					'Call fRptWriteReport("PASSWITHSCREENSHOT", sObjectName.ToString,sObjectName.ToString&" retrieved value is "&"'"&fnGetCelldata&"'")
				Case "OracleTable"
					fGetCelldata = RefObject.GetFieldValue(Cint(iRow),Cint(iCol))
					'Call fRptWriteReport("PASSWITHSCREENSHOT", sObjectName.ToString,sObjectName.ToString&" retrieved value is "&"'"&fnGetCelldata&"'")
			End select
		Else
		    Call  fRptWriteReport("Fail","Retrieve Value In a Table ","Not Capture table cell data")
		    Environment("ERRORFLAG") = False
		    fGetCelldata=False
		    Exit Function
		End If
	Else
	    Call  fRptWriteReport("Fail","Retrieve Value In a Table ","Not Capture table cell data")
	    Environment("ERRORFLAG") = False
	    fGetCelldata=False
	    Exit Function
	End If
		
	On Error GOTO 0	
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'	Function Name					:				  fSelectRowInTable
'	Objective						:				  Used to Select row in table
'	Input Parameters				:				  sObjectName,strRow
'	Output Parameters			    :				  Nil
'	Date Created					:				  20/04/2015
'	UFT Version						:				  UFT 15.0
'	Pre-requisites					:				  NIL  
'	Created By						:				  Cigniti
'	Modification Date		        :		   		  20/04/2015
'******************************************************************************************************************************************************************************************************************************************
Public Function fSelectRowInTable(sObjectName,intRow,strObjectName)
	
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Set RefObject = sObjectName				 
	If RefObject.Exist(MAX_WAIT) Then
		If RefObject.GetRoProperty("disabled") = False Then
			Select Case RefObject.GetRoProperty("micclass")
				Case "SAPUITable"
					RefObject.SelectRow Cint(intRow)	
					fSelectRowInTable= True
					Call fRptWriteReport("Pass","Select row in table "&strObjectName&" "&intRow&" row is selected")
			End Select 
		Else
			Call fRptWriteReport("Fail","Select row in table "&strObjectName&" "&intRow&" row is not selected")
			fSelectRowInTable= False
			Exit Function
		End If
	Else
		Call fRptWriteReport("Fail","Select row in table "&strObjectName&" "&intRow&" row is not selected")
		fSelectRowInTable= False
		Exit Function
	End If
	On error goto 0
	
End Function


'******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fnGetCalenderDate
'	Objective							:					Used to Get the Calender Date (eg:'Fri, 1 May, 2020 )
'	Input Parameters					:					nil
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:									
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Function fGetCalenderDate()

	currentDate = Date
    dDate = Day(currentDate)	    
    strMonthName=MonthName(Month(currentDate))	
    dMonth= left(strMonthName,3) 
    gfGetCurrentCalendarMonthName = dMonth		
    dYear = Year(currentDate)
    strCurDay = WeekDayName(Weekday(currentDate),True)
    strCurDate = strCurDay&","&"  "&dDate&" "&gfGetCurrentCalendarMonthName&","&"  "&dYear
    fGetCalenderDate = strCurDate
    
 End Function
 

'******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fGetText
'	Objective							:					Retrieving the innertext all Objects(Verification)
'	Input Parameters					:					nil
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:									
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fGetText(objElement,sPropertyName,sMessage)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	wait (MIN_WAIT)
	Dim sFlag
	sFlag = True
	sObjName=objElement.GetToProperty("TestobjName")
	If  objElement.Exist(MIN_WAIT) Then
		text=objElement.GetROProperty(sPropertyName) 
		Call fRptWriteReport("Pass",sMessage,"Value is Retrieved and the Value is " &text)
	Else
	    Call fRptWriteReport("Fail",sMessage,"Value is not Retrived and the Value " &text)
	    sFlag = False                                   
	End If
	fGetText = text
	If sFlag = False Then
	    Call fErrorChecking()
	End If
End Function


'******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fGetText
'	Objective							:					Retrieving the innertext all Objects(Verification)
'	Input Parameters					:					nil
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:									
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Function fRemoveSpecialCharectersFromString(strPattern,strString,strReplaceChar)
Dim objRe
Set objRe = New RegExp
objRe.Global = True
objRe.Pattern = strPattern
fRemoveSpecialCharectersFromString = objRe.Replace(strString,strReplaceChar)
End Function


'******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fGetText
'	Objective							:					Retrieving the innertext all Objects(Verification)
'	Input Parameters					:					nil
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:									
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function ReadExternalConfigFile(strConfigFilePath)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim ObjXML   
	Dim objRoot  
	Dim objNodeList  
	Dim strName  
	Dim strValue
	Dim strRerunFailuresFlag
	Dim strRerunIterations
	
	Set ObjXML = CreateObject ("Microsoft.XMLDOM")   
	ObjXML.async = False     
	ObjXML.load(strConfigFilePath)  
	Set objRoot = ObjXML.documentElement   
	Set objNodeList = objRoot.getElementsByTagName("Variable")   
	For Each Elem In objNodeList   
		Set strName = Elem.getElementsByTagName ("Name")(0)  
		Set strValue = Elem.getElementsByTagName ("Value")(0)
		
		If strName.Text = "Environment" Then
			gstrTestEnvironment = strValue.Text
		End If
		If strName.Text = "ChromeBrowser" Then
			gstrChromeBrowser = strValue.Text
		End If
		If strName.Text = "IEBrowser" Then
			gstrIEBrowser = strValue.Text
		End If
		If strName.Text = "AribaApplicationURL" Then
			gstrAribaApplicationURL = strValue.Text
		End If
		If strName.Text = "AribaApplicationUsername" Then
			gstrAribaApplicationUsername = strValue.Text
		End If
		If strName.Text = "AribaApplicationPassword" Then
			gstrAribaApplicationPassword = strValue.Text
		End If
		If strName.Text = "AribaSupplierApplicationURL" Then
			gstrAribaSupplierApplicationURL = strValue.Text
		End If
		If strName.Text = "AribaSupplierUsername" Then
			gstrAribaSupplierUsername = strValue.Text
		End If
		If strName.Text = "AribaSupplierPassword" Then
			gstrAribaSupplierPassword = strValue.Text
		End If
	Next
	
End Function

''******************************************************************************************************************************************************************************************************************************************
''	Function Name		 				:					fFillAdvancePaymentTemplatePDF
''	Objective							:					Used to fill Advance Payment Template PDF File
''	Input Parameters					:					strPDFFileName, strFieldValues -  String - Update the values in below order by | (pipeline) separated
''															Autodesk Entity Name, 
''															Name of Supplier, 
''															Supplier SAP ID, 
''															PO Number,
''															Payment Request Amount, 
''															Invoice Number, 
''															Business Reason for Advance Payment Request, 
''															Account Holder Name,
''															Bank Name,
''															Bank Account Number, 
''															Bank & Branch Code,
''															Invoice Date,
''															Payment DUE Date,
''															Special Instructions to be stipulated on the Paym, 
''															Company Code,
''															Payment Currency
''	Output Parameters					:					NIL
''	Date Created						:					
''	QTP Version							:					UFT 15.0
''	Pre-requisites						:					NILL  
''	Created By							:					Cigniti Techologies
''	Modification Date					:		   
''******************************************************************************************************************************************************************************************************************************************		
''Call fFillAdvancePaymentTemplatePDF(objDataDict, iRowCountRef, "CreditMemoInvoice")
'
''Call fFillAdvancePaymentTemplatePDF(objDataDict, iRowCountRef, "NormalInvoice")
'
''Call fFillAdvancePaymentTemplatePDF(objDataDict, iRowCountRef, "DownPaymentInvoice")  ' RPA-DWN-Y30
'
'Public Function fFillAdvancePaymentTemplatePDF(objDataDict, iRowCountRef, strInvoiceType)
'    intCompanyCode 	= objDataDict.Item("CompanyCode" & iRowCountRef)
'    intCompanyCode 	= Left(intCompanyCode,4)
'	intVendorName	= objDataDict.Item("VendorName" & iRowCountRef)
'	intPrice	 	= objDataDict.Item("Price" & iRowCountRef)
'	intQuntity	 	= objDataDict.Item("Quantity" & iRowCountRef)
'	strCurrencyName	= objDataDict.Item("CurrencyName" & iRowCountRef)
'	strVendorFullName = objDataDict.Item("VendorFullName" & iRowCountRef)
'	strAutodeskEntityName = objDataDict.Item("EntityName" & iRowCountRef)
'	
'	intPOnumber 	= fGetSingleValue("AutoPONumber","TestData",Environment("TestName"))
'	'intPrice		= CInt(intPrice) * CInt(intReEnterQuantity)
'	intPrice        = CInt(intPrice) * CInt(intQuntity)
'    intUpdatedPrice = fGetSingleValue("UpdatedPrice","TestData",Environment("TestName"))
'		If Len(intVendorName) >= 10 Then
'			intVendorName 	= Right(Left(intVendorName,10),6)
'		ElseIf Len(intVendorName) = 6 Then
'			intVendorName = intVendorName
'		Else
'			Call fRptWriteReport("Fail","Enter Supplier SAP ID","Supplier SAP ID is Not Valid")
'			ExitAction
'		End If
'		Call fGetCurrentDateFormatMMDDYYYY()
'		dtInvoiceDate 	= Environment.Value("dtCurrentDate") ' OR fGetCurrentDateFormat
'		dtDueDate		= Environment.Value("dtFutureDate")
'		strPDFFileName 	= "Source_AdvancePaymentTemplate"  'C:\CTAF_28\Projects\PTP\TestExecutionPlan\Files\PDFFiles
'	    'Creating Dotnet factory object to access methods in PDFLibrary Methods
'	    Set ObjDotNetFactory = DotNetFactory.CreateInstance("PDFLibrary.Utilities", gstrFrameworkUtilityLibrariesPath&"\PDFLibrary.dll")
'	    If strInvoiceType = "NormalInvoice" Then
'	        intInvoiceNumber= "IN" & intPOnumber
'			Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intInvoiceNumber,"TestData","AutoInvoiceNumber")
'	        strInputFieldNames = "txt1|3|txt4|txt6|7|txt9|Business Reason|Date2_af_date|Date3_af_date|text20|text21"
'	        strFieldValues 	= strAutodeskEntityName&"|"&strVendorFullName&"|"&intVendorName&"|"& intPOnumber &"|"& intPrice &"|"& intInvoiceNumber &"|TEST |"& dtInvoiceDate &"|"& dtDueDate &"|"& intCompanyCode &"|"& strCurrencyName &""
'	     ElseIf strInvoiceType = "CreditMemoInvoice" Then  
'	        intInvoiceNumber= "CM" & intPOnumber
'	        intPrice        = CInt(intPrice) - CInt(intUpdatedPrice)
'	        Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intInvoiceNumber,"TestData","AutoInvoiceNumber")
'	        Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intPrice,"TestData","CMPrice")
'	        strInputFieldNames = "txt1|3|txt4|txt6|7|txt9|Business Reason|Date2_af_date|Date3_af_date|text20|text21"
'	        strFieldValues     = "Autodesk Inc|GATES & COOPER, LLP|"&intVendorName&"|"& intPOnumber &"|"& intPrice &"|"& intInvoiceNumber &"|CREDIT MEMO |"& dtInvoiceDate &"|"& dtDueDate &"|"& intCompanyCode &"|"& strCurrencyName &""
'	    ElseIf strInvoiceType = "DownPaymentInvoice" Then  
'	        intInvoiceNumber= "DWN" & intPOnumber
'	        Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intInvoiceNumber,"TestData","AutoInvoiceNumber")
'	        strInputFieldNames = "txt1|3|txt4|txt6|7|txt9|Business Reason|Date2_af_date|Date3_af_date|text20|text21"
'	        strFieldValues     = strAutodeskEntityName&"|"&strVendorFullName&"|"&intVendorName&"|"& intPOnumber &"|"& intPrice &"|"& intInvoiceNumber &"|RPA TEST |"& dtInvoiceDate &"|"& dtDueDate &"|"& intCompanyCode &"|"& strCurrencyName &"" 
'	    End If
'	    'Input and output PDF File paths
'	    strPDFAdvncePymtTmpltePath 	= gstrFrameWorkFolder&"\"&gstrTestPlanName&"\Files\PDFFiles\"
'	    strSourcePDF 				= strPDFAdvncePymtTmpltePath & strPDFFileName&".pdf"
'	    strOutputPDF 				= strPDFAdvncePymtTmpltePath & Environment.Value("TestName")&"_"&strInvoiceType&".pdf"
'	    strLogFilePath 				= strPDFAdvncePymtTmpltePath & "PDFFillLogFile.txt"
'	    'Calling a method to fill PDF Files with input data
'	    strFlag = ObjDotNetFactory.PDFFillForm (strSourcePDF,strOutputPDF,strInputFieldNames, strFieldValues,strLogFilePath)
'	    If strFlag Then
'	        Call fRptWriteReport("Pass", "Fill Advance Payment Template PDF File","Advance Payment Template PDF File is filled sucessfully, Output File Name :"&strOutputPDF)
'	    Else   
'	        Call fRptWriteReport("Fail", "Fill Advance Payment Template PDF File","Failed to fill Advance Payment Template PDF File, Input File Name :"&strSourcePDF)          
'	    End If
'End Function	
''******************************************************************************************************************************************************************************************************************************************
''	Function Name		 				:					fSendEmailInvoice
''	Objective							:					Used to send email to autodesk with PDF invoice type
''	Input Parameters					:					strRecipientsEMail, strEmailSubject, strEmailBody, strInvoiceType
''	Output Parameters					:					NIL
''	Date Created						:					05/28/2020
''	QTP Version							:					UFT 15.0
''	Pre-requisites						:					NILL  
''	Created By							:					Cigniti Techologies
''	Modification Date					:		   			05/29/2020
''******************************************************************************************************************************************************************************************************************************************		
'''************ It's Email Send to Autodesk for NormalInvoice **************
''Call fSendEmailInvoice("NormalInvoice")
''Call fSendEmailInvoice("CreditMemoInvoice")
''Wait 720 ' After email send, we should wait minimum 10 min. It is as per manual TC step mentioned in TestRail
'
'Public Function fSendEmailInvoice(objDataDict, iRowCountRef, strInvoiceType)
'	Dim objAccount
'	Dim objMail
'	Dim objFSO
'	Dim intCompanyCode	
'	intCompanyCode 	= objDataDict.Item("CompanyCode" & iRowCountRef)
'    intCompanyCode 	= Left(intCompanyCode,4)
'    'strRecipientsEMail setup based on Company Code
'    Select Case intCompanyCode
'    	Case "1000"
'    	strRecipientsEMail = "test.AP.invoice.singapore@autodesk.com"
'		Case "1001"	
'		strRecipientsEMail = "test.AP.invoice.malaysia@autodesk.com"
'		Case "1002"
'		strRecipientsEMail = "test.AP.invoice.thailand@autodesk.com"
'		Case "1003"	
'		strRecipientsEMail = "test.AP.invoice.indonesia@autodesk.com"
'		Case "1004"	
'		strRecipientsEMail = "test.AP.invoice.philippines@autodesk.com"
'		Case "1005"	
'		strRecipientsEMail = "test.AP.invoice.vietnam@autodesk.com"
'		Case "1100"	
'		strRecipientsEMail = "test.AP.invoice.japan@autodesk.com"
'		Case "1200"	
'		strRecipientsEMail = "test.AP.invoice.australia@autodesk.com"
'		Case "1320"	
'		strRecipientsEMail = "test.AP.invoice.china@autodesk.com"
'		Case "1330"	
'		strRecipientsEMail = "test.AP.invoice.china@autodesk.com"
'		Case "1400"	
'		strRecipientsEMail = "test.AP.invoice.hongkong@autodesk.com"
'		Case "1600"	
'		strRecipientsEMail = "test.AP.invoice.taiwan@autodesk.com"
'		Case "1700"	
'		strRecipientsEMail = "test.AP.invoice.india@autodesk.com"
'		Case "1800"	
'		strRecipientsEMail = "test.AP.invoice.korea@autodesk.com"
'		Case "2000"	
'		strRecipientsEMail = "test.AP.invoice.switzerland@autodesk.com"
'		Case "2006"	
'		strRecipientsEMail = "test.AP.invoice.ireland@autodesk.com"
'		Case "2010"	
'		strRecipientsEMail = "test.AP.invoice.austria@autodesk.com"
'		Case "2020"	
'		strRecipientsEMail = "test.AP.invoice.netherlands@autodesk.com"
'		Case "2025"	
'		strRecipientsEMail = "test.AP.invoice.belgium@autodesk.com"
'		Case "2030"	
'		strRecipientsEMail = "test.AP.invoice.czech.republic@autodesk.com"
'		Case "2040"	
'		strRecipientsEMail = "test.AP.invoice.uk@autodesk.com"
'		Case "2041"	
'		strRecipientsEMail = "test.AP.invoice.uk@autodesk.com"
'		Case "2042"	
'		strRecipientsEMail = "test.AP.invoice.uae@autodesk.com"
'		Case "2043"	
'		strRecipientsEMail = "test.AP.invoice.jordan@autodesk.com"
'		Case "2050"	
'		strRecipientsEMail = "test.AP.invoice.france@autodesk.com"
'		Case "2060"	
'		strRecipientsEMail = "test.AP.invoice.germany@autodesk.com"
'		Case "2070"	
'		strRecipientsEMail = "test.AP.invoice.italy@autodesk.com"
'		Case "2080"	
'		strRecipientsEMail = "test.AP.invoice.spain@autodesk.com"
'		Case "2090"	
'		strRecipientsEMail = "test.AP.invoice.sweden@autodesk.com"
'		Case "2100"	
'		strRecipientsEMail = "test.AP.invoice.switzerland@autodesk.com"
'		Case "2104"	
'		strRecipientsEMail = "test.AP.invoice.south.africa@autodesk.com"
'		Case "2106"	
'		strRecipientsEMail = "test.AP.invoice.qatar@autodesk.com"
'		Case "2130"	
'		strRecipientsEMail = "test.AP.invoice.ireland@autodesk.com"
'		Case "2140"	
'		strRecipientsEMail = "test.AP.invoice.switzerland@autodesk.com"
'		Case "2150"	
'		strRecipientsEMail = "test.AP.invoice.netherlands@autodesk.com"
'		Case "2160"	
'		strRecipientsEMail = "test.AP.invoice.hungary@autodesk.com"
'		Case "2170"	
'		strRecipientsEMail = "test.AP.invoice.poland@autodesk.com"
'		Case "2180"	
'		strRecipientsEMail = "test.AP.invoice.turkey@autodesk.com"
'		Case "2190"	
'		strRecipientsEMail = "test.AP.invoice.russia@autodesk.com"
'		Case "2200"	
'		strRecipientsEMail = "test.AP.invoice.uk@autodesk.com"
'		Case "2240"	
'		strRecipientsEMail = "test.AP.invoice.netherlands@autodesk.com"
'		Case "2250"	
'		strRecipientsEMail = "test.AP.invoice.uk@autodesk.com"
'		Case "2911"	
'		strRecipientsEMail = "test.AP.invoice.israel@autodesk.com"
'		Case "2921"	
'		strRecipientsEMail = "test.AP.invoice.romania@autodesk.com"
'		Case "2935"	
'		strRecipientsEMail = "test.AP.invoice.saudi.arabia@autodesk.com"
'		Case "2936"	
'		strRecipientsEMail = "test.AP.invoice.denmark@autodesk.com"
'		Case "3000"	
'		strRecipientsEMail = "test.AP.invoice.usa@autodesk.com"
'		Case "3011"	
'		strRecipientsEMail = "test.AP.invoice.usa@autodesk.com"
'		Case "3500"	
'		strRecipientsEMail = "test.AP.invoice.canada@autodesk.com"
'		Case "3600"	
'		strRecipientsEMail = "test.AP.invoice.mexico@autodesk.com"
'		Case "3610"	
'		strRecipientsEMail = "test.AP.invoice.argentina@autodesk.com"
'		Case "3620"	
'		strRecipientsEMail = "test.AP.invoice.brazil@autodesk.com"
'		Case "3650"	
'		strRecipientsEMail = "test.AP.invoice.colombia@autodesk.com"
'		Case "3660"	
'		strRecipientsEMail = "test.AP.invoice.mexico@autodesk.com"
'    End Select
'    If strInvoiceType = "DownPaymentInvoice" Then
'    	strRecipientsEMail = "test.ap.advpayment@autodesk.com"
'    End If
'	'strEmailSubject and strEmailBody setup based on Invoice Type
'	strEmailSubject = strInvoiceType & " Generation Request"
'	strEmailBody = "Hi Team, "& vbLf &"This is "& strInvoiceType &" Generation Request...."
'	'attachement saved path
'	strPDFAdvncePymtTmpltePath = gstrFrameWorkFolder&"\"&gstrTestPlanName&"\Files\PDFFiles\"
'	strAttachmentPath = strPDFAdvncePymtTmpltePath & Environment.Value("TestName")&"_"&strInvoiceType&".pdf"
'	'Verify if TestPlan exists in TestPlan Folder
'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'	If objFSO.FileExists(strAttachmentPath) = True then
'		'Creating an object for Outlook Applixation
'		Set objOutlook = CreateObject("Outlook.Application")
'		'Iterating through all the emails accounts configured in the system
'		For Each objAccount In objOutlook.Session.Accounts
'			 If objAccount.AccountType = olPop3 Then 
'			 	Set objMail = objOutlook.CreateItem(olMailItem)
'				'Recipients details	 	
'				objMail.Recipients.Add(strRecipientsEMail)
'				'Email Subject
'				objMail.Subject = strEmailSubject
'				'Body Text
'				objMail.Body= strEmailBody
'				'Adding attachement to the email
'				objMail.Attachments.Add(strAttachmentPath)
'				objMail.Recipients.ResolveAll
'			 	'Verify the email id and send an email with respective mail id
'			 	If Instr(objAccount,"autodesk") Then
'			 		Set objMail.SendUsingAccount = objAccount
'			     	objMail.Send
'			     	Exit for
'			 	End If
'			 End If 
'		Next
'	End If
'End Function
''******************************************************************************************************************************************************************************************************************************************
''	Function Name		 				:					fGetCurrentDateFormatMMDDYYYY
''	Objective							:					Used to Get Current Date and future date format as MM/DD/YYYY
''	Input Parameters					:					N/A
''	Output Parameters					:					Environment.Value("dtCurrentDate") OR fGetCurrentDateFormatMMDDYYYY, and Environment.Value("dtFutureDate")
''	Date Created						:					06/01/2020
''	QTP Version							:					UFT 15.0
''	Pre-requisites						:					NILL  
''	Created By							:					Cigniti Techologies
''	Modification Date					:		   			
''******************************************************************************************************************************************************************************************************************************************		
'Public Function fGetCurrentDateFormatMMDDYYYY()
'	dtDay 	= Day(Date)
'	dtMonth = Month(Date)
'	dtYear 	= Year(Date)
'		If Len(dtDay) = 1 Then
'			dtDay = "0"&dtDay
'		End If
'		If Len(dtMonth) = 1 Then
'			dtMonth = "0"&dtMonth
'		End If
'			dtInvoiceDate 	= dtMonth&"/"&dtDay&"/"&dtYear
'			dtDueDate		= dtMonth&"/"&dtDay + 5&"/"&dtYear + 1
'			fGetCurrentDateFormatMMDDYYYY = dtInvoiceDate
'			Environment.Value("dtCurrentDate") = dtInvoiceDate
'			Environment.Value("dtFutureDate") = dtDueDate 'Replace(DateAdd("m",5,dtInvoiceDate),"-","/")
'End Function
'


'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 				:					fFillAdvancePaymentTemplatePDF
'	Objective							:					Used to fill Advance Payment Template PDF File
'	Input Parameters					:					strPDFFileName, strFieldValues -  String - Update the values in below order by | (pipeline) separated
'															Autodesk Entity Name, 
'															Name of Supplier, 
'															Supplier SAP ID, 
'															PO Number,
'															Payment Request Amount, 
'															Invoice Number, 
'															Business Reason for Advance Payment Request, 
'															Account Holder Name,
'															Bank Name,
'															Bank Account Number, 
'															Bank & Branch Code,
'															Invoice Date,
'															Payment DUE Date,
'															Special Instructions to be stipulated on the Paym, 
'															Company Code,
'															Payment Currency
'	Output Parameters					:					NIL
'	Date Created						:					
'	QTP Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
'Call fFillAdvancePaymentTemplatePDF(objDataDict, iRowCountRef, "CreditMemoInvoice")

'Call fFillAdvancePaymentTemplatePDF(objDataDict, iRowCountRef, "NormalInvoice")

'Call fFillAdvancePaymentTemplatePDF(objDataDict, iRowCountRef, "DownPaymentInvoice")  ' RPA-DWN-Y30

Public Function fFillAdvancePaymentTemplatePDF(objDataDict, iRowCountRef, strInvoiceType)
On Error Resume Next
	Dim intCompanyCode,intVendorName,intPrice,intQuntity,strCurrencyName,intPOnumber,intInvoiceNumber
	intCompanyCode 	= Trim(objDataDict.Item("CompanyCode" & iRowCountRef))
    intCompanyCode 	= Left(intCompanyCode,4)
	intVendorName	= Trim(objDataDict.Item("VendorName" & iRowCountRef))
	intPrice	 	= Trim(objDataDict.Item("Price" & iRowCountRef))
	intQuntity	 	= Trim(objDataDict.Item("Quantity" & iRowCountRef))
	strCurrencyName	= Trim(objDataDict.Item("CurrencyName" & iRowCountRef))
	intPOnumber 	= Trim(fGetSingleValue("AutoPONumber","TestData",Environment("TestName")))
	intPrice		= CInt(intPrice) * CInt(intQuntity)
		If Len(intVendorName) >= 10 Then
			intVendorName 	= Right(Left(intVendorName,10),6)
		ElseIf Len(intVendorName) = 6 Then
			intVendorName = intVendorName
		Else
			Call fRptWriteReport("Fail","Enter Supplier SAP ID","Supplier SAP ID is Not Valid")
			ExitAction
		End If
		Call fGetCurrentDateFormatMMDDYYYY()
		dtInvoiceDate 	= Environment.Value("dtCurrentDate") ' OR fGetCurrentDateFormatMMDDYYYY
		dtDueDate		= Environment.Value("dtFutureDate")
		strPDFFileName 	= "Source_AdvancePaymentTemplate"  'C:\CTAF_28\Projects\PTP\TestExecutionPlan\Files\PDFFiles
	    'Creating Dotnet factory object to access methods in PDFLibrary Methods
	    Set ObjDotNetFactory = DotNetFactory.CreateInstance("PDFLibrary.Utilities", gstrFrameworkUtilityLibrariesPath&"\PDFLibrary.dll")
	    If strInvoiceType = "NormalInvoice" Then
	        intInvoiceNumber= "IN" & intPOnumber
			Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intInvoiceNumber,"TestData","AutoInvoiceNumber")
	        strInputFieldNames = "txt1|3|txt4|txt6|7|txt9|Business Reason|Date2_af_date|Date3_af_date|text20|text21"
	        strFieldValues 	= "Autodesk Inc|GATES & COOPER, LLP.|"&intVendorName&"|"& intPOnumber &"|"& intPrice &"|"& intInvoiceNumber &"|TEST |"& dtInvoiceDate &"|"& dtDueDate &"|"& intCompanyCode &"|"& strCurrencyName &""
	    ElseIf strInvoiceType = "CreditMemoInvoice" Then   
	        intInvoiceNumber= "CM" & intPOnumber
			Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intInvoiceNumber,"TestData","AutoInvoiceNumber")
	        strInputFieldNames = "txt1|3|txt4|txt6|7|txt9|Business Reason|Date2_af_date|Date3_af_date|text20|text21"
	        strFieldValues 	= "Autodesk Inc|GATES & COOPER, LLP.|"&intVendorName&"|"& intPOnumber &"|"& intPrice &"|"& intInvoiceNumber &"|CREDIT MEMO |"& dtInvoiceDate &"|"& dtDueDate &"|"& intCompanyCode &"|"& strCurrencyName &""
	    ElseIf strInvoiceType = "DownPaymentInvoice" Then  
	        intInvoiceNumber= "DWN" & intPOnumber
	        Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intInvoiceNumber,"TestData","AutoInvoiceNumber")
	        strInputFieldNames = "txt1|3|txt4|txt6|7|txt9|Business Reason|Date2_af_date|Date3_af_date|text20|text21"
	        strFieldValues     = "Autodesk Inc|GATES & COOPER, LLP.|"&intVendorName&"|"& intPOnumber &"|"& intPrice &"|"& intInvoiceNumber &"|RPA TEST |"& dtInvoiceDate &"|"& dtDueDate &"|"& intCompanyCode &"|"& strCurrencyName &"" 
	    End If
	    'Input and output PDF File paths
	    strPDFAdvncePymtTmpltePath 	= gstrFrameWorkFolder&"\"&gstrTestPlanName&"\Files\PDFFiles\"
	    strSourcePDF 				= strPDFAdvncePymtTmpltePath & strPDFFileName&".pdf"
	    strOutputPDF 				= strPDFAdvncePymtTmpltePath & Environment.Value("TestName")&"_"&strInvoiceType&".pdf"
	    strLogFilePath 				= strPDFAdvncePymtTmpltePath & "PDFFillLogFile.txt"
	    'Calling a method to fill PDF Files with input data
	    strFlag = ObjDotNetFactory.PDFFillForm (strSourcePDF,strOutputPDF,strInputFieldNames, strFieldValues,strLogFilePath)
	    If strFlag Then
	        Call fRptWriteReport("Pass", "Fill Advance Payment Template PDF File","Advance Payment Template PDF File is filled sucessfully, Output File Name :"&strOutputPDF)
	    Else   
	        Call fRptWriteReport("Fail", "Fill Advance Payment Template PDF File","Failed to fill Advance Payment Template PDF File, Input File Name :"&strSourcePDF)          
	    End If
On Error GoTo 0
End Function
'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 				:					fFillPDF
'	Objective							:					Used to fill Advance Payment Template PDF File
'	Input Parameters					:					strPDFFileName, strFieldValues -  String - Update the values in below order by | (pipeline) separated
'															Autodesk Entity Name, 
'															Name of Supplier, 
'															Supplier SAP ID, 
'															PO Number,
'															Payment Request Amount, 
'															Invoice Number, 
'															Business Reason for Advance Payment Request, 
'															Account Holder Name,
'															Bank Name,
'															Bank Account Number, 
'															Bank & Branch Code,
'															Invoice Date,
'															Payment DUE Date,
'															Special Instructions to be stipulated on the Paym, 
'															Company Code,
'															Payment Currency
'	Output Parameters					:					NIL
'	Date Created						:					
'	QTP Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
'Call fFillPDF(objDataDict, iRowCountRef, "CreditMemoInvoice")

'Call fFillPDF(objDataDict, iRowCountRef, "NormalInvoice")

'Call fFillPDF(objDataDict, iRowCountRef, "DownPaymentInvoice")

Public Function fFillPDF(objDataDict, iRowCountRef, strInvoiceType)
On Error Resume Next
	Dim intCompanyCode,intVendorName,intPrice,intQuntity,strCurrencyName,intPOnumber,intInvoiceNumber,dtInvoiceDate,dtDueDate,strPDFFileName,strPDFAdvncePymtTmpltePath,strSourcePDF,strOutputPDF
	intCompanyCode 	= objDataDict.Item("CompanyCode" & iRowCountRef)
    intCompanyCode 	= Left(intCompanyCode,4)
	intVendorName	= objDataDict.Item("Vendor" & iRowCountRef)
	intPrice	 	= objDataDict.Item("Price" & iRowCountRef)
	intQuntity	 	= objDataDict.Item("Quantity" & iRowCountRef)
	strCurrencyName	= objDataDict.Item("Currency" & iRowCountRef)
	intPOnumber 	= fGetSingleValue("AutoPONumber","TestData",Environment("TestName"))
	intPrice		= CInt(intPrice) * CInt(intQuntity)
	intUpdatedPrice = fGetSingleValue("Credit Memo Amount","TestData",Environment("TestName"))
    intNewQuantity = fGetSingleValue("Credit Memo Quantity","TestData",Environment("TestName"))
    
		If Len(intVendorName) >= 10 Then
			intVendorName 	= Right(Left(intVendorName,10),6)
		ElseIf Len(intVendorName) = 6 Then
			intVendorName = intVendorName
		Else
			Call fRptWriteReport("Fail","Enter Supplier SAP ID","Supplier SAP ID is Not Valid")
			ExitAction
		End If
		Call fGetCurrentDateFormatMMDDYYYY()
		dtInvoiceDate 	= Environment.Value("dtCurrentDate") ' OR fGetCurrentDateFormatMMDDYYYY
		dtDueDate		= Environment.Value("dtFutureDate")
		'PDF source file name
		strPDFFileName 	= "Source_AdvancePaymentTemplate"
		'Input and output PDF File paths
	    strPDFAdvncePymtTmpltePath 	= gstrFrameWorkFolder&"\"&gstrTestPlanName&"\Files\PDFFiles\"
	    strSourcePDF 				= strPDFAdvncePymtTmpltePath & strPDFFileName&".pdf"
	    strOutputPDF 				= strPDFAdvncePymtTmpltePath & Environment.Value("TestName")&"_"&strInvoiceType&".pdf"
 			If strInvoiceType = "NormalInvoice" Then
 				intInvoiceNumber= "IN" & intPOnumber
 				Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intInvoiceNumber,"TestData","AutoInvoiceNumber")
 				arrValues = Array("Autodesk Inc","GATES & COOPER, LLP.",intVendorName,intPOnumber,intPrice,intInvoiceNumber,"TEST","","","","",dtInvoiceDate,dtDueDate,"",intCompanyCode,strCurrencyName)
			ElseIf strInvoiceType = "CreditMemoInvoice" Then
				intInvoiceNumber= "CM" & intPOnumber
				'intPrice        = CInt(intPrice) - CInt(intUpdatedPrice)
		      	'Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intPrice,"TestData","CMPrice")
				Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intInvoiceNumber,"TestData","AutoCreditMemoNumber")
				arrValues = Array("Autodesk Inc","GATES & COOPER, LLP.",intVendorName,intPOnumber,intUpdatedPrice,intInvoiceNumber,"CREDIT MEMO","","","","",dtInvoiceDate,dtDueDate,"",intCompanyCode,strCurrencyName)
			ElseIf strInvoiceType = "DownPaymentInvoice" Then
				intInvoiceNumber= "DWN" & intPOnumber
				intPrice = 500
				Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intInvoiceNumber,"TestData","AutoInvoiceNumber")
				arrValues = Array("Autodesk Inc","GATES & COOPER, LLP.",intVendorName,intPOnumber,intPrice,intInvoiceNumber,"RPA TEST","","","","",dtInvoiceDate,dtDueDate,"",intCompanyCode,strCurrencyName)
			End  If
				'Invoice number enter into data sheet
				'Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intInvoiceNumber,"TestData","AutoInvoiceNumber")
			    'Handle Mouse/Keyboard events
			    Setting.WebPackage("ReplayType") = 2
			    'Test Objects to save PDF file
			    Set objPDFWindow = Description.Create()
			    objPDFWindow("micclass").value = "Window"
			    objPDFWindow("regexpwndclass").value = "Acrobat.*"
			    
			    Set objPDFDialog = Description.Create()
			    objPDFDialog("micclass").value = "Dialog"
			    objPDFDialog("regexpwndtitle").value = "Acrobat.*"
			        
			    Set objSaveAsWindow = Description.Create()
			    objSaveAsWindow("micclass").value = "Window"
			    objSaveAsWindow("regexpwndtitle").value = "Save As"
			    
			    Set objSaveAsDialog = Description.Create()
			    objSaveAsDialog("micclass").value = "Dialog"
			    objSaveAsDialog("regexpwndtitle").value = "Save As"
			    
			    Set objTreeView = Description.Create()
			    objTreeView("micclass").value = "WinTreeView"
			    objTreeView("regexpwndtitle").value = "Tree View"
			    objTreeView("regexpwndclass").value = "SysTreeView32"    
			    
			    Set objAddressBarEdit = Description.Create()
			    objAddressBarEdit("micclass").value = "WinEdit"
			    objAddressBarEdit("regexpwndclass").value = "Edit"
			    objAddressBarEdit("index").value = "0"
			    
			    Set objFileName = Description.Create()
			    objFileName("micclass").value = "WinEdit"
			    objFileName("attached text").value = "File name:"
			    
			    Set objSaveButton = Description.Create()
			    objSaveButton("micclass").value = "WinButton"
			    objSaveButton("regexpwndtitle").value = "&Save"
			        
			    Set objOverwriteDialog = Description.Create()
			    objOverwriteDialog("micclass").value = "Dialog"
			    objOverwriteDialog("regexpwndtitle").value = "Confirm Save As"
			    
			    Set objYesButton = Description.Create()
			    objYesButton("micclass").value = "WinButton"
			    objYesButton("regexpwndtitle").value = "&Yes"
			    
			    'Invoke PDF Document from source path
			    SystemUtil.Run strSourcePDF
			    wait 10
			    If Window(objPDFWindow).Exist Then
			        'Activating PDF Window
			        Window(objPDFWindow).Activate
			        'Creating object for Shell Scripting and Mercury Device Replay
			        Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
			        Set objWshell = CreateObject("Wscript.Shell")
			        'ENtering the values into PDF Fields
			        For intValue=0 to Ubound(arrValues)
			            objWshell.SendKeys "{TAB}"
			            myDeviceReplay.SendString arrValues(intValue)
			            Wait 3
			        Next
			        Wait 2
			        'Clicking on close PDF to see save popup
			        Window(objPDFWindow).Close
			        Wait 5
			        'Verify Save as window is launched
			        If Window(objPDFWindow).Dialog(objPDFDialog).Exist(5) Then
			            Window(objPDFWindow).Dialog(objPDFDialog).WinButton(objYesButton).Click
			            Wait 7
			            If Window(objPDFWindow).Window(objSaveAsWindow).Exist(5) Then
			                objWshell.SendKeys "{ENTER}"
			                Wait 10
			                'Verify Save as confirmation window is opened
			                If Window(objPDFWindow).Dialog(objSaveAsDialog).Exist(5) Then
			                    Window(objPDFWindow).Dialog(objSaveAsDialog).WebEdit(objFileName).Set ""
			                    Wait 1                    
			                    'File save as with scriptname+invoicetype in same path
			                    myDeviceReplay.SendString strOutputPDF
			                    Wait 3
			                    'Click on Save Button
			                    Window(objPDFWindow).Dialog(objSaveAsDialog).WinButton(objSaveButton).Click
			                    'Verify if overwrite dialog is opened and click on yes
			                    If Window(objPDFWindow).Dialog(objSaveAsDialog).Dialog(objOverwriteDialog).Exist(5) Then
			                        Window(objPDFWindow).Dialog(objSaveAsDialog).Dialog(objOverwriteDialog).WinButton(objYesButton).Click
			                    End If
			                    Wait 3
			                End If
			            End If
			        End If
			    End If
			    'Releasing Mouse/Keyboard events
			    Setting.WebPackage("ReplayType") = 1
On Error GoTo 0
End Function

'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 				:					fSendEmailInvoice
'	Objective							:					Used to send email to autodesk with PDF invoice type
'	Input Parameters					:					strRecipientsEMail, strEmailSubject, strEmailBody, strInvoiceType
'	Output Parameters					:					NIL
'	Date Created						:					05/28/2020
'	QTP Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   			05/29/2020
'******************************************************************************************************************************************************************************************************************************************		
''************ It's Email Send to Autodesk for NormalInvoice **************
'Call fSendEmailInvoice("NormalInvoice")
'Call fSendEmailInvoice("CreditMemoInvoice")
'Wait 720 ' After email send, we should wait minimum 10 min. It is as per manual TC step mentioned in TestRail

Public Function fSendEmailInvoice(objDataDict, iRowCountRef, strInvoiceType)
'On Error Resume Next	
	Dim objAccount
	Dim objMail
	Dim objFSO
	Dim intCompanyCode	
		'Data retrive from data sheet
		intCompanyCode 	= objDataDict.Item("CompanyCode" & iRowCountRef)
	    intCompanyCode 	= Left(intCompanyCode,4)
	    'strRecipientsEMail setup based on Company Code
	    Select Case intCompanyCode
	    	Case "1000"
	    	strRecipientsEMail = "test.AP.invoice.singapore@autodesk.com"
			Case "1001"	
			strRecipientsEMail = "test.AP.invoice.malaysia@autodesk.com"
			Case "1002"
			strRecipientsEMail = "test.AP.invoice.thailand@autodesk.com"
			Case "1003"	
			strRecipientsEMail = "test.AP.invoice.indonesia@autodesk.com"
			Case "1004"	
			strRecipientsEMail = "test.AP.invoice.philippines@autodesk.com"
			Case "1005"	
			strRecipientsEMail = "test.AP.invoice.vietnam@autodesk.com"
			Case "1100"	
			strRecipientsEMail = "test.AP.invoice.japan@autodesk.com"
			Case "1200"	
			strRecipientsEMail = "test.AP.invoice.australia@autodesk.com"
			Case "1320"	
			strRecipientsEMail = "test.AP.invoice.china@autodesk.com"
			Case "1330"	
			strRecipientsEMail = "test.AP.invoice.china@autodesk.com"
			Case "1400"	
			strRecipientsEMail = "test.AP.invoice.hongkong@autodesk.com"
			Case "1600"	
			strRecipientsEMail = "test.AP.invoice.taiwan@autodesk.com"
			Case "1700"	
			strRecipientsEMail = "test.AP.invoice.india@autodesk.com"
			Case "1800"	
			strRecipientsEMail = "test.AP.invoice.korea@autodesk.com"
			Case "2000"	
			strRecipientsEMail = "test.AP.invoice.switzerland@autodesk.com"
			Case "2006"	
			strRecipientsEMail = "test.AP.invoice.ireland@autodesk.com"
			Case "2010"	
			strRecipientsEMail = "test.AP.invoice.austria@autodesk.com"
			Case "2020"	
			strRecipientsEMail = "test.AP.invoice.netherlands@autodesk.com"
			Case "2025"	
			strRecipientsEMail = "test.AP.invoice.belgium@autodesk.com"
			Case "2030"	
			strRecipientsEMail = "test.AP.invoice.czech.republic@autodesk.com"
			Case "2040"	
			strRecipientsEMail = "test.AP.invoice.uk@autodesk.com"
			Case "2041"	
			strRecipientsEMail = "test.AP.invoice.uk@autodesk.com"
			Case "2042"	
			strRecipientsEMail = "test.AP.invoice.uae@autodesk.com"
			Case "2043"	
			strRecipientsEMail = "test.AP.invoice.jordan@autodesk.com"
			Case "2050"	
			strRecipientsEMail = "test.AP.invoice.france@autodesk.com"
			Case "2060"	
			strRecipientsEMail = "test.AP.invoice.germany@autodesk.com"
			Case "2070"	
			strRecipientsEMail = "test.AP.invoice.italy@autodesk.com"
			Case "2080"	
			strRecipientsEMail = "test.AP.invoice.spain@autodesk.com"
			Case "2090"	
			strRecipientsEMail = "test.AP.invoice.sweden@autodesk.com"
			Case "2100"	
			strRecipientsEMail = "test.AP.invoice.switzerland@autodesk.com"
			Case "2104"	
			strRecipientsEMail = "test.AP.invoice.south.africa@autodesk.com"
			Case "2106"	
			strRecipientsEMail = "test.AP.invoice.qatar@autodesk.com"
			Case "2130"	
			strRecipientsEMail = "test.AP.invoice.ireland@autodesk.com"
			Case "2140"	
			strRecipientsEMail = "test.AP.invoice.switzerland@autodesk.com"
			Case "2150"	
			strRecipientsEMail = "test.AP.invoice.netherlands@autodesk.com"
			Case "2160"	
			strRecipientsEMail = "test.AP.invoice.hungary@autodesk.com"
			Case "2170"	
			strRecipientsEMail = "test.AP.invoice.poland@autodesk.com"
			Case "2180"	
			strRecipientsEMail = "test.AP.invoice.turkey@autodesk.com"
			Case "2190"	
			strRecipientsEMail = "test.AP.invoice.russia@autodesk.com"
			Case "2200"	
			strRecipientsEMail = "test.AP.invoice.uk@autodesk.com"
			Case "2240"	
			strRecipientsEMail = "test.AP.invoice.netherlands@autodesk.com"
			Case "2250"	
			strRecipientsEMail = "test.AP.invoice.uk@autodesk.com"
			Case "2911"	
			strRecipientsEMail = "test.AP.invoice.israel@autodesk.com"
			Case "2921"	
			strRecipientsEMail = "test.AP.invoice.romania@autodesk.com"
			Case "2935"	
			strRecipientsEMail = "test.AP.invoice.saudi.arabia@autodesk.com"
			Case "2936"	
			strRecipientsEMail = "test.AP.invoice.denmark@autodesk.com"
			Case "3000"	
			strRecipientsEMail = "test.AP.invoice.usa@autodesk.com"
			Case "3011"	
			strRecipientsEMail = "test.AP.invoice.usa@autodesk.com"
			Case "3500"	
			strRecipientsEMail = "test.AP.invoice.canada@autodesk.com"
			Case "3600"	
			strRecipientsEMail = "test.AP.invoice.mexico@autodesk.com"
			Case "3610"	
			strRecipientsEMail = "test.AP.invoice.argentina@autodesk.com"
			Case "3620"	
			strRecipientsEMail = "test.AP.invoice.brazil@autodesk.com"
			Case "3650"	
			strRecipientsEMail = "test.AP.invoice.colombia@autodesk.com"
			Case "3660"	
			strRecipientsEMail = "test.AP.invoice.mexico@autodesk.com"
	    End Select
	    If strInvoiceType = "DownPaymentInvoice" Then
	    	strRecipientsEMail = "test.ap.advpayment@autodesk.com"
	    End If
		'strEmailSubject and strEmailBody setup based on Invoice Type
		strEmailSubject = strInvoiceType & " Generation Request"
		strEmailBody = "Hi Team, "& vbLf &"This is "& strInvoiceType &" Generation Request...."
		'attachement saved path
		strPDFAdvncePymtTmpltePath = gstrFrameWorkFolder&"\"&gstrTestPlanName&"\Files\PDFFiles\"
		strAttachmentPath = strPDFAdvncePymtTmpltePath & Environment.Value("TestName")&"_"&strInvoiceType&".pdf"
		'Verify if TestPlan exists in TestPlan Folder
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FileExists(strAttachmentPath) = True then
			'Creating an object for Outlook Applixation
			Set objOutlook = CreateObject("Outlook.Application")
			'Iterating through all the emails accounts configured in the system
			For Each objAccount In objOutlook.Session.Accounts
				 If objAccount.AccountType = olPop3 Then 
				 	Set objMail = objOutlook.CreateItem(olMailItem)
					'Recipients details	 	
					objMail.Recipients.Add(strRecipientsEMail)
					'Email Subject
					objMail.Subject = strEmailSubject
					'Body Text
					objMail.Body= strEmailBody
					'Adding attachement to the email
					objMail.Attachments.Add(strAttachmentPath)
					objMail.Recipients.ResolveAll
				 	'Verify the email id and send an email with respective mail id
				 	If Instr(objAccount,"autodesk") Then
				 		Set objMail.SendUsingAccount = objAccount
				     	objMail.Send
				     	Exit for
				 	End If
				 End If 
			Next
		End If
'On Error GoTo 0	
End Function
'******************************************************************************************************************************************************************************************************************************************
'	Function Name		 				:					fGetCurrentDateFormatMMDDYYYY
'	Objective							:					Used to Get Current Date and future date format as MM/DD/YYYY
'	Input Parameters					:					N/A
'	Output Parameters					:					Environment.Value("dtCurrentDate") OR fGetCurrentDateFormatMMDDYYYY, and Environment.Value("dtFutureDate")
'	Date Created						:					06/01/2020
'	QTP Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti
'	Modification Date					:		   			
'******************************************************************************************************************************************************************************************************************************************		
Public Function fGetCurrentDateFormatMMDDYYYY()
On Error Resume Next
	dtDay 	= Day(Date)
	dtMonth = Month(Date)
	dtYear 	= Year(Date)
		If Len(dtDay) = 1 Then
			dtDay = "0"&dtDay
		End If
		If Len(dtMonth) = 1 Then
			dtMonth = "0"&dtMonth
		End If
		dtInvoiceDate 	= dtMonth&"/"&dtDay&"/"&dtYear
		dtDueDate		= dtMonth&"/"&dtDay + 5&"/"&dtYear
		fGetCurrentDateFormatMMDDYYYY = dtInvoiceDate
		Environment.Value("dtCurrentDate") = dtInvoiceDate
		Environment.Value("dtFutureDate") = dtDueDate 'Replace(DateAdd("m",5,dtInvoiceDate),"-","/")
On Error GoTo 0	
End Function



'******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					ReadGlobalConfigFile
'	Objective							:					Used to read global Config XML File
'	Input Parameters					:					Nil
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0				
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti Technologies
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fReadGlobalConfigFile()

	Dim ObjXML   
	Dim objRoot  
	Dim objNodeList  
	Dim strName  
	Dim strValue
	Dim strRerunFailuresFlag
	Dim strRerunIterations
	
	Set ObjXML = CreateObject ("Microsoft.XMLDOM")   
	ObjXML.async = False     
	ObjXML.load(gstrGlobalConfigPath)  
	Set objRoot = ObjXML.documentElement   
	Set objNodeList = objRoot.getElementsByTagName("Variable")   
	For Each Elem In objNodeList   
		Set strName = Elem.getElementsByTagName ("Name")(0)  
		Set strValue = Elem.getElementsByTagName ("Value")(0)
		
		If strName.Text = "Environment" Then
			gstrEnvironmentName = strValue.Text
		End If
	Next
	
End Function



'******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					ReadFrameworkConfig
'	Objective							:					Used to framework Config XML File
'	Input Parameters					:					Nil
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0				
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti Technologies
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fReadFrameworkConfig()
	
	Dim objXMLDoc
	Dim objXMLRoot
	Dim objAllEnvironmentNodes
	Dim intNode
	Dim objEnvironmentNode
	Dim strEnvironmentName
	
	Set objXMLDoc = XMLUtil.CreateXMLFromFile(gstrProjectConfigFilePath)
	Set objXMLRoot = objXMLDoc.GetRootElement()
	
	'Get all the company node
	Set objAllEnvironmentNodes = objXMLRoot.ChildElementsByPath("Environment")
	
	'Iterate through all environments
	For intNode = 1 To objAllEnvironmentNodes.Count
	    Set objEnvironmentNode = objAllEnvironmentNodes.Item(intNode)
	    
	    'Get all the attribute of company node
	    Set attributeCol = objEnvironmentNode.Attributes
	    strFrameworkConfgEnvName = attributeCol.ItemByName("Name").Value
	    
	    'Reading default and respective environment values
	    If Trim(strFrameworkConfgEnvName) = "DEFAULT" OR Trim(strFrameworkConfgEnvName) = Trim(gstrEnvironmentName) Then
	    	If Trim(strFrameworkConfgEnvName) = "DEFAULT" Then
	    		gstrChromeBrowser = objEnvironmentNode.ChildElementsByPath("ChromeBrowser").Item(1).Value
				gstrIEBrowser = objEnvironmentNode.ChildElementsByPath("IEBrowser").Item(1).Value
				gstrFireFoxBrowser = objEnvironmentNode.ChildElementsByPath("FireFoxBrowser").Item(1).Value   	
	    	ElseIf Trim(strFrameworkConfgEnvName) = Trim(gstrEnvironmentName) Then	
	    		'Reading Ariba Environment Details
				gstrAribaBuyerURL = objEnvironmentNode.ChildElementsByPath("AribaBuyerURL").Item(1).Value
				gstrAribaBuyerUsername = objEnvironmentNode.ChildElementsByPath("AribaBuyerUsername").Item(1).Value
				gstrAribaBuyerPassword = objEnvironmentNode.ChildElementsByPath("AribaBuyerPassword").Item(1).Value
				gstrAribaSupplierURL = objEnvironmentNode.ChildElementsByPath("AribaSupplierURL").Item(1).Value
				gstrAribaSupplierUsername = objEnvironmentNode.ChildElementsByPath("AribaSupplierUsername").Item(1).Value
				gstrAribaSupplierPassword = objEnvironmentNode.ChildElementsByPath("AribaSupplierPassword").Item(1).Value
				'Reading Fiori Environment Details
				gstrFioriLanguage = objEnvironmentNode.ChildElementsByPath("FioriLanguage").Item(1).Value
				gstrFioriAnalystAPURL = objEnvironmentNode.ChildElementsByPath("FioriAnalystAPURL").Item(1).Value
				gstrFioriAnalystAPUsername = objEnvironmentNode.ChildElementsByPath("FioriAnalystAPUsername").Item(1).Value
				gstrFioriAnalystAPPassword = objEnvironmentNode.ChildElementsByPath("FioriAnalystAPPassword").Item(1).Value
				gstrFioriManagerAPURL = objEnvironmentNode.ChildElementsByPath("FioriManagerAPURL").Item(1).Value
				gstrFioriManagerAPUsername = objEnvironmentNode.ChildElementsByPath("FioriManagerAPUsername").Item(1).Value
				gstrFioriManagerAPPassword = objEnvironmentNode.ChildElementsByPath("FioriManagerAPPassword").Item(1).Value
				'Reading Concur Environment Details
				gstrConcurEmployeeURL = objEnvironmentNode.ChildElementsByPath("ConcurEmployeeURL").Item(1).Value
				gstrConcurEmployeeUsername = objEnvironmentNode.ChildElementsByPath("ConcurEmployeeUsername").Item(1).Value
				gstrConcurEmployeePassword = objEnvironmentNode.ChildElementsByPath("ConcurEmployeePassword").Item(1).Value
				gstrConcurManagerURL = objEnvironmentNode.ChildElementsByPath("ConcurManagerURL").Item(1).Value
				gstrConcurManagerUsername = objEnvironmentNode.ChildElementsByPath("ConcurManagerUsername").Item(1).Value
				gstrConcurManagerPassword = objEnvironmentNode.ChildElementsByPath("ConcurManagerPassword").Item(1).Value
				'Reading SAP ECC Environment Details
				gstrSAPECCClient = objEnvironmentNode.ChildElementsByPath("SAPECCClient").Item(1).Value
				gstrSAPECCUsername = objEnvironmentNode.ChildElementsByPath("SAPECCUsername").Item(1).Value
				gstrSAPECCPassword = objEnvironmentNode.ChildElementsByPath("SAPECCPassword").Item(1).Value
				gstrSAPECCLanguage = objEnvironmentNode.ChildElementsByPath("SAPECCLanguage").Item(1).Value
				gstrSAPECCServer = objEnvironmentNode.ChildElementsByPath("SAPECCServer").Item(1).Value
				gstrSAPECCInstance = objEnvironmentNode.ChildElementsByPath("SAPECCInstance").Item(1).Value
				gstrSAPECCApplication = objEnvironmentNode.ChildElementsByPath("SAPECCApplication").Item(1).Value
	    	End If
	    End If	    
	Next

End Function

