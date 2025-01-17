'*********************************************************************************************************************************************************************
'    Function Name                                :                    fAribaLogin
'    Objective                                    :                    Log-in Ariba Application 
'    Input Parameters                             :                     Nil
'    Output Parameters                            :                     Nil
'    Date Created                                 :                     
'    UFT/QTP Version                              :                     UFT 15.0
'    Pre-requisites                               :                     NIL  
'    Created By                                   :                     Cigniti
'    Modification Date                            :  
'*********************************************************************************************************************************************************************
Public Function fAribaLogin()    
    On error resume next    
    'Verify if Step Failed, If yes, it will Exit from function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
       
    If gstrIEBrowser = "YES" Then
        SystemUtil.Run "iexplore.exe",gstrAribaBuyerURL,,,3  
    ElseIf gstrChromeBrowser = "YES" Then        
        SystemUtil.Run "Chrome.exe",gstrAribaBuyerURL,,,3
        Wait(MIN_WAIT)
    ElseIf gstrFireFoxBrowser = "YES" Then
        SystemUtil.Run "firefox.exe",gstrAribaBuyerURL,,,3
    End If
    
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MID_WAIT)	       
	Set objBrwAndPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
	objBrwAndPage.Sync
    If (objBrwAndPage.WebEdit("txtUserName").Exist) and (objBrwAndPage.WebEdit("txtPassword").Exist) Then
        objBrwAndPage.Sync                
        Call fRptWriteReport("PASSWITHSCREENSHOT","Verify Ariba Application login page","Ariba login page is launched successfully")
    Else
        Call fRptWriteReport("Fail","Verify Ariba Application login page","Ariba login page is not displayed")
        Call fRptWriteResultsSummary() 
        Exit Function
    End  If
    
    'Enter data into Username txt field
    Call fEnterText(objBrwAndPage.WebEdit("txtUserName"),gstrAribaBuyerUsername,"UserName")
    
    'Enter data into Password txt field
    Call fEnterText_SetSecureMode(objBrwAndPage.WebEdit("txtPassword"),gstrAribaBuyerPassword,"Password")    
    'Click on Login button
    Call fClick(objBrwAndPage.WebButton("btnLogin"), "Login")
    objBrwAndPage.Sync
    Wait MIN_WAIT
    
     Set objBrwAndPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
     objBrwAndPage.Sync
    
    'Verification of Home page
    If fVerifyObjectExist(objBrwAndPage.Image("imgCompanyLogo")) Then
        Call fRptWriteReport("PASSWITHSCREENSHOT","Verify Ariba Application login","Ariba Application is login successfully with Username "&gstrAribaApplicationUsername)
    Else
        Call fRptWriteReport("Fail","Verify Ariba Application login","Ariba Application is not login with Username "&gstrAribaApplicationUsername)
        Environment("ERRORFLAG") = False
        Call fRptWriteResultsSummary() 
        Exit Function 
    End If  
On error goto 0    
       
End Function

'*********************************************************************************************************************************************************************
'    Function Name								:					fAribaLogOut
'    Objective									:					Log out from Ariba Application
'    Input Parameters							: 					Nil
'    Output Parameters							: 					Nil
'    Date Created								: 					
'    UFT/QTP Version							: 					UFT 15.0
'    Pre-requisites								: 					NIL  
'    Created By									: 					Cigniti
'    Modification Date							:  
'*********************************************************************************************************************************************************************
Public Function fAribaLogout()
	
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Set Browser and Page 
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MID_WAIT)
    Set objBrwAndPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
    objBrwAndPage.Sync
	
	'Wait till Logout btn is displayed
	If fSynUntilObjExists(objBrwAndPage.Link("lnkLogOutBtn"),MID_WAIT) Then	
		'Click on Logout btn link
	    Call fClick(objBrwAndPage.Link("lnkLogOutBtn"), "UserLogout")		
		'Click on Log on button
		Call fClick(objBrwAndPage.Link("lnkLogOut"), "Logout")
		Wait(MIN_WAIT)
		Call fSynUntilObjExists(objBrwAndPage.WebEdit("txtUserName"),MID_WAIT)		
			'Verification of Home page
			If fVerifyObjectExist(objBrwAndPage.WebEdit("txtUserName")) Then
				Call fRptWriteReport("PASSWITHSCREENSHOT", "Logout  from Ariba page","Successfully logged out from Ariba application")
			Else
				Call fRptWriteReport("Fail", "Logout  from Ariba page","Unable to logout from Ariba application")
				Call fRptWriteResultsSummary()        
	        	Exit Function
			End  If
	End  If	
	'Close all Open browsers
    Call fCloseAllOpenBrowsers("ALL")     
    On error goto 0
End Function

'******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fSelectByCategory
'	Objective							:					Used to Select the item by Catogory
'	Input Parameters					:					strMenuItemText,strSubMenuItemText
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	Pre-requisites						:					NILL  
'	Created By							:					Cigniti					
'	Modification Date					:		   		
'******************************************************************************************************************************************************************************************************************************************		
Public Function fSelectByCategory(strMenuItemText,strSubMenuItemText)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	If fClick(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement").WebTable("outertext:="&Ucase(strMenuItemText),"html tag:=TABLE"),strMenuItemText) Then
		fSelectByItem = false
			If fClick(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement").Link("outertext:="&Ucase(strSubMenuItemText),"html tag:=A"),strSubMenuItemText) Then
				fSelectByItem = True
				'Updated by sravanthi : 26-May
				Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify navigated to "&strMenuItemText&"->"&strSubMenuItemText,"Navigated successfully to "&strMenuItemText&"->"&strSubMenuItemText)
				
				fSelectByCategory = True
			Else
				Call fRptWriteReport("Fail","Verify page navigated to Create Requisition Screen","Failed to navigate Create Requisition Screen")
				Call fRptWriteResultsSummary() 
				Exit Function
			End If
	Else
		Call fRptWriteReport("Fail","Verify page navigated to Create Requisition Screen","Failed to navigate Create Requisition Screen")
	    Call fRptWriteResultsSummary() 
	    Exit Function
	End If
On error goto 0	
End Function

'******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fChooseValueforSearchField
'	Objective							:					Used to Choose value in the Popup page eg: Supplier,CompanyCode
'	Input Parameters					:					strType,strSearchValue
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					UFT 15.0
'	QC Version							:		
'	Pre-requisites						:					NILL  
'	Created By							:					
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fChooseValueforSearchField(strType,strSearchValue)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	fChooseValueforSearchField = False
	
	objPage.Sync
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)
	'Set the object to a page
	Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
	
	If fSynUntilObjExists(objPage.WebElement("weChoosePopupWindow"),MID_WAIT) = True Then			
		'Select Type in list eg: Name,CompanyCode
		fSelect objPage.SAPWebExtList("WlstType"),strType,"listType"
		'Enter value in search field
		fEnterText objPage.WebEdit("txtSearchfield"),strSearchValue,"SearchName"
		'Click on Search button
		fClick objPage.WebButton("btnSearch"),"SearchButton"
		Wait 3
		objPage.Sync
		introw =objPage.WebTable("wbtSelectvalue").RowCount 
		indexcount = 0
		For icount  = 1 To introw
			If objPage.WebTable("wbtSelectvalue").GetCellData(icount,1) <> "" Then
				indexcount = indexcount+1
			End If			
			'Verify the Value in the table
			If Instr(objPage.WebTable("wbtSelectvalue").GetCellData(icount,1),strSearchValue) > 0 Then				
				flgSearchItem = True		
				'Click on Select button
				objPage.WebButton("btnSelect").SetTOProperty "index",indexcount-1
				fClick objPage.WebButton("btnSelect"),"SelectButton"				
				Exit for
			End IF	
		Next
		
		If flgSearchItem Then
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Page navigated to 'Choose Value' popup window","Navigated successfully and " &strSearchValue& "  Value has been selected")
		Else
			Call fRptWriteReport("Fail", "Verify Page navigated to 'Choose Value' Popup window","Failed to Navigated 'Choose Value' Popup window and " &strSearchValue& " value is not selected")			
			Call fRptWriteResultsSummary() 
			Exit Function
		End If	
	Else
		Call fRptWriteReport("Fail", "Verify Page navigated to 'Choose Value' Popup window","Failed to Navigated 'Choose Value' Popup window and " &strSearchValue& " value is not selected")			
		Call fRptWriteResultsSummary() 
		Exit Function
	End If
	On error goto 0
End Function
'******************************************************************************************************************************************************************************************************************************************
'    Function Name                        :                fSelectAribaMainMenuItems
'    Objective                            :                Used to select submenu Items in Ariba Application
'    Input Parameters                     :                strMenuItemText,strItemSelection
'    Output Parameters                    :                Nil
'    Date Created                         :                29/04/2020
'    UFT/QTP Version                      :                UFT 15.0
'    Pre-requisites                       :                NIL  
'    Created By                           :                Cigniti
'    Modification Date                    :       
'***************************************************************************************************************************************************************************************
Public Function fSelectAribaMainMenuItems(strMenuItemText,strItemSelection)  
     On error resume next   
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	     
	'Variable Declaration
	Dim objHomePage, objWbtChildObject, oChildObjectCount,flgSelectionItem
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MID_WAIT)
	Set objHomePage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")    
    'Click on Menu
    If fClick(objHomePage.WebElement("innertext:="&strMenuItemText&".*","html tag:=A","role:=tab"),strMenuItemText) Then        
        'Click on Menu  
		Set objWbtChildObject = description.Create
		objWbtChildObject("micclass").value = "link"
		intItemSelectionCount = 0
		Call fSynUntilObjExists(objHomePage.WebTable("wbtCommonActions"),MID_WAIT)
		Set oChildObjectCount = objHomePage.WebTable("wbtCommonActions").ChildObjects(objWbtChildObject)		
		For jCount = 0 To oChildObjectCount.count-1	
			If oChildObjectCount(jCount).getroproperty("innertext") = strItemSelection Then
				oChildObjectCount(jCount).click	
				flgSelectionItem = True
				Exit for
			End If	
		Next        
        If flgSelectionItem Then
            Call fRptWriteReport("Pass", "Verify page navigated to "&strMenuItemText&"->"&strItemSelection,"Navigated successfully to "&strMenuItemText&"->"&strItemSelection)
            fSelectAribaMainMenuItems = True
        Else
            Call fRptWriteReport("Fail","Verify page navigated to "&strMenuItemText&"->"&strItemSelection,"Failed to navigate "&strMainMenuInnerText&"->"&strSubMenuName)
            Call fRptWriteResultsSummary() 
            Exit Function
        End If
    Else
    	Call fRptWriteReport("Fail","Verify navigated to "&strMenuItemText&"->"&strItemSelection,"Failed to navigate "&strMenuItemText)
        Call fRptWriteResultsSummary() 
        Exit Function
    End If    
    Set objHomePage			=	Nothing
    Set objWbtChildObject	=	Nothing
    Set oChildObjectCount	=	Nothing  
On error goto 0    
End Function
'******************************************************************************************************************************************************************************
'	Function Name						:		fAribaSupplierLogin
'	Objective							:		Used to Open Ariba Supplier login
'	Input Parameters					:		strBrName -  Browser Name 
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fAribaSupplierLogin(objDataDict,iRowCountRef)
	On error resume next		
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	

	intCompanyCode 	= objDataDict.Item("Company Code" & iRowCountRef)
'	intCompanyCode 	= Left(intCompanyCode,4)

  Select Case Trim(intCompanyCode)
  		Case  "3000"
  		gstrAribaSupplierUsername =	"us1_adsk@autodesk.com"
  		Case  "1000"
  		gstrAribaSupplierUsername =	"sg1_adsk@autodesk.com"
		Case  "2080"
		gstrAribaSupplierUsername =	"es1@autodesk.com"
		Case  "3500"	
		gstrAribaSupplierUsername =	"ca4@autodesk.com"
		Case "1100"
		gstrAribaSupplierUsername =	"jp1@autodesk.com"	
		Case "2006"
		gstrAribaSupplierUsername =	"ie1@autodesk.com"			
		Case "2050"
		gstrAribaSupplierUsername =	"fr5@autodesk.com"		
		Case "2040"
		gstrAribaSupplierUsername =	"uk4@autodesk.com"				
  End  Select
	
	
	
   If gstrIEBrowser = "YES" Then
        SystemUtil.Run "iexplore.exe",gstrAribaSupplierURL,,,3  
    ElseIf gstrChromeBrowser = "YES" Then        
        SystemUtil.Run "Chrome.exe",gstrAribaSupplierURL,,,3
        Wait(MIN_WAIT)
    ElseIf gstrFireFoxBrowser = "YES" Then
        SystemUtil.Run "firefox.exe",gstrAribaSupplierURL,,,3
    End If
    
 	Set objBrwAndPage = Browser("brAribaSpendManagement").Page("pgAribaNetworkSupplier")    
	    If (objBrwAndPage.WebEdit("txtUserName").Exist) and (objBrwAndPage.WebEdit("txtPassword").Exist) Then
	        objBrwAndPage.Sync                
	        Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify 'Ariba Network Supplier' login page exist","'Ariba Network Supplier' login screen launched successfully")
	    Else
	    	Call fRptWriteReport("Fail", "Verify Ariba Network Supplier login page exist","Unable to launch the Ariba Network Supplier application")
	       	Call fRptWriteResultsSummary() 
            Exit Function
	    End  If
	  'Enter data into Username txt field
    Call fEnterText(objBrwAndPage.WebEdit("txtUserName"),gstrAribaSupplierUsername,"UserName")
	'Enter data into Password txt field
	Call fEnterText_SetSecureMode(objBrwAndPage.WebEdit("txtPassword"),gstrAribaSupplierPassword,"Password")
	'Click on Login button
	Call fClick(objBrwAndPage.WebButton("btnLogin"), "Login")
	Call fSynUntilObjExists(objBrwAndPage.WebElement("weAribaNetwork"),MAX_WAIT)
	Wait(2)
		'Verification of Home page
	    If fVerifyObjectExist(objBrwAndPage.WebElement("weAribaNetwork")) Then
	    	Call fRptWriteReport("PASSWITHSCREENSHOT", "User log-in to 'Ariba Network Supplier' application", "User "&chr(34)&strAribaSupplierUserName&chr(34)&" should be able to log-in to the application")
	    Else
			Call fRptWriteReport("Fail", "User log-in to 'Ariba Network Supplier' application", "User "&chr(34)&strAribaSupplierUserName&chr(34)&" is unable to log-in to the application")
			Environment("ERRORFLAG") = False
			Call fRptWriteResultsSummary() 
            Exit Function     
		End If 
On error goto 0		
End Function
'******************************************************************************************************************************************************************************
'	Function Name						:		fAribaGeneratePurchaseOrder
'	Objective							:		Used to Capture the PO Number
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fAribaGeneratePurchaseOrder(objDataDict,iRowCountRef)
    On error resume next
    
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	
    'Clear the testdata sheet
     Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"","TestData","AutoPONumber")
     
     'Read Data from excel
    strSerachScreen = objDataDict.Item("SearchScreen" & iRowCountRef)
    strReqID = fGetSingleValue("RequisitionNumber","TestData",Environment("TestName"))
    Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MID_WAIT)
    Set objHomePage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
    flgSearchTextInAribaPage = fSearchTextInAribaPage (strSerachScreen)

    If flgSearchTextInAribaPage Then
         Call fRptWriteReport("Pass", "Verify Clicked on "&strSerachScreen,"Clicked successfully on "&strSerachScreen)
        
            'To Validate Whether the screen is navigated to required Screen    
            flgNavigateScreen = fVerifyProperty(objHomePage.SAPWebExtList("html tag:=DIV","name:=_wh1eo"),"selection",strSerachScreen)            
           		 If flgNavigateScreen Then
	                Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Screen navigation","Navigated to "&strSerachScreen&" Screen Successfully")
	'			   'Enter RequsitionID
					Call fSynUntilObjExists(objHomePage.WebEdit("acc_name:=Requisition ID:","html tag:=INPUT"),MIN_WAIT)
	                objHomePage.WebEdit("acc_name:=Requisition ID:","html tag:=INPUT").set strReqID
	                objHomePage.WebButton("class:=w-btn w-btn-primary aw7_w-btn-primary","html tag:=BUTTON","name:=Search").Click
	                Call fSynUntilObjExists(objHomePage.WebTable("wbtOrders"),MID_WAIT)
	                'To fetch Purchase Order ID
	                objHomePage.WebTable("wbtOrders").Highlight
	                strOrderID = objHomePage.WebTable("wbtOrders").GetCellData(2,2)   
'						 Get strOrderID , If OrderID is not created then please Click on Search button and Get OrderID number
						If Isnull(strOrderID) or IsEmpty(strOrderID) or Len(strOrderID) < 1 Then
							objHomePage.WebEdit("acc_name:=Requisition ID:","html tag:=INPUT").set strReqID
			                objHomePage.WebButton("class:=w-btn w-btn-primary aw7_w-btn-primary","html tag:=BUTTON","name:=Search").Click
			                Call fSynUntilObjExists(objHomePage.WebTable("wbtOrders"),MID_WAIT)
			                Call fSynUntilObjExists(objHomePage.WebTable("wbtOrders"),MID_WAIT)
			                'To fetch Purchase Order ID	                
			                strOrderID = objHomePage.WebTable("wbtOrders").GetCellData(2,2)   
						End If
		                If isnumeric(strOrderID) Then
		'                     Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Genrated PO ID","Generated PO ID: "&strOrderID&" Successfully")
							  Call fRptWriteReport("PASSWITHSCREENSHOT","Verify PurchaseOrder ID","PurchaseOrder ID is Generated as "&strOrderID)
		                      fAribaGeneratePurchaseOrder = strOrderID
		'                    Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strOrderID,"TestData","AutoPONumber")
							Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,cdbl(strOrderID),"TestData","AutoPONumber")
							
							Call fAribaGetPaymentTermID() ' Get Payment Term ID
							
		                Else
		                    Call fRptWriteReport("Fail", "Verify Genrated PO ID","Failed to Generate PO ID")
		                    Call fRptWriteResultsSummary() 
		                    Exit Function
		                End If                    
            	Else
	                Call fRptWriteReport("Fail", "Verify Navigated to "&strSerachScreen&" Screen","Failed to Navigated to "&strSerachScreen&" Screen")
	                Call fRptWriteResultsSummary() 
	                Exit Function
            	End If        
    Else
        Call fRptWriteReport("Fail", "Verify Clicked on "&strSerachScreen,"Failed to Click on "&strSerachScreen)
        Call fRptWriteResultsSummary() 
        Exit Function        
    End If
    On error goto 0
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fAribaGeneratePurchaseOrder
'	Objective							:		Used to Capture the PO Number
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fAribaGeneratePurchaseOrderTwo(objDataDict,iRowCountRef)
    On error resume next
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    
    'Clear the testdata sheet
     Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"","TestData","AutoPONumber")
     
     'Read Data from excel
    strSerachScreen = objDataDict.Item("SearchScreen" & iRowCountRef)    
    strReqID = fGetSingleValue("RequisitionNumber","TestData",Environment("TestName"))
     Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MID_WAIT)
    Set objHomePage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
    flgSearchTextInAribaPage = fSearchTextInAribaPage (strSerachScreen)

    If flgSearchTextInAribaPage Then
         Call fRptWriteReport("Pass", "Verify Clicked on "&strSerachScreen,"Clicked successfully on "&strSerachScreen)
        
            'To Validate Whether the screen is navigated to required Screen    
            flgNavigateScreen = fVerifyProperty(objHomePage.SAPWebExtList("html tag:=DIV","name:=_wh1eo"),"selection",strSerachScreen)            
           		 If flgNavigateScreen Then
	                Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Screen navigation","Navigated to "&strSerachScreen&" Screen Successfully")
	'			   'Enter RequsitionID
	                objHomePage.WebEdit("acc_name:=Requisition ID:","html tag:=INPUT").set strReqID
	                objHomePage.WebButton("class:=w-btn w-btn-primary aw7_w-btn-primary","html tag:=BUTTON","name:=Search").Click
	                Call fSynUntilObjExists(objHomePage.WebTable("wbtOrders"),MID_WAIT)
	                'To fetch Purchase Order ID
	                objHomePage.WebTable("wbtOrders").Highlight
	                strOrderID = objHomePage.WebTable("wbtOrders").GetCellData(2,2)   
'						05/19/2020 - Get strOrderID , If OrderID is not created then please Click on Search button and Get OrderID number
						If Isnull(strOrderID) or IsEmpty(strOrderID) or Len(strOrderID) < 1 Then
							objHomePage.WebEdit("acc_name:=Requisition ID:","html tag:=INPUT").set strReqID
			                objHomePage.WebButton("class:=w-btn w-btn-primary aw7_w-btn-primary","html tag:=BUTTON","name:=Search").Click
			                Call fSynUntilObjExists(objHomePage.WebTable("wbtOrders"),MID_WAIT)
			                Call fSynUntilObjExists(objHomePage.WebTable("wbtOrders"),MID_WAIT)
			                'To fetch Purchase Order ID	                
			                strOrderID = objHomePage.WebTable("wbtOrders").GetCellData(2,2)   
						End If
		                If isnumeric(strOrderID) Then
		                     Call fRptWriteReport("PASSWITHSCREENSHOT","Verify PurchaseOrder ID","PurchaseOrder ID is Generated as"&strOrderID)
		                      fAribaGeneratePurchaseOrder = strOrderID
		  					  Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strOrderID,"TestData","AutoPONumber")
							
		                Else
		                    Call fRptWriteReport("Fail", "Verify Genrated PO ID","Failed to Generate PO ID")
		                    Call fRptWriteResultsSummary() 
		                    Exit Function
		                End If                    
            	Else
	                Call fRptWriteReport("Fail", "Verify Navigated to "&strSerachScreen&" Screen","Failed to Navigated to "&strSerachScreen&" Screen")
	                Call fRptWriteResultsSummary() 
	                Exit Function
            	End If        
    Else
        Call fRptWriteReport("Fail", "Verify Clicked on "&strSerachScreen,"Failed to Click on "&strSerachScreen)
        Call fRptWriteResultsSummary() 
        Exit Function        
    End If
    
    On error goto 0
End Function

'******************************************************************************************************************************************************************************************************************************************
''    Function Name                         	:                fSearchTextInAribaPage
''    Objective                                	:                Used to Search Text In Ariba Page 
''    Input Parameters                        	:                strSearchText
''    Output Parameters                     	:                Nil
''    Date Created                          	:                28/April/2020
''    UFT/QTP Version                         	:                15.0
''    Pre-requisites                        	:                NIL  
''    Created By                            	:                Cigniti
''    Modification Date                        	:                   
'*************************************************************************************************************************************************************************************** 
Public Function fSearchTextInAribaPage(strSearchText)
	On error resume next	
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Variable declaration
	Dim objHomePage,objWbtSearchItems,objSearchItems,flgSearchItem
	
	flgSearchItem = False	
    ' Set Browser and Page 
    Set objHomePage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
    fSearchTextInAribaPage = False
    'Click on Search button
    If fClick(objHomePage.WebElement("weSearch"),"Search button") Then
        Wait MIN_WAIT
        'Select the Item in Search Results
        
        flgSearchItem = false
		Set objWbtSearchItems = description.Create
		objWbtSearchItems("micclass").value = "link"			
        Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement").WebTable("wbtSearchMenuItems"),MID_WAIT)
        Call fRptWriteReport("Pass","Verify Click on Search Button in Ariba page","Successfully Clicked on Search Button in Ariba Page")
        
		set objSearchItems = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement").WebTable("wbtSearchMenuItems").ChildObjects(objWbtSearchItems)
		
		For jCount = 0 To objSearchItems.count-1		
			If objSearchItems(jCount).getroproperty("innertext") = strSearchText Then
				objSearchItems(jCount).click	
				flgSearchItem = True
				Exit for
			End If			
		Next
        
        If flgSearchItem Then
            Call fRptWriteReport("Pass", "Verify "&strSearchText& "Clicked","Clicked successfully on "&strSearchText& "field")
            fSearchTextInAribaPage = True
        Else
            Call fRptWriteReport("Fail","Verify Clicked to "&strSearchText,"Failed to Click"&strSearchText)
            Call fRptWriteResultsSummary() 
            Exit Function
        End If
    Else
    	Call fRptWriteReport("Fail","Verify Click on Search Button in Ariba page","Failed to Click on Search Button in Ariba Page")
        Call fRptWriteResultsSummary() 
        Exit Function
    End If  
	Set objHomePage			=	Nothing
	Set objWbtSearchItems	=	Nothing
	Set objSearchItems		=	Nothing	           
    On error goto 0        
End Function

'******************************************************************************************************************************************************************************************************************************************
''    Function Name                        :                fselectRecentManageCreateMenu
''    Objective                            :                Used to select submenu Items under Recent\Manage\Create Ariba Application
''    Input Parameters                     :                strMenuItem,strSubItem
''    Output Parameters                    :                Nil
''    Date Created                         :                29/04/2020
''    UFT/QTP Version                      :                15.0
''    Pre-requisites                       :                NIL  
''    Created By                           :                Cigniti
''    Modification Date                    :       
'***************************************************************************************************************************************************************************************
Public Function fselectRecentManageCreateMenu(strMenuItem,strSubItem)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Variable declaration
	Dim objHomePage,objMenuList,objChildObject,oChildObjectCount,flgSelectionItem
	fselectRecentManageCreateMenu = False ' -05/26/2020 - Updated
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MID_WAIT)
    ' Set Browser and Page 
    Set objHomePage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
    
    'Click on Create\Recent\Manage Dropdowns
    Call fClick(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement").Link("html tag:=A","innertext:="&strMenuItem&" ","name:="&strMenuItem&" "),strMenuItem)
    
    If strMenuItem = "Recent" Then
		strMenuItem = "RecentlyViewed"
	End If
    
    Set objMenuList = objHomePage.SAPWebExtMenu("html tag:=DIV","id:="&strMenuItem)    
   	fSynUntilObjExists objMenuList,5 
   	
   	'Verify whether the Create\Recent\Manage Dropdowns are clicked
    If objMenuList.Exist(1) Then    	    	  
	    Set objChildObject = description.Create
		objChildObject("micclass").value = "link"
		intItemSelectionCount = 0		
		Set oChildObjectCount = objMenuList.ChildObjects(objChildObject)		
		For jCount = 0 To oChildObjectCount.count-1			
			If Trim(oChildObjectCount(jCount).getroproperty("innertext")) = strSubItem Then
				oChildObjectCount(jCount).click	
				flgSelectionItem = True
				Exit for
			End If				
		Next    	
    	If flgSelectionItem Then
            Call fRptWriteReport("Pass", "Verify navigated to "&strMenuItem,"Navigated successfully to "&strMenuItem&"->"&strSubItem)
            fselectRecentManageCreateMenu = True
        Else
            Call fRptWriteReport("Fail","Verify navigated to "&strMenuItem,"Failed to navigate "&strMenuItem&"->"&strSubItem)
            Call fRptWriteResultsSummary() 
            Exit Function
        End If     	
    Else    
    	Call fRptWriteReport("Fail","Verify Clicked on "&strMenuItem,"Failed to Click on "&strMenuItem)
        Call fRptWriteResultsSummary() 
        Exit Function
    End If   		
	'On error goto 0	
	Set objHomePage			=	Nothing
	Set objMenuList			=	Nothing
	Set objChildObject		=	Nothing
	Set oChildObjectCount	=	Nothing
	On error goto 0
End Function


'******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fAribaCreateRequisition
'	Objective							:					Used to Create Requisition for Ariba and click on Continue Shopping
'	Input Parameters					:					objDataDict,iRowCountRef
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					15.0
'	QC Version							:		
'	Pre-requisites						:					NILL  
'	Created By							:					
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fAribaCreateRequisition(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	fAribaCreateRequisition = False
	
	'Clear the testdata sheet
	Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"","TestData","AutoPONumber")
	
	'Fetching Data from the Testdata file
	strLineItems = objDataDict.Item("LineItems" & iRowCountRef)
	
	strTitle = objDataDict.Item("Title" & iRowCountRef)
	strOnBehalfOf = objDataDict.Item("BehalfOf" & iRowCountRef)
	intCompanyCode = objDataDict.Item("CompanyCode" & iRowCountRef)
	strDeliverTo = objDataDict.Item("DeliverTo" & iRowCountRef)
	strComments = objDataDict.Item("Commnets" & iRowCountRef)
	strAttachRequ = objDataDict.Item("AttchmentRequired" & iRowCountRef)
	strApprovalFlowdata = objDataDict.Item("ApprovalFlowtext" & iRowCountRef)	
	strType = objDataDict.Item("TypeName" & iRowCountRef)
	strTypeOfCompany = objDataDict.Item("TypeOfCompany" & iRowCountRef)
	strMenuItem = objDataDict.Item("MenuItem" & iRowCountRef)
	strSubItem = objDataDict.Item("SubItem" & iRowCountRef)	
	strSerachScreen = objDataDict.Item("SearchScreen" & iRowCountRef)	
	
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MID_WAIT)
	'Set the object to the page
	Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
	
	'Select the Menu Items to Create Requisition
	Call fselectRecentManageCreateMenu(strMenuItem,strSubItem)
	
	'Enter value in title	
	fEnterText objPage.WebEdit("txtTitle"),strTitle,"Title"
	
	'Enter value in On Behalf of Field
	fSelect objPage.SAPWebExtMenu("wMnuOnBehalfOf"),"Search more","SearchMore"
	
	'Select the Value on Behalf of
	Call fChooseValueforSearchField(strType,strOnBehalfOf)
	
	'Enter Date in Date Filed
	dtCalenderDate = fGetCalenderDate()
	Wait MIN_WAIT
	fEnterText objPage.SAPWebExtCalendar("wCalCalenderdate"),dtCalenderDate,"Delay Purchase"
	
	'Enter Company Code
	fSelect objPage.SAPWebExtMenu("wMnuCompanyCode"),"Search more","SearchMore"
	
	'Select value for Companycode 
	Call fChooseValueforSearchField(strTypeOfCompany,intCompanyCode)	
	
	'Enter Value in Deliver to
	fEnterText objPage.WebEdit("txtDeliverTo"),strDeliverTo,"DeliverTo"
	
	'Enter Value in comments
	fEnterText objPage.WebEdit("txtComments"),strComments,"Comments"
	
		'Select the checkbox if it is unchecked
		strCheckStatus = objPage.SAPWebExtCheckBox("wChkVisibleSupplier").GetROProperty("state")
			If strCheckStatus = False Then
				fEnterText objPage.SAPWebExtCheckBox("wChkVisibleSupplier"),strchecked,"VisibleCheck"
			End If
	
	'click on ContinueShopping button
	fClick objPage.WebButton("btnContinueShopping"),"ContinueShopping"		
	
	'Verify Catlog Home Page
	If fSynUntilObjExists(objPage.WebElement("weCatalogHome"),MID_WAIT) Then
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Catlog Home Page exist","Catlog Home page exists Successfully")
		fAribaCreateRequisition = True
	Else
		Call fRptWriteReport("Fail", "Verify Catlog Home Page","Unable to load the Catlog Home Screen")	
		Call fRptWriteResultsSummary()        
	    Exit Function
	End If	
		If strLineItems="MULTI" Then
			Call fCreateMultiLineNonCatalogItem(objDataDict,iRowCountRef)		
		Else
			'Create Non_Catelog
			Call fCreateNonCatalogItem(objDataDict,iRowCountRef)
			Wait MIN_WAIT	
		End If	
	
	'View Requisition
	Call fViewRequisition()
	
	'Get the Requisition Number
	strReqID = fAribaCaptureReqID()
	
	'Capture PO Number
	intPONumber = fAribaGeneratePurchaseOrder(objDataDict,iRowCountRef)		
	Wait MIN_WAIT	
	On error goto 0
End Function

'******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fAribaCreateRequisition
'	Objective							:					Used to Create Requisition for Ariba and click on Continue Shopping
'	Input Parameters					:					objDataDict,iRowCountRef
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					15.0
'	QC Version							:		
'	Pre-requisites						:					NILL  
'	Created By							:					
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fCreateNonCatalogItem(objDataDict,iRowCountRef)
 	On error resume next
 	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	fCreateNonCatalogItem = False
 	
 	strType = objDataDict.Item("TypeOfPurchase" & iRowCountRef)
 	strPurchaseOrg = objDataDict.Item("Purchase Org" & iRowCountRef)
' 	strDescription = objDataDict.Item("Description" & iRowCountRef)
 	strCommodityType = objDataDict.Item("CommodityType" & iRowCountRef)
 	strCommodityValue = objDataDict.Item("Commodity Code" & iRowCountRef)
 	strItemCategory = objDataDict.Item("ItemCategory" & iRowCountRef)
 	strCostCenter =  objDataDict.Item("CostCenter" & iRowCountRef)
	strCurrencyType = objDataDict.Item("CurrencyType" & iRowCountRef)
	strCurrencyName = objDataDict.Item("Currency" & iRowCountRef)
	strVendorType = objDataDict.Item("VendorType" & iRowCountRef)
	strVendorName = objDataDict.Item("Vendor" & iRowCountRef)
	strBillType = objDataDict.Item("BillType" & iRowCountRef)
	strBillName = objDataDict.Item("BillName" & iRowCountRef)
	strCostCenter = objDataDict.Item("CostCenter" & iRowCountRef)
	strCostcenterValue = objDataDict.Item("CostCenterValue" & iRowCountRef)
	strShipToType = objDataDict.Item("ShipToType" & iRowCountRef)
	strShipToValue = objDataDict.Item("ShipToValue" & iRowCountRef)
	strText = objDataDict.Item("Text" & iRowCountRef)
	intQuantity = objDataDict.Item("Quantity" & iRowCountRef)
	intPrice = objDataDict.Item("Price" & iRowCountRef)	

	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MID_WAIT) 
	'Set the object to the page
	Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")	
	objPage.Sync
 	Wait MIN_WAIT	 	
	'Click on NonCatolog button
	fClick objPage.WebButton("btnAddNonCatalogItem"),"Add Non-Catolog" 		
	'Verify Create Non-Catolog Screen
 	If fSynUntilObjExists(objPage.WebElement("weCreateNonCatalogItem"),MID_WAIT) Then 			
		'Selec SearchMore in Purch Org: 
		fSelect objPage.SAPWebExtMenu("wMnuPurchasingOrg"),"Search more","SearchMore"
		
		'Select the Value on Purchase Org
		Call fChooseValueforSearchField(strType,strPurchaseOrg)
		
		'Enter Description
		fEnterText objPage.WebEdit("txtDescription"),"test","Description"
		objPage.sync
		
		'Enter Commodity Code
		fSelect objPage.SAPWebExtMenu("wMnuCommidityCode"),"Search more","SearchMore"
		
		'Select the Commodity Type
		Call fChooseValueforSearchField(strCommodityType,strCommodityValue)		
		
		'Enter Item Category
		fSelect objPage.SAPWebExtList("wlstItemCatogery"),strItemCategory,"ItemCategory"
		
 		'Select the Account Type
		 fSelect objPage.SAPWebExtList("wlstCostCenter"),strCostCenter,"CostCenter"	
		 
		 'Enter Quantity
		 fEnterText objPage.WebEdit("txtQuantity"),intQuantity,"Quantity"
		 
		 'Select the CurrencyType
		 fSelect objPage.SAPWebExtMenu("wMnuPriceType"),"Other...","Others"
		 Call fChooseValueforSearchField(strCurrencyType,strCurrencyName)

		'Enter Price
		fEnterText objPage.WebEdit("txtPrice"),intPrice,"Price"		
		
		Wait MIN_WAIT
		'Click on Update button
		fClick objPage.WebButton("btnUpdateAmount"),"Update Amount"	
		 
		'Verify the Updated Amount
		 intAmountValue = intQuantity*intPrice
		 Wait MIN_WAIT
		 Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")	
		 'Get the Value from app
		 intAmount = fWebtableGetCelldata(objPage.WebTable("wbtAmount"),11,3,"AmountFields")
		 arrAmount = Split(intAmount," ")
'		 If strComp(clng(intAmountValue),clng(Right(arrAmount(0),Len(arrAmount(0))-1))) = 0
			If strComp(intAmountValue,clng(arrAmount(0))) = 0 Then
				Call fRptWriteReport("Pass", "Verify Updated Amount","Amount field is updated as "&intAmount)
				fCreateNonCatalogItem = True
			Else
				Call fRptWriteReport("Fail","Verify Updated Amount","Amount is not updated")	
				Call fRptWriteResultsSummary()        
		        Exit Function
			End If
		Wait MIN_WAIT		
		'Enter Vendor under Supplier information
		fSelect objPage.SAPWebExtMenu("wMnuVendor"),"Search more","SearchMore"
		Call fChooseValueforSearchField(strVendorType,strVendorName)
		 
		'Click on Add to cart
		fClick objPage.WebButton("btnAddtoCart"),"Add to Cart"	
		
		'Verify the Proceed to check window appears
		If fSynUntilObjExists(objPage.WebButton("btnProceedCheckout"),MIN_WAIT) Then
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Proceed to Check","Procced to check is verified Successfully")
			'Click on button
			fClick objPage.WebButton("btnProceedCheckout"),"Proceed to Check"	
		Else
			Call fRptWriteReport("Fail", "Verify Proceed to Check","Procced to check is not appeared")		
			Call fRptWriteResultsSummary()        
	        Exit Function
		End If		
		
		'Verify the Cart Summary
		Call fVerifyCartSummary(arrAmount(0))
				
		'Select the Actions under LineItems
		fSelect objPage.SAPWebExtMenu("wMnuActionItem"),"Edit","Edit"
		Wait MIN_WAIT
		'Select BillTo Field
		fSelect objPage.SAPWebExtMenu("wMnuBillTo"),"Search more","SearchMore"
		Call fChooseValueforSearchField(strBillType,strBillName)
		Wait MIN_WAIT
		fSelect objPage.SAPWebExtMenu("wMnuCostCenter"),"Search more","SearchMore"
		Call fChooseValueforSearchField(strCostCenter,strCostcenterValue)
		Wait MIN_WAIT
		'Enter ShipTo
		fSelect objPage.SAPWebExtMenu("wMnuShipTo"),"Search more","SearchMore"
		Call fChooseValueforSearchField(strShipToType,strShipToValue)
		
		'************************************************************************************************
		Call fAribaRequisitionAccountingLineItem(objDataDict,iRowCountRef)	
		'************************************************************************************************
		'Click on OK button
		fClick objPage.WebButton("btnOK"),"OK"			
		Wait MIN_WAIT   		
		'Click on Submit button
		fClick objPage.WebButton("btnSubmit"),"Submit"		
		Wait MIN_WAIT		
		'Verify the Requistion is submitted 
		strRequisitionText = objPage.WebElement("weViewRequisition").GetROProperty("innertext")
		If Instr(strRequisitionText,strText) > 0 Then
			fCreateNonCatalogItem = True
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify the Requsition is submitted","Requisition is submitted Successfully")
		Else
			Call fRptWriteReport("Fail", "Verify the Requsition is submitted","Failed to submit the Requisition")		
			Call fRptWriteResultsSummary()        
	        Exit Function
		End If			
 	End If 	
On error goto 0 	
 End Function

'******************************************************************************************************************************************************************************************************************************************
'   Function Name                         :                    fViewRequisition
'    Objective                            :                    Used to Navigate View Requisition Screen and Deleted Approvers
'    Input Parameters                     :                    objDataDict,iRowCountRef
'    Output Parameters                    :                    NIL
'    Date Created                         :                    
'    UFT Version                          :                    15.0
'    QC Version                           :        
'    Pre-requisites                       :                    NILL  
'    Created By                           :                    
'    Modification Date                    :           
'******************************************************************************************************************************************************************************************************************************************        
Public Function fViewRequisition()

 	 On error resume next 
 	 'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	 fViewRequisition = False
	  Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")   
	  Wait MIN_WAIT	  
	 'Clcik on View Requisition
	 If fSynUntilObjExists(objPage.WebButton("btnViewRequisition"),MIN_WAIT) Then
	     fClick objPage.WebButton("btnViewRequisition"),"View Requisition"    
	     Call fSynUntilObjExists(objPage.WebTable("wbtApprovalFlow"),MID_WAIT)
	     Call fSynUntilObjExists(objPage.WebTable("wbtApprovalFlow"),MID_WAIT)
	     Call fSynUntilObjExists(objPage.WebTable("wbtApprovalFlow"),MID_WAIT)
	     Call fSynUntilObjExists(objPage.WebTable("wbtApprovalFlow"),MID_WAIT)
	     'Verify whehther the page navigated to View Requisition Screen
	     If fSynUntilObjExists(objPage.WebTable("wbtApprovalFlow"),MID_WAIT) Then
	         Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify View Requisition Screen is navigated","Navigated to View Requisition Screen Successfully")                 
	         Wait MIN_WAIT
	         Set objCloseChild = Description.Create
	        objCloseChild("micclass").Value = "WebElement"
	        objCloseChild("class").Value = "a-graph-node-icon w-apv-delete-icon"
	        objCloseChild("html tag").Value = "SPAN"
	        objCloseChild("title").Value = "Delete"
	        objCloseChild("visible").value = True
	        Call fPgDown()
	        Set objCloseApprover = objPage.WebTable("wbtApprovalFlow").ChildObjects(objCloseChild)
	        objCountOfCloseApprover = objCloseApprover.Count                
	        
	        If objCountOfCloseApprover > 0 Then	            
	            Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Approvers are available to Delete","Approvers are available to Delete")	            
	            'Click on X-Symbolol on each Approver
	            For jCount = 0 To objCountOfCloseApprover-1
	                objCloseApprover(jCount).Click	                
	                If fSynUntilObjExists(objPage.WebButton("btnDeleteApprover"),MID_WAIT) Then
	                    fClick objPage.WebButton("btnDeleteApprover"),"Delete Approver --> OK"     
	                    objPage.Sync
				Wait MIN_WAIT	                    
	                    	If objPage.WebElement("weWaitPopUp").Exist(1) Then 
					Wait MIN_WAIT
				End  IF

	                End If					             
	            Next	            
	           If objPage.WebTable("wbtApprovalFlow").Exist(2) Then 
					 Set objCloseApprover = objPage.WebTable("wbtApprovalFlow").ChildObjects(objCloseChild)
			        objCountOfCloseApprover = objCloseApprover.Count                
			        
			        If objCountOfCloseApprover > 0 Then	            
			            Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Approvers are available to Delete","Approvers are available to Delete")	            
			            'Click on X-Symbolol on each Approver
			            For jCount = 0 To objCountOfCloseApprover-1
			                objCloseApprover(jCount).Click	                
			                If fSynUntilObjExists(objPage.WebButton("btnDeleteApprover"),MID_WAIT) Then
			                    fClick objPage.WebButton("btnDeleteApprover"),"Delete Approver --> OK"     
			                    objPage.Sync
								Wait MID_WAIT								
								If objPage.WebElement("weWaitPopUp").Exist(1) Then
									Wait MIN_WAIT
								End  IF
			                End If					             
			            Next
			        End  If
			     End  IF   
	            If fSynUntilObjExists(objPage.WebButton("btnShowApprovalFlow"),MID_WAIT) Then
	                Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Approvers are not available to Delete","Approvers are not available to Delete")
	                fViewRequisition = True
	            Else
	                Call fRptWriteReport("Fail", "Verify Approvers are not available to Delete","Approvers are available to Delete")
	            End If
	            
	        Else
	            Call fRptWriteReport("Fail", "Verify Approvers are available to Delete","Approvers are not available to Delete.. Please verify Test data Once..")
	        End If	  
	     Else
	         Call fRptWriteReport("Fail","Verify View Requisition Screen is navigated","Failed to Navigate to View Requisition Screen")
	     End If
	     
	 End If	 
	'On error goto 0	
	Set objPage = Nothing
	Set objCloseChild = Nothing
	Set objCloseApprover = Nothing     
	On error goto 0
 End Function
'******************************************************************************************************************************************************************************
'	Function Name						:		fVerifyCartSummary
'	Objective							:		Used to Verify Cart Summary
'	Input Parameters					:		strAmount
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fVerifyCartSummary(strAmount)	 
	 On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If			 
	 fVerifyCartSummary = False	 
	 Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)
	 Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")    
	 'Verify Cart Summary
	 If fSynUntilObjExists(objPage.WebElement("weCartSummary"),MID_WAIT) Then		
			'Get the Cart Summary text
			strCartSummary = objPage.WebElement("weCartSummary").GetROProperty("innertext")
			arrCartSum = Split(strCartSummary," ")
			If StrComp(arrCartSum(3),strAmount) = 0 Then
				Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Cart Summary","Cart Summary is displayed as "&strCartSummary)
	                fVerifyCartSummary = True
	         Else
	              Call fRptWriteReport("Fail", "Verify Approvers are not available to Delete","Approvers are available to Delete")
	              Call fRptWriteResultsSummary()        
	        	  Exit Function
	         End If		
	End  If
	On error goto 0
 End Function
   
 '******************************************************************************************************************************************************************************
'	Function Name						:		fAribaCaptureReqID
'	Objective							:		Used to Capture Req ID in Ariba
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fAribaCaptureReqID()
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	fAribaCaptureReqID = False
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement") ,MIN_WAIT) '
	Set obAribajPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")  	
	'Get the Status of Requisition
	Call fSynUntilObjExists(obAribajPage.WebElement("weStatus"),MIN_WAIT)
	strReqStatus = obAribajPage.WebElement("weStatus").GetROProperty("innertext")
		If Instr(strReqStatus,"Approved") > 0 Then
			Call fRptWriteReport("Pass", "Verify Status after Creating Requisition","Status is diaplayed as "&strReqStatus)
		Else
			Call fRptWriteReport("Fail", "Verify Status after Creating Requisition","Status is not diaplayed as "&strReqStatus)	
			Call fRptWriteResultsSummary()        
	        Exit Function
		End If	
	'Get the Requisition ID
	strReqID = obAribajPage.WebElement("weReqID").GetROProperty("innertext")
	arrReqID = Split(strReqID," ")
	strRequisitionNumber = arrReqID(0)	
		If strRequisitionNumber <> "" Then		
			fAribaCaptureReqID = strRequisitionNumber
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Capture the Requisition ID","Requistion is diaplayed as "&strRequisitionNumber)
			Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strRequisitionNumber,"TestData","RequisitionNumber")
		Else
			Call fRptWriteReport("Fail", "Capture the Requisition Number","Unable to capture Req Number")	
			Call fRptWriteResultsSummary()        
	        Exit Function			
		End If
		On error goto 0
End Function
   
 '******************************************************************************************************************************************************************************
'	Function Name						:		fAribaSupplierCreateInvoiceWithPONumber
'	Objective							:		Used to Search PO Number in Ariba Supplier 
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fAribaSupplierCreateInvoiceWithPONumber(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	fCreateInvoiceWithPONumber = False
	
	'Fetching Data from the Testdata file
	strActionType = objDataDict.Item("ActionType" & iRowCountRef)
	
	'Search for PO Number
	Call fAribaSupplierSearchPONumber(objDataDict,iRowCountRef)	
	
	'Select the Order and Select Invoice Type 
	Call fAribaSupplierSelectOrderByActionType(strActionType)	
	
	'Create Invoice details
	strInvoiceNumber = fAribaSupplierCreateInvoice(objDataDict,iRowCountRef) '	 (strTaxPercentage,strFilePath)	
		If strInvoiceNumber <> "" Then
			Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strInvoiceNumber,"TestData","AutoInvoiceNumber")
			Call fRptWriteReport("Pass","Create Invoice Number","Invoice Number :"&strInvoiceNumber&" is generated successfully")
			fAribaSupplierCreateInvoiceWithPONumber = True
		Else
			Call fRptWriteReport("Fail","Create Invoice Number","Invoice Number is not generated")
		End If
	On error goto 0	
End Function 
 
'******************************************************************************************************************************************************************************
'	Function Name						:		fSearchPONumber
'	Objective							:		Used to Search PO Number in Ariba Supplier 
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fAribaSupplierSearchPONumber(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objANSPage
	Dim intPONumber
	
	fAribaSupplierSearchPONumber = False

	'Get the PO Number from testdata sheet
	intPONumber = fGetSingleValue("AutoPONumber","TestData",Environment("TestName")) 
	
	If intPONumber <> "" Then		
		Set objANSPage = Browser("brAribaSpendManagement").Page("pgAribaNetworkSupplier")   
		
		'Enter the PO Number in search text
		fEnterText objANSPage.WebEdit("txtSearchtext"),intPONumber,"PO Number"
		'Click on Search icon
		 fClick objANSPage.WebElement("weSearch"),"Search"  		 
		 
		 If fSynUntilObjExists(objANSPage.WebElement("weSearchFilters"),MID_WAIT) Then
		 	Call fRptWriteReport("PASSWITHSCREENSHOT", "Search PO Number","Search PO Number - "&intPONumber)
	         fAribaSupplierSearchPONumber = True
	     Else
	         Call fRptWriteReport("Fail", "Search PO Number","Failed to search PO Number")
	     End If		
	End If
On error goto 0	
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fSelectOrderByActionType
'	Objective							:		Used to Select the Action At the end of PO line item 
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fAribaSupplierSelectOrderByActionType(strActionType)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	fAribaSupplierSelectOrderByActionType = False
	
	intPONumber = fGetSingleValue("AutoPONumber","TestData",Environment("TestName")) 	
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaNetworkSupplier"),MIN_WAIT)
	Set objANSPage = Browser("brAribaSpendManagement").Page("pgAribaNetworkSupplier") 
	'Get the Row count	
	intRow = objANSPage.WebTable("wbtOrderActionType").RowCount
	For icount = 1 To intRow
		'Get the Cell values from table
		strOrder = objANSPage.WebTable("wbtOrderActionType").GetCellData(icount,3)
		'Validating the Order Number
		If strcomp(strOrder,intPONumber) = 0 Then
			objANSPage.WebTable("wbtOrderActionType").Highlight
			objANSPage.WebTable("wbtOrderActionType").ChildItem(icount,1,"SAPWebExtRadioButton",1).Click
			objANSPage.Link("lnkActions").Click
			fSelect objANSPage.SAPWebExtMenu("wMnuOrderActions"),strActionType,"SelecActionType"
			Call fRptWriteReport("Pass", "Select the Action type in table","Action type is selected as "&strActionType)
			fAribaSupplierSelectOrderByActionType = True
			Exit for	
		End If		
	Next
	If fAribaSupplierSelectOrderByActionType Then
		Call fRptWriteReport("Pass", "Select the Action type in table","Action type is selected as "&strActionType)
	Else
		Call fRptWriteReport("Fail", "Select the Action type in table","Action type is selected as "&strActionType)	
	End If
	On error goto 0
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fCreateInvoice
'	Objective							:		Used to Create Invoice for generated PO Number
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fAribaSupplierCreateInvoice(objDataDict,iRowCountRef)	
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objANSPage
	Dim strTaxPercentage
	Dim strFilePath
	Dim intPONumberCapture
	Dim intNewSupplierQuantity
	
	fAribaSupplierCreateInvoice = False
	'Read data from excel sheet
	strTaxPercentage = objDataDict.Item("TaxPercentage" & iRowCountRef)
	strTaxRateAN = objDataDict.Item("Tax Rate(AN)" & iRowCountRef)
	strFileName = objDataDict.Item("FileName" & iRowCountRef)	
	strTaxCode = objDataDict.Item("Tax Code(Non US/Canada)" & iRowCountRef)
	
	intPONUmber = fGetSingleValue("AutoPONumber","TestData",Environment("TestName")) 	
	intNewSupplierQuantity = fGetSingleValue("NewQuantity","TestData",Environment("TestName")) 
	
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaNetworkSupplier"),MIN_WAIT)
	Set objANSPage = Browser("brAribaSpendManagement").Page("pgAribaNetworkSupplier") 
	objANSPage.Sync
	If fSynUntilObjExists(objANSPage.WebElement("wePONumber"),MID_WAIT) Then
		'Get the PO Number and validate with actual PO Number
		intPONumberCapture = objANSPage.WebElement("wePONumber").GetROProperty("innertext")
		If StrComp(Trim(intPONumberCapture),intPONUmber) = 0 Then
			Call fRptWriteReport("Pass", "Validate the PO number","PO Number is validated successfully")			
		Else	
			Call fRptWriteReport("Fail", "Validate the PO number","Failed to Validate PO Number")	
			Call fRptWriteResultsSummary() 
            Exit Function
		End If
		
		'Enter the Invoice Number
		fEnterText objANSPage.WebEdit("txtInvoiceNumber"),"IN"&intPONumberCapture,"InvoiceNumber"
		
		'Enter Calender Date
		dtDate = fGetCurrentDate()
		fEnterText objANSPage.SAPWebExtCalendar("wCalInvoiceDate"),dtDate,"CalenderDate"
		
		If strTaxCode <> "" Then
			'Enter Supplier VAT/Tax ID
		    fEnterText objANSPage.WebEdit("txtTaxCode"),strTaxCode,"Supplier VAT/TAX ID"	
		End If		
		'Europe - 'Enter Customer VAT/Tax ID
		If fVerifyObjectExist(objANSPage.WebEdit("txtCustomerVATTaxID")) Then
				If strTaxCode <> "" Then
					'Enter Customer VAT/Tax ID
				    Call fEnterText(objANSPage.WebEdit("txtCustomerVATTaxID"),strTaxCode,"Customer VAT/Tax ID")	
				End If
		End If
	
			
		'Select the Checkbox
		 fClick objANSPage.SAPWebExtCheckBox("wChkAgree"),"I Agree"
		 objANSPage.Sync
		 Wait MIN_WAIT
		 'Enter PO Number
		 fEnterText objANSPage.WebEdit("txtPONumber"),intPONumberCapture,"PO Number"
		 
		 'Click on Attachmnets in AddHeader listbox
		 fSelect objANSPage.SAPWebExtMenu("wMnuAddAttachments"),"Attachment","AddToHeader-> Attachment"		
		 Call fAribaSupplierAttachFile(strFileName)
		
		'06/14/2020 - Ratnakar Eda - TAX details are not appeared when user create PO with 121 Comm Code.
		'If Tax Category details are appeared then Select TAX details
		'Check Line item based on Row count - Need to update
		If fVerifyObjectExist(objANSPage.SAPWebExtCheckBox("wChkTaxCategory")) Then
			'Click on Tax checkbox			
			'Call fClick(objANSPage.SAPWebExtCheckBox("wChkTaxCategory"),"TaxCategory")
			'Call fClick(objANSPage.WebElement("weTaxlist"),"TaxList")
			'Select the Tax percentage
			'Call fSelect(objANSPage.SAPWebExtMenu("wMnuTaxPercentage"),strTaxPercentage,"Tax Percentage")
			'Select the line item
			objANSPage.WebTable("wbtInsertLineItem").ChildItem(3,1,"SAPWebExtCheckBox",1).Click
			fClick objANSPage.WebElement("weLineItem"),"LineItem"  
			fClick objANSPage.Link("lnkTax"),"Tax"  
			Wait MIN_WAIT
			'Enter Tax Rate
			If fVerifyObjectExist(objANSPage.WebEdit("txtTAXRate%")) Then
				Call fEnterText(objANSPage.WebEdit("txtTAXRate%"),strTaxRateAN,"Rate(%)")
				Call fShellScriptTabOut()
				wait 3 ' Update TAX Amount
				Call fEnterText(objANSPage.WebEdit("txtTAXRate%"),strTaxRateAN,"Rate(%)")
			End  IF
			
		Else
		
			 'Enter PO Number
			 Call fEnterText(objANSPage.WebEdit("txtPONumber"),intPONumberCapture,"PO Number")
			 objANSPage.WebTable("tblPOItems").ChildItem(2,1,"SAPWebExtCheckBox",0).Click '06/14/2020 - Ratnakar Eda - Need to update based on ROW count 
			' Added code for Commodity Code - 121
			If fClick(objANSPage.WebButton("btnCreate"),"Create") Then
				If fClick(objANSPage.Link("lnkService"),"Service") Then
					Call fSynUntilObjExists(objANSPage.WebElement("weLineItem"),MIN_WAIT)
					If fVerifyObjectExist(objANSPage.WebElement("weLineItem")) Then
						fClick objANSPage.WebElement("weLineItem"),"LineItem"  
						fClick objANSPage.Link("lnkTax"),"Tax"  
						Wait MIN_WAIT
						'Select the Tax percentage
						'Call fSelect(objANSPage.SAPWebExtMenu("wMnuTaxPercentage"),strTaxPercentage,"Tax Percentage")
						'Select TAX Percentage
						'Call fClick(objANSPage.WebEdit("txtTAXPercentage"),"Tax Percentage")
						'Call fSelect(objANSPage.SAPWebExtMenu("wMnuTaxPercentage"),strTaxPercentage,"Tax Percentage")
						'Enter Rate(%) 
						Call fEnterText(objANSPage.WebEdit("txtTAXRatePercentage"),strTaxRateAN,"Rate(%)")
						Call fShellScriptTabOut()

						'Click on Create button
						Call fClick(objANSPage.WebButton("btnPOCreate"),"Create")
						Call fSynUntilObjExists(objANSPage.WebButton("btnNext"),MIN_WAIT)
					Else
						Call fRptWriteReport("Fail", "Select Line Item","Line Item drop down not displayed or not clicked")		
						Call fRptWriteResultsSummary() 
		            	Exit Function
					End If			
				Else
					Call fRptWriteReport("Fail", "Select and Fill Blanket PO Items","Goods link not identified or Clicked")		
					Call fRptWriteResultsSummary() 
	            	Exit Function
				End If
			
			Else
				Call fRptWriteReport("Fail", "Select and Fill Blanket PO Items","Create Button not identified or Clicked")		
				Call fRptWriteResultsSummary() 
            	Exit Function
			End If		
			
		End If

			If Isnumeric(intNewSupplierQuantity) Then
				Call fEnterText(objANSPage.WebEdit("txtLineItemsQuantity"),intNewSupplierQuantity,"Updated Quantity")
				Call fClick(objANSPage.WebButton("btnCreateInvUpdate"),"Update") 
				Call fSynUntilObjExists(objANSPage.WebButton("btnNext"),MIN_WAIT)
			End If
			'Click on Next button
			If fVerifyObjectExist(objANSPage.WebButton("btnNext")) Then		
				Call fClick(objANSPage.WebButton("btnNext"),"Click->Next")
				 Wait MIN_WAIT
			End  IF	 
		 
		 If fSynUntilObjExists(objANSPage.WebTable("wbtStandardInvoice"),MID_WAIT) Then
		 	'Get the Invoice Number
		 	strInvoiceNumber = objANSPage.WebElement("weInvoiceNumber").GetROProperty("innertext")
		 	fAribaSupplierCreateInvoice = strInvoiceNumber
			'Click on Submit button
			fClick objANSPage.WebButton("btnSubmit"),"Submit"
			If fSynUntilObjExists(objANSPage.WebElement("weInvoiceDetail"),MID_WAIT) Then
				Call fRptWriteReport("PassWithScreenshot", "Create and Capture Invoice Number","Invoice is Created and Captured Invoice Number as "&strInvoiceNumber)
			Else
				Call fRptWriteReport("Fail", "Create and Capture Invoice Number","Unable to Create Invoice Number")		
				Call fRptWriteResultsSummary() 
            	Exit Function
			End If		
		 End If		
	End If
	On error goto 0
End Function
'******************************************************************************************************************************************************************************
'	Function Name						:		fAribaSupplierAttachFile
'	Objective							:		Used to Attach a file
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fAribaSupplierAttachFile(strFileName)  
On error resume next
'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaNetworkSupplier"),MIN_WAIT)
    Set oSuppPage = Browser("brAribaSpendManagement").Page("pgAribaNetworkSupplier")
    oSuppPage.Sync
    strFilePath = gstrDocumentFolder&"\"&strFileName
'    strFilePath = "C:\CTAF\Documentation\SampleTest.txt"
    arrFileName = split(strFilePath,"\")
    strFileName = arrFileName(Ubound(arrFileName))
    Wait MIN_WAIT
    'Attach file
    oSuppPage.WebFile("wfAttachFile").Highlight
    oSuppPage.WebFile("wfAttachFile").Submit
    wait MIN_WAIT  
    Window("windowstyle:= 533659648","text:=Ariba Network Supplier.*").Dialog("text:=Open").WinEdit("attached text:=File &name:","windowstyle:=1409286272").Type strFilePath
    Wait MIN_WAIT
    Window("windowstyle:= 533659648","text:=Ariba Network Supplier.*").Dialog("text:=Open").WinButton("regexpwndclass:=Button","regexpwndtitle:=&Open").Click
    Wait MID_WAIT
    'Verify whehter the file was Selected
    strActFileValue = oSuppPage.WebFile("wfAttachFile").GetROProperty("value")    
     If Isnull(strActFileValue) or Isempty(strActFileValue) or Len(strActFileValue) < 1 Then
          oSuppPage.WebFile("wfAttachFile").Highlight
         oSuppPage.WebFile("wfAttachFile").Submit
        wait MIN_WAIT      
        Window("windowstyle:= 533659648","text:=Ariba Network Supplier.*").Dialog("text:=Open").WinEdit("attached text:=File &name:","windowstyle:=1409286272").Type strFilePath
        Wait MIN_WAIT   
        Window("windowstyle:= 533659648","text:=Ariba Network Supplier.*").Dialog("text:=Open").WinButton("regexpwndclass:=Button","regexpwndtitle:=&Open").Click
        Call fSynUntilObjExists(oSuppPage.WebFile("wfAttachFile"),MID_WAIT)
         Wait MIN_WAIT
         'Verify whehter the file was Selected
        strActFileValue = oSuppPage.WebFile("wfAttachFile").GetROProperty("value")
     End  IF
   
    If strActFileValue = strFileName Then       
        oSuppPage.WebButton("class:=w-btn","html tag:=BUTTON","name:=Add Attachment").Click
        Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify file: "&strFileName&" was selectd to attach","File: "&strFileName&" Was Selected to attach Successfully")           
       
        'Verify whether file attached 
        if fSynUntilObjExists(oSuppPage.WebTable("wbtFileAttachment"),MIN_WAIT) Then
            Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify file: "&strFileName&" was attached","File: "&strFileName&" Was attached Successfully")           
        Else
            Call fRptWriteReport("Fail", "Verify file: "&strFileName&" was attached","Failed to attach File: "&strFileName)
            Call fRptWriteResultsSummary()       
            Exit Function
        End If       
    Else
        Call fRptWriteReport("Fail", "Verify file: "&strFileName&" was selectd to attach","File: "&strFileName&" Was not Selected to attach")
        Call fRptWriteResultsSummary()       
        Exit Function
    End If   
On error goto 0    
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fAribaFillAccountingLineItem_UPDATED
'	Objective							:		Used to Fill the account line items
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Public Function fAribaFillAccountingLineItem_UPDATED(objDataDict,iRowCountRef)	
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim objPg
	Dim strAccountAssignment
	Dim intGLAccount
	Dim strProjectWBS
	Dim intCostCEnter
	
	'Read data from excel
	strAccountAssignment = objDataDict.Item("AccountAssignment" & iRowCountRef)
	
	Select Case Ucase(strAccountAssignment)
		
		Case Ucase("P (Project)")
	
				intGLAccount = objDataDict.Item("GLAccount" & iRowCountRef)
				strProjectWBS = objDataDict.Item("Project/WBS" & iRowCountRef)
				intCostCEnter= objDataDict.Item("CostCenterValue" & iRowCountRef)
			
				Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)
				Set objPg=Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
				objPg.Sync
				
				'Select AccountAssignment
				Call fSelect(objPg.SAPWebExtList("lstAccountAssignment"),strAccountAssignment,"Project")
				'Select GL Account
				Call fSelect (objPg.WebElement("weGLAccount"),"Search more","SearchMore")
				Call fClick(objPg.Link("lnkSearchMore"),"SearchMore")
			 	Call fChooseValueforSearchField("General Ledge",intGLAccount)
			 	Call fSynUntilObjExists(objPg.WebElement("weLineItemCostCenter"),MIN_WAIT)
			 	'Select Cost center
			 	Call fSelect (objPg.WebElement("weLineItemCostCenter"),"Cost Center","CostCenter")
				Call fClick(objPg.Link("lnkSearchMore"),"SearchMore")
			 	Call fChooseValueforSearchField("Cost Center",intCostCEnter)
				Call fSynUntilObjExists(objPg.WebElement("weProjectWBS"),MIN_WAIT)
			 	'Select Project/WBS
				Call fSelect (objPg.WebElement("weProjectWBS")," Project/WBS"," Project/WBS")
				Call fClick(objPg.Link("lnkSearchMore"),"SearchMore")
				Call fChooseValueforSearchField("Project/WBS",strProjectWBS)
	
		Case Ucase("A (Asset)")
		
		Case Ucase("K (Cost center)")
		
		Case Ucase("V (No Cost Object(7/8*)")
	
	End Select
	Set objPg = Nothing
On error goto 0
End Function



'*********************************************************************************************************************************
''	Function Name					:				fAribaCreateRequisitionWithSummaryDetails
''	Objective						:				strMenuItem,strSubItem,strTitle,strOnBehalfOf,strCompanyCode,strDeliverTo,strComments,strCheckStatus
''	Input Parameters				:				Nil
''	Output Parameters			    :				Nil
''	Date Created					:				12/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**********************************************************************************************************************************
Public Function fAribaCreateRequisitionWithSummaryDetails(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If		
	Dim objPage
	Dim strMenuItem
	Dim strSubItem
	Dim strTitle
	Dim strOnBehalfOf
	Dim dtCalenderDate
	Dim strCompanyCode
	Dim strDeliverTo
	Dim strComments
	Dim strCheckStatus

	fAribaCreateRequisitionWithSummaryDetails = False
	
	'Fetching Data from the Testdata file
	strTitle = objDataDict.Item("Title" & iRowCountRef)
	strOnBehalfOf = objDataDict.Item("BehalfOf" & iRowCountRef)
	intCompanyCode = objDataDict.Item("Company Code" & iRowCountRef)
	strDeliverTo = objDataDict.Item("DeliverTo" & iRowCountRef)
	strComments = objDataDict.Item("Comments" & iRowCountRef)
	'strAttachRequ = objDataDict.Item("AttchmentRequired" & iRowCountRef)
	strType = objDataDict.Item("TypeName" & iRowCountRef)
	strTypeOfCompany = objDataDict.Item("TypeOfCompany" & iRowCountRef)
	strMenuItem = objDataDict.Item("MenuItem" & iRowCountRef)
	strSubItem = objDataDict.Item("SubItem" & iRowCountRef)	
	
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
	'Set the object to the page
	Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
	objPage.Sync
	'Select the Menu Items to Create Requisition
	Call fselectRecentManageCreateMenu(strMenuItem,strSubItem)
	Call fSynUntilObjExists(objPage.WebEdit("txtTitle"),MIN_WAIT)
	'Enter value in title	
	Call fEnterText (objPage.WebEdit("txtTitle"),strTitle,"Title")
	'Enter value in On Behalf of Field
'	Call fEnterText(objPage.WebEdit("txtOnBehalfOf"),strOnBehalfOf,"On Behalf Of")
	fSelect objPage.SAPWebExtMenu("wMnuOnBehalfOf"),"Search more","SearchMore"	
	Call fChooseValueforSearchField("Name",strOnBehalfOf)
	'Enter Date in Date Filed
	dtCalenderDate = fGetCalenderDate()
	Call fSynUntilObjExists(objPage.SAPWebExtCalendar("wCalCalenderdate"),MIN_WAIT)
	Call fEnterText(objPage.SAPWebExtCalendar("wCalCalenderdate"),dtCalenderDate,"Delay Purchase")
	
	'Enter Company Code
'	Call fEnterText(objPage.WebEdit("txtCompanyCode"),strCompanyCode,"Company Code")
	fSelect objPage.SAPWebExtMenu("wMnuCompanyCode"),"Search more","SearchMore"
	Call fChooseValueforSearchField("CompanyCode",intCompanyCode)	
	
	Call fShellScriptTabOut()
	Call fSynUntilObjExists(objPage.WebEdit("txtDeliverTo"),MIN_WAIT)
	'Enter Value in Deliver to
	Call fEnterText(objPage.WebEdit("txtDeliverTo"),strDeliverTo,"DeliverTo")
	'Enter Value in comments
	Call fEnterText(objPage.WebEdit("txtComments"),strComments,"Comments")
	'Verify Catlog Home Page
		If fSynUntilObjExists(objPage.WebButton("btnContinueShopping"),MID_WAIT) Then
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Requisition summary details","Requisition details are filled in Summary section")
			fAribaCreateRequisitionWithSummaryDetails = True
		Else
			Call fRptWriteReport("Fail", "Fill Requisition summary details","Requisition details are not filled in Summary section")
			Call fRptWriteResultsSummary()        
	        Exit Function
		End If
		'click on ContinueShopping button
		Call fClick (objPage.WebButton("btnContinueShopping"),"ContinueShopping"	)
		If fSynUntilObjExists(objPage.WebElement("weCatalogHome"),MID_WAIT) Then
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Catlog Home Page","Catlog Home page exists Successfully")
			fAribaCreateRequisitionWithSummaryDetails = True
		Else
			Call fRptWriteReport("Fail", "Verify Catlog Home Page","Unable to load the Catlog Home Screen")	
			Call fRptWriteResultsSummary()        
	        Exit Function
		End If
		On error goto 0
End  Function 


'*********************************************************************************************************************************
''	Function Name					:				fAribaSearchAndSelectCatalogAndCheckout
''	Objective						:				Search and Select Catalog item
''	Input Parameters				:				strCatalogItem - Catalog Name
''	Output Parameters			    :				Nil
''	Date Created					:				12/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**********************************************************************************************************************************
Public Function fAribaSearchAndSelectCatalogAndCheckout(objDataDict,iRowCountRef)
    On error resume next
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
        Dim objPage
        Dim strCatalogItem
            
        fAribaSearchAndSelectCatalogAndCheckout = False    
        strCatalogItem = objDataDict.Item("CatalogItemName" & iRowCountRef)
        Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
        Set objPage =  Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
        objPage.Sync
            'Enter Catalog Name
            If fEnterText (objPage.WebEdit("txtCatalogSearch"),strCatalogItem,"Catalog Search") Then
                'Click on Catalog search button
                Call fclick(objPage.WebButton("btnCatalogSearch"),"Catalog Search")
                Call fclick(objPage.WebButton("btnCatalogSearch"),"Catalog Search")
                Call fSynUntilObjExists(objPage.WebButton("btnAddtoCart"),MIN_WAIT)
                If fclick(objPage.WebButton("btnCatalogSearch"),"Catalog Search") Then
                    Call fSynUntilObjExists(objPage.WebButton("btnAddtoCart"),MIN_WAIT)                 
                        'Click on Add to cart
                        If fClick(objPage.WebButton("btnAddtoCart"),"Add to Cart"    ) Then
                            Call fRptWriteReport("Pass", "Add Catalog Item","Catalog Item added Successfully")
                        Else
                            Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
                            fAribaSearchAndSelectCatalogAndCheckout = False
                           Call fRptWriteResultsSummary() 
            				Exit Function                           
                        End If
                Else
                    Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
                    fAribaSearchAndSelectCatalogAndCheckout = False    
                    Call fRptWriteResultsSummary() 
            		Exit Function
                End  IF            
                
            Else
                Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
                fAribaSearchAndSelectCatalogAndCheckout = False        
                Call fRptWriteResultsSummary() 
            	Exit Function
            End  IF
    
            ' If ProceedCheckout button not visable then click on Add to cart again
            If fVerifyObjectExist(objPage.WebButton("btnProceedCheckout")) Then
                '
            Else
                Call fClick(objPage.WebButton("btnAddtoCart"),"Add to Cart"    )
            End  IF
            
            'Verify the Proceed to check window appears
            If fSynUntilObjExists(objPage.WebButton("btnProceedCheckout"),MIN_WAIT) Then
                Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Proceed to Check","Procced to check is Verified Successfully")
                'Click on Proceed To Checkout
                If fClick(objPage.WebButton("btnProceedCheckout"),"Proceed to Check") then
                    Call fSynUntilObjExists(objPage.WebEdit("txtTitle"),MIN_WAIT)
                    Call fRptWriteReport("PASSWITHSCREENSHOT", "Create Requisition","Application Redirected to Create Requisition Page")
                    fAribaSearchAndSelectCatalogAndCheckout = True    
                Else
                    Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
                    fAribaSearchAndSelectCatalogAndCheckout = False    
                    Call fRptWriteResultsSummary() 
            		Exit Function   
                End  IF            
            Else
                Call fRptWriteReport("Fail", "Verify Proceed to Check","Procced to check is not appeared")
                fAribaSearchAndSelectCatalogAndCheckout = False    
                Call fRptWriteResultsSummary() 
            	Exit Function    
            End If        
        Set objPage =  Nothing
    On error goto 0
End Function
		
'*********************************************************************************************************************************
''	Function Name					:				fAribaPerformOperationOnLineItem
''	Objective						:				Selected Operation On Added Line Item
''	Input Parameters				:				strLineItemActionType - Edit / Copy / Delete
''	Output Parameters			    :				Nil
''	Date Created					:				12/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**********************************************************************************************************************************		
Public Function fAribaPerformOperationOnLineItem(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim objPg
	Dim strLineItemActionType

	fAribaPerformOperationOnLineItem = False
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
	Set objPage =  Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
	objPage.Sync
	'Read data from excel
	strLineItemActionType = objDataDict.Item("LineItemActionType" & iRowCountRef)
	
	Select Case Ucase(strLineItemActionType)
		Case Ucase("Edit") ' Select Edit Option
			'Select the Actions under LineItems
			If fSelect(objPage.SAPWebExtMenu("wMnuActionItem"),"Edit","Actions drop down") Then
				Wait MIN_WAIT
				Call fSynUntilObjExists(objPage.SAPWebExtList("lstAccountAssignment"),MIN_WAIT)
				fAribaPerformOperationOnLineItem = True
			Else
				Call fRptWriteReport("Fail","Select value "&Chr(34)&strItem&chr(34)&" in "&chr(34)&strText&Chr(34) ,"Value is not seleted in "&chr(34)&strText&Chr(34))
				Call fRptWriteResultsSummary() 
            	Exit Function
			End  If
		Case Ucase("Copy") ' Select Copy Option
			'TBD
		Case Ucase("Delete") ' Select Delete Option
			'TBD		
	End Select	
	On error goto 0
End Function
'*************************************************************************************************************************************************************
''	Function Name					:				fAribaRequisitionUpdatLineItemDetails
''	Objective						:				Update Line Item Details
''	Input Parameters				:				strCurrencyType,strCurrencyName,strCommodityValue,intQuantity,strPurchaseOrg,strVendorName,intPrice
''	Output Parameters			    :				Nil
''	Date Created					:				19/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************************************		
Public Function fAribaRequisitionUpdatLineItemDetails(objDataDict,iRowCountRef)	
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim objPg
	Dim  strCurrencyType
	Dim  strCurrencyName
	Dim  strCommodityValue
	Dim  intQuantity
 	Dim  strPurchaseOrg
	Dim  strVendorName
	Dim  intPrice 
	
	fAribaRequisitionUpdatLineItemDetails = False
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
	Set objPage=Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
	objPage.Sync
	'Read data from excel
	strCurrencyType = objDataDict.Item("CurrencyType" & iRowCountRef)
	strCurrencyName = objDataDict.Item("CurrencyName" & iRowCountRef)
	strCommodityValue = objDataDict.Item("CommodityValue" & iRowCountRef)
	intQuantity = objDataDict.Item("Quantity" & iRowCountRef)
 	strPurchaseOrg = objDataDict.Item("PurchaseOrg" & iRowCountRef)
 	'strVendorType = objDataDict.Item("VendorType" & iRowCountRef)
	strVendorName = objDataDict.Item("VendorName" & iRowCountRef)
	intPrice = objDataDict.Item("Price" & iRowCountRef)

	'Select the CurrencyType
	Call fSelect(objPage.SAPWebExtMenu("weMnuPriceType"),"Other...","Others")
	Call fChooseValueforSearchField(strCurrencyType,strCurrencyName)
	Call fSynUntilObjExists(objPage.WebEdit("txtQuantity"),MIN_WAIT)	
	'Select the CurrencyType
	If fEnterText(objPage.WebEdit("txtQuantity"),intQuantity,"Quantity") then
		'Enter Price
		If fEnterText(objPage.WebEdit("txtPrice"),intPrice,"Price") then
			''Enter Commodity Code
			If fEnterText(objPage.WebEdit("txtCommidityCode"),strCommodityValue,"Commidity Code") then
				Wait 5
				'Enter Purch Org
				If fEnterText(objPage.WebEdit("txtPurchaseOrg"),strPurchaseOrg,"Purchase Org") then
					'Enter Vendor under Supplier information
					Wait (5)
					If fEnterText(objPage.WebEdit("txtVendor"),strVendorName,"Vendor") then
						Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Line Item Details","Line Item Details are filled")
						fAribaRequisitionUpdatLineItemDetails = True
					Else
						Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
						fAribaRequisitionUpdatLineItemDetails = False			
						Call fRptWriteResultsSummary() 
            			Exit Function
					End  IF	
				Else
					Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
					fAribaRequisitionUpdatLineItemDetails = False			
					Call fRptWriteResultsSummary() 
            		Exit Function
				End  IF	
			Else
				Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))	
				fAribaRequisitionUpdatLineItemDetails = False			
				Call fRptWriteResultsSummary() 
            	Exit Function
			End  IF	
		Else
			Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))	
			fAribaRequisitionUpdatLineItemDetails = False			
			Call fRptWriteResultsSummary() 
            Exit Function
		End  IF	
	Else
		Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
		fAribaRequisitionUpdatLineItemDetails = False			
		Call fRptWriteResultsSummary() 
        Exit Function
	End  IF	
	Set objPg = Nothing
	On error goto 0
End Function	

'*************************************************************************************************************************************************************
''	Function Name					:				fAribaRequisitionAccountingLineItem
''	Objective						:				Add / Update Accounting Line Item Details
''	Input Parameters				:				strAccountAssignment,strBillName,strGLAccountValue,strCostcenterValue
''	Output Parameters			    :				Nil
''	Date Created					:				19/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************************************		
'Public Function fAribaRequisitionAccountingLineItem_OLD(objDataDict,iRowCountRef)	
'	On error resume next
'	'Verify if Step Failed, If yes, it will not run the function
'    If Environment("StepFailed") = "YES" Then
'		Exit Function
'	End If	
'	'Declarations
'	Dim objPg
'	Dim strAccountAssignment
'	Dim strGLAccountValue
'	Dim strProjectWBS
'	Dim intCostCEnter
'	fAribaRequisitionAccountingLineItem = False
'	
'	'Read data from excel
'	strAccountAssignment = objDataDict.Item("AccountAssignment" & iRowCountRef)
'	strBillName = objDataDict.Item("BillName" & iRowCountRef)
'	strGLAccountValue = objDataDict.Item("GLAccount" & iRowCountRef)
'	strProjectWBS = objDataDict.Item("Project/WBS" & iRowCountRef)
'	strCostcenterValue = objDataDict.Item("CostCenterValue" & iRowCountRef)	
'	intAccNumber = objDataDict.Item("AssetNumber" & iRowCountRef)			
'	
''		blnSplitAccFlag = objDataDict.Item("SplitAccountingFlag" & iRowCountRef)	
''		strCostcenterValueTwo = objDataDict.Item("CostCenterValueTwo" & iRowCountRef)	
''		intSplitPercOne = objDataDict.Item("SplitPercentageOne" & iRowCountRef)	
''		intSplitPercTwo = objDataDict.Item("SplitPercentageTwo" & iRowCountRef)	
'		strSplitBy = objDataDict.Item("Split By" & iRowCountRef)	
'		strSplitValue = objDataDict.Item("Split Value" & iRowCountRef) 
'		arrSplitValue = Split(strSplitValue,",")
'	
'	Call fPgDown() ' Page down
'	Select Case Ucase(strAccountAssignment)
'		
'		Case Ucase("P (Project)")
'				Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
'				Set objPg=Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
'				objPg.Sync
'				'Select AccountAssignment
'				If fSelect(objPg.SAPWebExtList("lstAccountAssignment"),strAccountAssignment,"Account Assignment") Then
'					Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
'					'Enter BillTo
'					If fEnterText(objPg.WebEdit("txtAccountingBillTo"),strBillName,"Bill To") Then
'						Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
'						
'						'Enter GL Account
'						If fEnterText(objPg.WebEdit("txtAccountingGLAccount"),strGLAccountValue,"GL Account")Then
'								Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
'								'Cost Center
'								If fEnterText(objPg.WebEdit("txtAccountingCostCenter"),strCostcenterValue,"Cost Center")  Then
'									Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
'									'Select Project WBS
'									If fEnterText(objPg.WebEdit("txtProjectWBS"),strProjectWBS,"Project WBS") Then
'										fAribaRequisitionAccountingLineItem = True	
'										Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Accounting - by Line Item","Accounting - by Line Item details are filled")
'									Else
'										Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
'										fAribaRequisitionAccountingLineItem = False	    
'						        		Exit Function
'					        		End  IF
'								Else
'									Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
'									fAribaRequisitionAccountingLineItem = False	    
'					        		Exit Function
'				        		End  IF
'						Else
'							Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
'							fAribaRequisitionAccountingLineItem = False	    
'			        		Exit Function
'						End IF					
'					Else
'						Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
'						fAribaRequisitionAccountingLineItem = False	    
'		        		Exit Function
'					End  IF
'				Else
'					Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
'					fAribaRequisitionAccountingLineItem = False	    
'	        		Exit Function
'				End  IF			
'		Case Ucase("A (Asset)")		
'				strBillName = objDataDict.Item("BillName" & iRowCountRef)		
'				Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)
'				Set objPg=Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
'				objPg.Sync	    		
'				'Select AccountAssignment
'				objPg.SAPWebExtList("lstAccountAssignment").Highlight
'				Call fSelect(objPg.SAPWebExtList("lstAccountAssignment"),strAccountAssignment,"Account")		
'				
'				'Enter BillTo
'				Call fEnterText(objPg.WebEdit("txtAccountingBillTo"),strBillName,"Bill To")
'				Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
'				
'				'Select Project/WBS
'				Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
'				Call fEnterText (objPg.WebEdit("txtProjectWBS"),strProjectWBS,"Project/WBS")
'				
'				'Select Asset Number
'				Call fSynUntilObjExists(objPg.WebElement("weAccNumber"),MIN_WAIT)
'				Call fClick (objPg.WebElement("weAccNumber"),"Account Number")
'					
'				Call fClick (objPg.Link("lnkSearchmore"),"Search more")
'				Call fChooseValueforSearchField("Asset",intAccNumber) 		
'	
'		Case Ucase("K (Cost center)")
'			
'				Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)
'				Set objPg=Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
'				objPg.Sync
'				'Select AccountAssignment
'				Call fSelect(objPg.SAPWebExtList("lstAccountAssignment"),strAccountAssignment,"Account Assignment")
'				Call fSynUntilObjExists(objPg.WebElement("weLineItemCostCenter"),MIN_WAIT)
'				'Enter BillTo Field
'				If fEnterText(objPg.WebEdit("txtAccountingBillTo"),strBillName,"Bill To") Then					
'					'Enter GL Account
'					If fEnterText(objPg.WebEdit("txtAccountingGLAccount"),strGLAccountValue,"GL Account")Then
'						Call fSynUntilObjExists(objPg.WebEdit("txtAccountingCostCenter"),MIN_WAIT)
'						Call fSynUntilObjExists(objPg.WebEdit("txtAccountingCostCenter"),MIN_WAIT)
'						Call fSynUntilObjExists(objPg.WebEdit("txtAccountingCostCenter"),MIN_WAIT)
'						'Cost Center
'						If strCostcenterValue <> ""  Then
'							Call fEnterText(objPg.WebEdit("txtAccountingCostCenter"),strCostcenterValue,"Cost Center")
'							fAribaRequisitionAccountingLineItem = True
'							Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Accounting - by Line Item","Accounting - by Line Item details are filled")
'						Else
'							Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
'							fAribaRequisitionAccountingLineItem = False	
'							Call fRptWriteResultsSummary()        
'	        				Exit Function
'						End IF						
'						'Perform Split Accounting functionality
'						If strSplitBy <> "" Then
'							arrCostcenterValue = Split(strCostcenterValue,",")
'							Call fAribaSplitAccounting(arrCostcenterValue(1),arrSplitValue(0),arrSplitValue(1))
'						End  If		
'					Else
'						Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
'						fAribaRequisitionAccountingLineItem = False	
'						Call fRptWriteResultsSummary()        
'	        			Exit Function
'					End  IF
'				Else
'					Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
'					fAribaRequisitionAccountingLineItem = False	
'					Call fRptWriteResultsSummary()        
'	        		Exit Function
'				End  IF
'		Case Ucase("V (No Cost Object(7/8*)")
'	
'	End Select
'	
'		Set objPg = Nothing
'On error goto 0
'End Function


Public Function fAribaRequisitionAccountingLineItem(objDataDict,iRowCountRef)	
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim objPg
	Dim strAccountAssignment
	Dim strGLAccountValue
	Dim strProjectWBS
	Dim intCostCEnter
	fAribaRequisitionAccountingLineItem = False
	
	'Read data from excel
	strAccountAssignment = objDataDict.Item("Account Assignment" & iRowCountRef)
	strBillName = objDataDict.Item("Bill To" & iRowCountRef)
	strGLAccountValue = objDataDict.Item("GLAccount" & iRowCountRef)
	strProjectWBS = objDataDict.Item("Project/WBS" & iRowCountRef)
	strCostcenterValue = objDataDict.Item("Cost Center" & iRowCountRef)	
	intAccNumber = objDataDict.Item("Asset Number" & iRowCountRef)			
	
	strSplitBy = objDataDict.Item("Split By" & iRowCountRef)	
	strSplitValue = objDataDict.Item("Split Value" & iRowCountRef) 
	arrSplitValue = Split(strSplitValue,",")
	
	Call fPgDown() ' Page down
	Select Case Ucase(strAccountAssignment)
		
		Case Ucase("P (Project)")
				Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
				Set objPg=Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
				objPg.Sync
				'Select AccountAssignment
				If fSelect(objPg.SAPWebExtList("lstAccountAssignment"),strAccountAssignment,"Account Assignment") Then
					Call fShellScriptTabOut()
'					Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
					'Enter BillTo
					'If fEnterText(objPg.WebEdit("txtAccountingBillTo"),strBillName,"Bill To") Then
						'Call fShellScriptTabOut()
						
						'Select BillTo Field
					If fSelect(objPg.SAPWebExtMenu("wMnuBillTo"),"Search more","SearchMore") Then
						Call fChooseValueforSearchField("Address",strBillName)
'						Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
						
						'Enter GL Account
'						If fEnterText(objPg.WebEdit("txtAccountingGLAccount"),strGLAccountValue,"GL Account")Then
'								Call fShellScriptTabOut()
'								Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
								Call fPgDown_withNumber(2)
								'Select GL Account
							If strGLAccountValue <> "" Then
								Call fSelect(objPg.SAPWebExtMenu("wMnuGLAccount"),"Search more","SearchMore") 
'								Call fClick(objPg.Link("lnkSearchMore"),"SearchMore")
							 	Call fChooseValueforSearchField("General Ledge",strGLAccountValue)
							 	objPg.Sync
'							 	Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
								Call fPgDown_withNumber(2)
								'Cost Center
								If fEnterText(objPg.WebEdit("txtAccountingCostCenter"),strCostcenterValue,"Cost Center")  Then
									Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
									'Select Project WBS
									If fEnterText(objPg.WebEdit("txtProjectWBS"),strProjectWBS,"Project WBS") Then
										Call fShellScriptTabOut()
										fAribaRequisitionAccountingLineItem = True	
										Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Accounting - by Line Item","Accounting - by Line Item details are filled")
									Else
										Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
										fAribaRequisitionAccountingLineItem = False	    
						        		Exit Function
					        		End  IF
								Else
									Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
									fAribaRequisitionAccountingLineItem = False	    
					        		Exit Function
				        		End  IF
						Else
							Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
							fAribaRequisitionAccountingLineItem = False	    
			        		Exit Function
						End IF					
					Else
						Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
						fAribaRequisitionAccountingLineItem = False	    
		        		Exit Function
					End  IF
				Else
					Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
					fAribaRequisitionAccountingLineItem = False	    
	        		Exit Function
				End  IF			
		Case Ucase("A (Asset)")		
'				strBillName = objDataDict.Item("BillName" & iRowCountRef)		
'				Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)
'				Set objPg=Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
'				objPg.Sync	    		
'				'Select AccountAssignment
'				objPg.SAPWebExtList("lstAccountAssignment").Highlight
'				Call fSelect(objPg.SAPWebExtList("lstAccountAssignment"),strAccountAssignment,"Account")		
'				
'				'Enter BillTo
'				Call fEnterText(objPg.WebEdit("txtAccountingBillTo"),strBillName,"Bill To")
'				Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
'				
'				'Select Project/WBS
'				Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
'				Call fEnterText (objPg.WebEdit("txtProjectWBS"),strProjectWBS,"Project/WBS")
'				
'				'Select Asset Number
'				Call fSynUntilObjExists(objPg.WebElement("weAccNumber"),MIN_WAIT)
'				Call fClick (objPg.WebElement("weAccNumber"),"Account Number")
'					
'				Call fClick (objPg.Link("lnkSearchmore"),"Search more")
'				Call fChooseValueforSearchField("Asset",intAccNumber) 	

			Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
				Set objPg=Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
				objPg.Sync
				'Select AccountAssignment
				If fSelect(objPg.SAPWebExtList("lstAccountAssignment"),strAccountAssignment,"Account Assignment") Then
					Call fShellScriptTabOut()
'					Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
					'Enter BillTo
'					If fEnterText(objPg.WebEdit("txtAccountingBillTo"),strBillName,"Bill To") Then
'						Call fShellScriptTabOut()
'						Call fSynUntilObjExists(objPg.WebEdit("txtAssetNumber"),MIN_WAIT)
						'Select BillTo Field
					If fSelect(objPg.SAPWebExtMenu("wMnuBillTo"),"Search more","SearchMore") Then
						Call fChooseValueforSearchField("Address",strBillName)
						Call fSynUntilObjExists(objPg.WebEdit("txtAssetNumber"),MIN_WAIT)
							'Asset Number
							If fEnterText(objPg.WebEdit("txtAssetNumber"),intAccNumber,"Asset Number")  Then	
									Call fShellScriptTabOut()							
									fAribaRequisitionAccountingLineItem = True	
									Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Accounting - by Line Item","Accounting - by Line Item details are filled")
								Else
									Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
									fAribaRequisitionAccountingLineItem = False	    
					        		Exit Function
				        	End  IF
						Else
							Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
							fAribaRequisitionAccountingLineItem = False	    
			        		Exit Function
		        		End  IF
						
					Else
						Call fRptWriteReport("Fail","Requisition Accounting LineItem","Requisition Accounting by Line Item details are not filled")
						fAribaRequisitionAccountingLineItem = False	    
		        		Exit Function
					End  IF	



	
		Case Ucase("K (Cost center)")
			
				Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)
				Set objPg=Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
				objPg.Sync
				'Select AccountAssignment
				Call fSelect(objPg.SAPWebExtList("lstAccountAssignment"),strAccountAssignment,"Account Assignment")
'				Call fSynUntilObjExists(objPg.WebElement("weLineItemCostCenter"),MIN_WAIT)
				'Enter BillTo Field
'				If fEnterText(objPg.WebEdit("txtAccountingBillTo"),strBillName,"Bill To") Then	
'					Call fShellScriptTabOut()	
				If fSelect(objPg.SAPWebExtMenu("wMnuBillTo"),"Search more","SearchMore") Then
						Call fChooseValueforSearchField("Address",strBillName)
						'Call fSynUntilObjExists(objPg.WebEdit("txtProjectWBS"),MIN_WAIT)
					'Enter GL Account
'					If fEnterText(objPg.WebEdit("txtAccountingGLAccount"),strGLAccountValue,"GL Account")Then
'						Call fShellScriptTabOut()	
'						Call fSynUntilObjExists(objPg.WebEdit("txtAccountingCostCenter"),MIN_WAIT)
						Call fPgDown_withNumber(2)
						Wait(MIN_WAIT)
					'Select GL Account
					If fSelect(objPg.SAPWebExtMenu("wMnuGLAccount"),"Search more","SearchMore") Then
					 	Call fChooseValueforSearchField("General Ledge",strGLAccountValue)
					 	objPg.Sync
'					 	Call fSynUntilObjExists(objPg.WebEdit("txtAccountingCostCenter"),MIN_WAIT)
						'Cost Center
						If strCostcenterValue <> ""  Then
							If Instr(strCostcenterValue,",") Then
								strCostValue = Trim(Split(strCostcenterValue,",")(0))
							Else
								strCostValue = strCostcenterValue
							End If
'							Call fPgDown_withNumber(2)
							Wait(MIN_WAIT)
'							Call fEnterText(objPg.WebEdit("txtAccountingCostCenter"),strCostValue,"Cost Center")
							fSelect objPg.SAPWebExtMenu("wMnuCostCenter"),"Search more","SearchMore"
							Call fChooseValueforSearchField("Cost Center",strCostValue)
'							Call fShellScriptTabOut()	
							fAribaRequisitionAccountingLineItem = True
							Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Accounting - by Line Item","Accounting - by Line Item details are filled")
							
						Else
							Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
							fAribaRequisitionAccountingLineItem = False	
							Call fRptWriteResultsSummary()        
	        				Exit Function
						End IF						
						'Perform Split Accounting functionality
						If strSplitBy <> "" Then						
							arrCostcenterValue = Split(strCostcenterValue,",")
							Call fAribaSplitAccounting(Trim(arrCostcenterValue(1)),Trim(arrSplitValue(0)),Trim(arrSplitValue(1)))
						End  If			
					Else
						Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
						fAribaRequisitionAccountingLineItem = False	
						Call fRptWriteResultsSummary()        
	        			Exit Function
					End  IF
				Else
					Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
					fAribaRequisitionAccountingLineItem = False	
					Call fRptWriteResultsSummary()        
	        		Exit Function
				End  IF
		Case Ucase("V (No Cost Object(7/8*)")
	
	End Select
	
		Set objPg = Nothing
On error goto 0
End Function


'*************************************************************************************************************************************************************
''	Function Name					:				fAribaRequisitionShippingByLineItem
''	Objective						:				Add / Update Shipping By Line Item data
''	Input Parameters				:				strShipToValue
''	Output Parameters			    :				Nil
''	Date Created					:				19/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************************************	
Public Function fAribaRequisitionShippingByLineItem(objDataDict,iRowCountRef)
 	On error resume next
 	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	'Dim strShipToType
 	Dim objPg
 	Dim strShipToValue
 	Dim strPurchaseGroup
 

 	fAribaRequisitionShippingByLineItem = False
 	
 	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
 	Set objPg = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
 	objPg.Sync
 	Call fPgDown() ' Page down
 	'Read data from excel 
	strShipToValue = objDataDict.Item("Ship To" & iRowCountRef)
	strPurchaseGroup = objDataDict.Item("Purchase Group" & iRowCountRef)
		'Enter Ship To
		'If fEnterText(objPg.WebEdit("txtShipTo"),strShipToValue,"Ship To")  Then
		'	Call fShellScriptTabOut()
			
			'Enter ShipTo
		If fSelect(objPg.SAPWebExtMenu("wMnuShipTo"),"Search more","SearchMore") Then
			Call fChooseValueforSearchField("Plant",strShipToValue)
			objPg.Sync
			Call fShellScriptTabOut()
			
			If fEnterText(objPg.WebEdit("txtPurchaseGroup"),strPurchaseGroup,"Purchase Group") Then
				Call fShellScriptTabOut()
				fAribaRequisitionShippingByLineItem = True
				Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Shipping - by Line Item","Shipping - by Line Item details are filled")
			Else
				Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
				fAribaRequisitionShippingByLineItem = False	
				Call fRptWriteResultsSummary()        
		        Exit Function
			End IF			
		Else
			Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
			fAribaRequisitionShippingByLineItem = False	
			Call fRptWriteResultsSummary()        
	        Exit Function
		End IF			
	Set objPg = Nothing
 On error goto 0
 End Function
'*************************************************************************************************************************************************************
''	Function Name					:				fAribaSubmitRequisition
''	Objective						:				Submit New Requisition 
''	Input Parameters				:				Nil
''	Output Parameters			    :				Nil
''	Date Created					:				19/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************************************	
Public Function fAribaSubmitRequisition(objDataDict,iRowCountRef)	
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	Dim objPage
 	Dim dtValidityStartDate
 	Dim dtValidityEndDate
 	
 	dtValidityStartDate = objDataDict.Item("Validity Start Date" & iRowCountRef)
 	dtValidityEndDate =objDataDict.Item("Validity End Date" & iRowCountRef) 	
	
	fAribaSubmitRequisition = False
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
	Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
		'Click on OK button
		If fClick(objPage.WebButton("btnOK"),"OK") Then
			Call fSynUntilObjExists(objPage.WebButton("btnSubmit"),MIN_WAIT) 
			'Enter Valid date
			If dtValidityStartDate <> "" and dtValidityEndDate <>"" and IsEmpty(dtValidityStartDate) = FALSE Then
				dtValidityStartDate = fGetFutureDateAdd(dtValidityStartDate)
				dtValidityEndDate = fGetFutureDateAdd(dtValidityEndDate)
					If fEnterText(objPage.WebEdit("txtValidityStartDate"),dtValidityStartDate,"Validity Start Date") Then
							If fEnterText(objPage.WebEdit("txtValidityEndDate"),dtValidityEndDate,"Validity End Date") Then
								'TBD
							Else
								fAribaSubmitRequisition = False
								Call fRptWriteReport("Fail","Enter Validity End Date","Date not entered in Validity End Date field")
								Call fRptWriteResultsSummary() 
								Exit Function
							End  IF	
					Else
						fAribaSubmitRequisition = False
						Call fRptWriteReport("Fail","Enter Validity Start Date","Date not entered in Validity Start Date field")
						Call fRptWriteResultsSummary() 
						Exit Function
					End If
			End  IF		

				'Click on Submit button
				If fClick(objPage.WebButton("btnSubmit"),"Submit") Then
					Call fSynUntilObjExists(objPage.WebElement("weViewRequisition"),MIN_WAIT)
					Call fSynUntilObjExists(objPage.WebElement("weViewRequisition"),MIN_WAIT)
					fAribaSubmitRequisition = TRUE	
				Else
					Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
					fAribaSubmitRequisition = False	
					Call fRptWriteResultsSummary() 
            	    Exit Function
				End If			
		Else
			Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
			fAribaSubmitRequisition = False	
			Call fRptWriteResultsSummary() 
            Exit Function
		End  IF
		'Verify the Requistion is submitted 
		If fVerifyObjectExist(objPage.WebElement("weViewRequisition")) Then
			strRequisitionText = objPage.WebElement("weViewRequisition").GetROProperty("innertext")
			If Instr(Ucase(strRequisitionText),Ucase("The requisition has been submitted")) > 0 Then
				fAribaSubmitRequisition = True
				Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify the Requsition is submitted","Requisition is submitted Successfully")
			Else
				Call fRptWriteReport("Fail", "Verify the Requsition is submitted","Failed to submit the Requisition")	
				fAribaSubmitRequisition = False	
				Call fRptWriteResultsSummary() 
	            Exit Function				
			End If	
		Else
			Call fRptWriteReport("Fail", "View Requsition","View Requisition not displayed")
			Exit Function
		End If	
	Set objPg = Nothing	
On error goto 0	
End Function
'*************************************************************************************************************************************************************
''	Function Name					:				fFioriSearchOrderStatus
''	Objective						:				Submit New Requisition 
''	Input Parameters				:				Nil
''	Output Parameters			    :				Nil
''	Date Created					:				19/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************************************	 
Public Function fFioriSearchOrderStatus(objDataDict,iRowCountRef)
  	On error resume next
  	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	Dim objPage
 	Dim objPageHome
	Dim intOrderNumberColNo
	Dim intOrderStatusColNo
	Dim intOrderNumber
	Dim StrActualOrderStatus
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
	Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
	objPage.Sync
	Set objPageHome = Browser("brAribaSpendManagement").Page("pgAribaNetworkSupplier")
	fFioriSearchOrderStatus = False
	'Read data from excel 
	intAutoPONumber = fGetSingleValue("AutoPONumber","TestData",Environment("TestName")) 
	strExpOrderStatus = objDataDict.Item("ExpOrderStatus" & iRowCountRef) 
	'Click Exit
	If fClick(objPageHome.Link("lnkExit"),"Exit") then
		Call fSynUntilObjExists(objPage.Link("lnkSupplierInbox"),MIN_WAIT) 
		'Click Inbox
		If fClick(objPage.Link("lnkSupplierInbox"),"Inbox") Then
			Call fSynUntilObjExists(objPage.Link("lnkSuppOrdersAndReleases"),MIN_WAIT) 
			If fClick(objPage.Link("lnkSuppOrdersAndReleases"),"Orders and Releases") Then
				Call fSynUntilObjExists(objPage.WebTable("tblOrdersAndReleases"),MIN_WAIT)  
				fFioriSearchOrderStatus = True	
			Else
				Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
				fFioriSearchOrderStatus = False	
				Call fRptWriteResultsSummary() 
            	Exit Function
			End If
		Else
			Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
			fFioriSearchOrderStatus = False	
			Call fRptWriteResultsSummary() 
            Exit Function
		End If
	Else
		Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
		fFioriSearchOrderStatus = False	
		Call fRptWriteResultsSummary() 
        Exit Function
	End IF 

	'Get Column number based on Column name
	If fFioriSearchOrderStatus = True	 Then
		intOrderNumberColNo = fGetTableHeaderColumnNumber(objPage.WebTable("tblOrdersAndReleases"),1,1,"Order Number")
		intOrderNumber = fGetRowNumberInTableBasedonColumnData (objPage.WebTable("tblOrdersAndReleases"),intOrderNumberColNo,intAutoPONumber)	
		intOrderStatusColNo = fGetTableHeaderColumnNumber(objPage.WebTable("tblOrdersAndReleases"),1,1,"Order Status")
		If intOrderNumber > 0 Then
			StrActualOrderStatus = fWebtableGetCelldata(objPage.WebTable("tblOrdersAndReleases"),intOrderNumber,intOrderStatusColNo,"Orders and Releases")
			If Ucase(strExpOrderStatus)= Ucase(StrActualOrderStatus) Then
				Call fRptWriteReport("PASSWITHSCREENSHOT","Verify Order Status",strExpOrderStatus &"- Status displayed for created order") 															
			Else
				Call fRptWriteReport("Fail","Verify Order Status",strExpOrderStatus &"- Status not displayed for newly created order") 																			
				Call fRptWriteResultsSummary()        
	        	Exit Function
			End If
		Else
			Call fRptWriteReport("Fail","Verify Order Number",intAutoPONumber &"- Order number not displayed in table") 	
			Call fRptWriteResultsSummary()        
	        Exit Function			
		End If
	End If
		Set objPage = Nothing	
		Set objPageHome = Nothing
On error goto 0
 End Function
 
 '*************************************************************************************************************************************************************
''    Function Name                   :                fAribaCreateRequsitionAsNonCatalogItem
''    Objective                       :                Create Non-Catalog item
''    Input Parameters                :                strPurchaseOrg,strDescription,strCommodityValue,strItemCategory,strCostCenter,strCurrencyType,strCurrencyName
                                                       'strVendorName,intQuantity,intPrice
''    Output Parameters               :                Nil
''    Date Created                    :                25/May/2020
''    UFT/QTP Version                 :                15.0
''    Pre-requisites                  :                NIL  
''    Created By                      :                Cigniti
''    Modification Date               :                06/08/2020 - Modified based on PTP Test data sheet   
'**************************************************************************************************************************************************************    
Public Function fAribaCreateRequsitionAsNonCatalogItem_OLDFUNCTION(objDataDict,iRowCountRef)
     On error resume next
     'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
     fAribaCreateRequsitionAsNonCatalogItem = False

    strPurchaseOrg = objDataDict.Item("PurchaseOrg" & iRowCountRef)
     strDescription = objDataDict.Item("PurchaseOrg" & iRowCountRef)
     strCommodityValue = objDataDict.Item("CommodityValue" & iRowCountRef)
     strItemCategory = objDataDict.Item("ItemCategory" & iRowCountRef)
     strCostCenter =  objDataDict.Item("CostCenter" & iRowCountRef)
    strCurrencyType = objDataDict.Item("CurrencyType" & iRowCountRef)
    strCurrencyName = objDataDict.Item("CurrencyName" & iRowCountRef)
    strVendorName = objDataDict.Item("VendorName" & iRowCountRef)
    intQuantity = objDataDict.Item("Quantity" & iRowCountRef)
    intPrice = objDataDict.Item("Price" & iRowCountRef)
	
    'Set the object to the page
    Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
    Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")   
	Call fSynUntilObjExists(objPage.WebButton("btnAddNonCatalogItem"),MID_WAIT)   
	wait (2)
    'Click on NonCatolog button
    If fClick(objPage.WebButton("btnAddNonCatalogItem"),"Add Non-Catolog") Then
    	Call fSynUntilObjExists(objPage.WebElement("weCreateNonCatalogItem"),MIN_WAIT) 
    	If objPage.WebButton("btnAddNonCatalogItem").Exist(1) Then
    		 Call fClick(objPage.WebButton("btnAddNonCatalogItem"),"Add Non-Catolog")
    		 'Verify Create Non-Catolog Screen
	         Call fSynUntilObjExists(objPage.WebElement("weCreateNonCatalogItem"),MIN_WAIT) 
	         Wait MIN_WAIT
    	End If  
         'Enter PurchaseOrg
         If fEnterText(objPage.WebEdit("txtPurchaseOrg"),strPurchaseOrg,"Purchase Org")  Then
             'Description
             If fEnterText(objPage.WebEdit("txtDescription"),strDescription,"Description") Then
                 'Commidity Code
                 If fEnterText(objPage.WebEdit("txtCommidityCode"),strCommodityValue,"Commodity Code") Then
                     Call fSynUntilObjExists(objPage.SAPWebExtList("wlstItemCatogery"),MID_WAIT) 
                     Call fSynUntilObjExists(objPage.SAPWebExtList("wlstItemCatogery"),MID_WAIT) 
                     Call fSynUntilObjExists(objPage.SAPWebExtList("wlstItemCatogery"),MID_WAIT) 
                     ' Item Code
                     If fSelect(objPage.SAPWebExtList("wlstItemCatogery"),strItemCategory,"ItemCategory") then 
                         'Account Type
                         If fSelect(objPage.SAPWebExtList("wlstCostCenter"),strCostCenter,"Cost Center") then 
                             'Enter Quantity
                            If fEnterText(objPage.WebEdit("txtQuantity"),intQuantity,"Quantity") then'                           
						         'Select the CurrencyType
						         Call fSelect(objPage.SAPWebExtMenu("wMnuPriceType"),"Other...","Others")
						         Call fChooseValueforSearchField(strCurrencyType,strCurrencyName)
'                               'Enter Price
                                If fEnterText(objPage.WebEdit("txtPrice"),intPrice,"Price") then
                                    fAribaCreateRequsitionAsNonCatalogItem = True        
                                Else
                                    Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
                                    fAribaCreateRequsitionAsNonCatalogItem = False            
                                    Call fRptWriteResultsSummary() 
            						Exit Function
                                End  IF    
                            Else
                                Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
                                fAribaCreateRequsitionAsNonCatalogItem = False            
                                Call fRptWriteResultsSummary() 
            					Exit Function
                            End  IF
                         Else
                             Call fRptWriteReport("Fail", "Select Item from Account Type",strCostCenter &"- item not selected from list")
                              fAribaCreateRequsitionAsNonCatalogItem = False            
                             Call fRptWriteResultsSummary() 
            				 Exit Function
                         End  IF
                     
                     Else
                          Call fRptWriteReport("Fail", "Select Item from Item Category ",strItemCategory &"- item not selected from list")
                          fAribaCreateRequsitionAsNonCatalogItem = False            
                       	  Call fRptWriteResultsSummary() 
            			  Exit Function
                     End  IF
                 Else
                     Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
                    fAribaCreateRequsitionAsNonCatalogItem = False            
                    Call fRptWriteResultsSummary() 
            		Exit Function
                 End If
             Else
                 Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
                fAribaCreateRequsitionAsNonCatalogItem = False            
                Call fRptWriteResultsSummary() 
            	Exit Function
             End If
         Else
             Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
            fAribaCreateRequsitionAsNonCatalogItem = False            
            Call fRptWriteResultsSummary() 
            Exit Function
         End  IF
    Else
        Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
        fAribaCreateRequsitionAsNonCatalogItem = False    
        Call fRptWriteResultsSummary() 
        Exit Function
    End IF
    
        If fAribaCreateRequsitionAsNonCatalogItem = True then
         'click on Update amount
             If fClick(objPage.WebButton("btnUpdateAmount"),"Update Amount") Then 
                  intAmountValue = intQuantity*intPrice
                 Wait MIN_WAIT
                 Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")    
                 'Get the Value from app
                 intAmount = fWebtableGetCelldata(objPage.WebTable("wbtAmount"),11,3,"AmountFields")
                 arrAmount = Split(intAmount," ")
				If strComp(clng(intAmountValue),clng(Right(arrAmount(0),Len(arrAmount(0))-1))) = 0 Then
                    Call fRptWriteReport("Pass", "Verify Updated Amount","Amount field is updated as "&intAmount)
                    If fEnterText(objPage.WebEdit("txtVendor"),strVendorName,"Vendor") then
                        Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Catalog Details","catalog details are filled")
                        If fClick(objPage.WebButton("btnAddtoCart"),"Add to Cart")     Then  
                            Call fSynUntilObjExists(objPage.WebButton("btnProceedCheckout"),MID_WAIT)
                            fAribaCreateRequsitionAsNonCatalogItem = True
                            Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Catalog Details","catalog details are filled")
                        Else
                            Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
                            fAribaCreateRequsitionAsNonCatalogItem = False    
                           Call fRptWriteResultsSummary() 
            				Exit Function
                        End If
                        
                    Else
                        Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
                        fAribaCreateRequsitionAsNonCatalogItem = False            
                        Call fRptWriteResultsSummary() 
            			Exit Function    
                    End  IF    
                Else
                    Call fRptWriteReport("Fail","Verify Updated Amount","Amount is not updated")
                    fAribaCreateRequsitionAsNonCatalogItem = False    
                    Call fRptWriteResultsSummary() 
            		Exit Function    
                End If
             Else
                 Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
                fAribaCreateRequsitionAsNonCatalogItem = False    
                Call fRptWriteResultsSummary() 
            	Exit Function
             End  If
        End  IF 
 
        'Verify the Proceed to check window appears
        If fSynUntilObjExists(objPage.WebButton("btnProceedCheckout"),MIN_WAIT) Then
            Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Proceed to Check","Procced to check is verified Successfully")
            'Click on button
            fClick objPage.WebButton("btnProceedCheckout"),"Proceed to Check"    
        Else
            Call fRptWriteReport("Fail", "Verify Proceed to Check","Procced to check is not appeared")    
            fAribaCreateRequsitionAsNonCatalogItem = False    
            Call fRptWriteResultsSummary() 
            Exit Function         
        End If                
    Set objPage = Nothing  
On error goto 0    
 End Function
 
 Public Function fAribaCreateRequsitionAsNonCatalogItem(objDataDict,iRowCountRef)
     On error resume next
     'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
     fAribaCreateRequsitionAsNonCatalogItem = False

	strPurchaseOrg = objDataDict.Item("Purchase Org" & iRowCountRef)
	strDescription = objDataDict.Item("Full Description" & iRowCountRef)
	strCommodityValue = objDataDict.Item("CommodityValue" & iRowCountRef)
	'strItemCategory = objDataDict.Item("ItemCategory" & iRowCountRef)
	strCostCenter =  objDataDict.Item("CostCenter" & iRowCountRef)
	strCurrencyType = objDataDict.Item("CurrencyType" & iRowCountRef)
	strCurrencyName = objDataDict.Item("Currency" & iRowCountRef)
	strVendorName = objDataDict.Item("Vendor" & iRowCountRef)
	intQuantity = objDataDict.Item("Quantity" & iRowCountRef)
	intPrice = objDataDict.Item("Price" & iRowCountRef)
	strItemCategory = objDataDict.Item("Commodity Code" & iRowCountRef)
	intMaxAmount = objDataDict.Item("Max Amount" & iRowCountRef)
	intExpectedAmount = objDataDict.Item("Expected Amount" & iRowCountRef)
	strMaxAmountCurrency = objDataDict.Item("Max Amount Currency" & iRowCountRef)
	strExpectedAmountCurrency = objDataDict.Item("Expected Amount Currency" & iRowCountRef)
	strSupplierChoice = objDataDict.Item("Supplier Choice" & iRowCountRef)
	
    'Set the object to the page
    Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
    Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")   
	Call fSynUntilObjExists(objPage.WebButton("btnAddNonCatalogItem"),MID_WAIT)   
	wait (2)
    'Click on NonCatolog button
    If fClick(objPage.WebButton("btnAddNonCatalogItem"),"Add Non-Catolog") Then
    	Call fSynUntilObjExists(objPage.WebElement("weCreateNonCatalogItem"),MIN_WAIT) 
    	If objPage.WebButton("btnAddNonCatalogItem").Exist(1) Then
    		 Call fClick(objPage.WebButton("btnAddNonCatalogItem"),"Add Non-Catolog")
    		 'Verify Create Non-Catolog Screen
	         Call fSynUntilObjExists(objPage.WebElement("weCreateNonCatalogItem"),MIN_WAIT) 
	         Wait MIN_WAIT
    	End If  
         'Enter PurchaseOrg
         If strPurchaseOrg <> "" Then
'        	 fEnterText(objPage.WebEdit("txtPurchaseOrg"),strPurchaseOrg,"Purchase Org")  
			'Selec SearchMore in Purch Org: 
			fSelect objPage.SAPWebExtMenu("wMnuPurchasingOrg"),"Search more","SearchMore"		
			Call fChooseValueforSearchField("Purchase Org.",strPurchaseOrg)	
'         	Call fShellScriptTabOut()
             'Description
             If fEnterText(objPage.WebEdit("txtDescription"),strDescription,"Full Description") Then
                 'Commidity Code
'                 If fEnterText(objPage.WebEdit("txtCommidityCode"),strCommodityValue,"Commodity Code") Then
'                     Call fSynUntilObjExists(objPage.SAPWebExtList("wlstItemCatogery"),MID_WAIT) 
                     ' Item Code
'                     If fSelect(objPage.SAPWebExtList("wlstItemCatogery"),strItemCategory,"ItemCategory") then 
'                         'Account Type
'                         If fSelect(objPage.SAPWebExtList("wlstCostCenter"),strCostCenter,"Cost Center") then 
							  'Enter Material Group 
                            If strItemCategory <> "" then
'                            	fEnterText(objPage.WebEdit("txtMaterialGroup"),strItemCategory,"Material Group")
'                            	Call fShellScriptTabOut()
								'Enter Commodity Code
								fSelect objPage.SAPWebExtMenu("wMnuMaterialGroup"),"Search more","SearchMore"								
								'Select the Commodity Type
								Call fChooseValueforSearchField("Material Group",strItemCategory)		
								
							
								'Enter Max Amount and Expected amount
                            	If intMaxAmount <> "" and Trim(intMaxAmount) <> "Empty" Then ' Need to modify based on framework format
                            		Wait (3)
                            	 	Call fSelect(objPage.SAPWebExtMenu("wMnuMaxAmount"),"Other...","Others")
								    Call fChooseValueforSearchField(strCurrencyType,strMaxAmountCurrency)
								    Wait(3)
                            		Call fEnterText(objPage.WebEdit("txtMaxAmount"),intMaxAmount,"Max Amount") ' Enter Max Amount
                            		Call fSelect(objPage.SAPWebExtMenu("wMnuExpectedAmount"),"Other...","Others")
								    Call fChooseValueforSearchField(strCurrencyType,strExpectedAmountCurrency)
									Wait(3)
									Call fEnterText(objPage.WebEdit("txtExpectedAmount"),intExpectedAmount,"Expected Currency Amount") ' Enter Expected Currency Amount
									Call fSelect(objPage.SAPWebExtList("lstSupplierChoice"),strSupplierChoice,"Supplier Choice") ' Select Supplier choice
									fAribaCreateRequsitionAsNonCatalogItem = True
								Else
									'Enter Quantity
								   Call fSynUntilObjExists(objPage.SAPWebExtMenu("wMnuPriceType"),MIN_WAIT) 

		                           If fEnterText(objPage.WebEdit("txtQuantity"),intQuantity,"Quantity") then'                           
								         'Select the CurrencyType
								         Call fSelect(objPage.SAPWebExtMenu("wMnuPriceType"),"Other...","Others")
								         Call fChooseValueforSearchField(strCurrencyType,strCurrencyName)
			'                               'Enter Price
											Wait MIN_WAIT	
			                                If fEnterText(objPage.WebEdit("txtPrice"),intPrice,"Price") then
			                                    fAribaCreateRequsitionAsNonCatalogItem = True        
			                                Else

			                                    Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
			                                    fAribaCreateRequsitionAsNonCatalogItem = False            
			                                    Call fRptWriteResultsSummary() 
			            						Exit Function
			                             	End  IF    

	                            Else
		                                Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
		                                fAribaCreateRequsitionAsNonCatalogItem = False            
		                                Call fRptWriteResultsSummary() 
		            					Exit Function
		                            End  IF
                            	End If
                            Else
                            	 Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
                                fAribaCreateRequsitionAsNonCatalogItem = False            
                                Call fRptWriteResultsSummary() 
            					Exit Function
                            End  IF
                         Else
                             Call fRptWriteReport("Fail", "Select Item from Account Type",strCostCenter &"- item not selected from list")
                              fAribaCreateRequsitionAsNonCatalogItem = False            
                             Call fRptWriteResultsSummary() 
            				 Exit Function
                         End  IF
                     
'                     Else
'                          Call fRptWriteReport("Fail", "Select Item from Item Category ",strItemCategory &"- item not selected from list")
'                          fAribaCreateRequsitionAsNonCatalogItem = False            
'                       	  Call fRptWriteResultsSummary() 
'            			  Exit Function
'                     End  IF
'                 Else
'                     Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
'                    fAribaCreateRequsitionAsNonCatalogItem = False            
'                    Call fRptWriteResultsSummary() 
'            		Exit Function
'                 End If
'             Else
'                 Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
'                fAribaCreateRequsitionAsNonCatalogItem = False            
'                Call fRptWriteResultsSummary() 
'            	Exit Function
'             End If
         Else
             Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
            fAribaCreateRequsitionAsNonCatalogItem = False            
            Call fRptWriteResultsSummary() 
            Exit Function
         End  IF
    Else
        Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
        fAribaCreateRequsitionAsNonCatalogItem = False    
        Call fRptWriteResultsSummary() 
        Exit Function
    End IF
    
        If fAribaCreateRequsitionAsNonCatalogItem = True then
         'click on Update amount
             If fClick(objPage.WebButton("btnUpdateAmount"),"Update Amount") Then 
                  intAmountValue = intQuantity*intPrice
                 Wait MIN_WAIT
                 Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")    
                 'Get the Value from app
                 intAmount = fWebtableGetCelldata(objPage.WebTable("wbtAmount"),11,3,"AmountFields")
                 arrAmount = Split(intAmount," ")
				If strComp(clng(intAmountValue),clng(Right(arrAmount(0),Len(arrAmount(0))-1))) = 0 Then
                    Call fRptWriteReport("Pass", "Verify Updated Amount","Amount field is updated as "&intAmount)
                    
                    'Enter Vendor under Supplier information
					If strVendorName <> "" Then
						Call fSelect(objPage.SAPWebExtMenu("wMnuVendor"),"Search more","SearchMore")
						Call fChooseValueforSearchField("ID",strVendorName)
                    'If fEnterText(objPage.WebEdit("txtVendor"),strVendorName,"Vendor") then
                        Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Catalog Details","catalog details are filled")
                        If fClick(objPage.WebButton("btnAddtoCart"),"Add to Cart")     Then  
                            Call fSynUntilObjExists(objPage.WebButton("btnProceedCheckout"),MID_WAIT)
                            fAribaCreateRequsitionAsNonCatalogItem = True
                            Call fRptWriteReport("PASSWITHSCREENSHOT", "Fill Catalog Details","catalog details are filled")
                        Else
                            Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
                            fAribaCreateRequsitionAsNonCatalogItem = False    
                           Call fRptWriteResultsSummary() 
            				Exit Function
                        End If
                        
                    Else
                        Call fRptWriteReport("Fail","Enter value in "&chr(34)&strFieldName&Chr(34) ,"Value is not entered in "&chr(34)&strFieldName&Chr(34))
                        fAribaCreateRequsitionAsNonCatalogItem = False            
                        Call fRptWriteResultsSummary() 
            			Exit Function    
                    End  IF    
                Else
                    Call fRptWriteReport("Fail","Verify Updated Amount","Amount is not updated")
                    fAribaCreateRequsitionAsNonCatalogItem = False    
                    Call fRptWriteResultsSummary() 
            		Exit Function    
                End If
             Else
                 Call fRptWriteReport("Fail","Click on "&strButtonName, strButtonName&" "&"not been clicked")
                fAribaCreateRequsitionAsNonCatalogItem = False    
                Call fRptWriteResultsSummary() 
            	Exit Function
             End  If
        End  IF 
 
        'Verify the Proceed to check window appears
        If fSynUntilObjExists(objPage.WebButton("btnProceedCheckout"),MIN_WAIT) Then
            Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Proceed to Check","Procced to check is verified Successfully")
            'Click on button
            fClick objPage.WebButton("btnProceedCheckout"),"Proceed to Check"    
        Else
            Call fRptWriteReport("Fail", "Verify Proceed to Check","Procced to check is not appeared")    
            fAribaCreateRequsitionAsNonCatalogItem = False    
            Call fRptWriteResultsSummary() 
            Exit Function         
        End If                
    Set objPage = Nothing  
On error goto 0    
 End Function
 
 
 'Page Down
 Function fPgDown()
 	 ' W Shell script for page down
        Set objWshShell = CreateObject("WScript.shell")
        objWshShell.SendKeys "{PGDN}"
 End Function


'******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fCreateMultiLineNonCatalogItem
'	Objective							:					Used to Create Multiline Non Catalog Items for Ariba
'	Input Parameters					:					objDataDict,iRowCountRef,iLineItemCount
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					15.0
'	QC Version							:		
'	Pre-requisites						:					NILL  
'	Created By							:					
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fCreateMultiLineNonCatalogItem(objDataDict,iRowCountRef)
 	On error resume next
 	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	fCreateNonCatalogItem = False 	
 	strType = objDataDict.Item("TypeOfPurchase" & iRowCountRef)
 	strPurchaseOrg = objDataDict.Item("PurchaseOrg" & iRowCountRef)
' 	strDescription = objDataDict.Item("Description" & iRowCountRef)
 	strCommodityType = objDataDict.Item("CommodityType" & iRowCountRef)
 	strCommodityValue = objDataDict.Item("CommodityValue" & iRowCountRef)
 	strItemCategory = objDataDict.Item("ItemCategory" & iRowCountRef)
 	strCostCenter =  objDataDict.Item("CostCenter" & iRowCountRef)
	strCurrencyType = objDataDict.Item("CurrencyType" & iRowCountRef)
	strCurrencyName = objDataDict.Item("CurrencyName" & iRowCountRef)
	strVendorType = objDataDict.Item("VendorType" & iRowCountRef)
	strVendorName = objDataDict.Item("VendorName" & iRowCountRef)
	strBillType = objDataDict.Item("BillType" & iRowCountRef)
	strBillName = objDataDict.Item("BillName" & iRowCountRef)
	strCostCenter = objDataDict.Item("CostCenter" & iRowCountRef)
	strCostcenterValue = objDataDict.Item("CostCenterValue" & iRowCountRef)
	strShipToType = objDataDict.Item("ShipToType" & iRowCountRef)
	strShipToValue = objDataDict.Item("ShipToValue" & iRowCountRef)
	strText = objDataDict.Item("Text" & iRowCountRef)
	intQuantity = objDataDict.Item("Quantity" & iRowCountRef)
	intPrice = objDataDict.Item("Price" & iRowCountRef)
	iLineItemCount = objDataDict.Item("LineItemCount" & iRowCountRef)
		
	'Set the object to the page
	Call fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement"),MIN_WAIT)				
	Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")	
	objPage.Sync
 	Wait MIN_WAIT	
 	If fSynUntilObjExists(objPage.WebButton("btnAddNonCatalogItem"),MID_WAIT) Then		
		For iItem = 1 To iLineItemCount
 	
	 	'Click on NonCatolog button
		fClick objPage.WebButton("btnAddNonCatalogItem"),"Add Non-Catolog" 		
		'Verify Create Non-Catolog Screen
	 	If fSynUntilObjExists(objPage.WebElement("weCreateNonCatalogItem"),MID_WAIT) Then 			
			'Selec SearchMore in Purch Org: 
			fSelect objPage.SAPWebExtMenu("wMnuPurchasingOrg"),"Search more","SearchMore"
			
			'Select the Value on Purchase Org
			Call fChooseValueforSearchField(strType,strPurchaseOrg)
			
			'Enter Description
			fEnterText objPage.WebEdit("txtDescription"),"test","Description"
			objPage.Sync		
			
			'Enter Commodity Code
			fSelect objPage.SAPWebExtMenu("wMnuCommidityCode"),"Search more","SearchMore"
			
			'Select the Commodity Type
			Call fChooseValueforSearchField(strCommodityType,strCommodityValue)		
			
			'Enter Item Category
			fSelect objPage.SAPWebExtList("wlstItemCatogery"),strItemCategory,"ItemCategory"
			
	 		'Select the Account Type
			 fSelect objPage.SAPWebExtList("wlstCostCenter"),strCostCenter,"CostCenter"	
			 
			 'Enter Quantity
			 fEnterText objPage.WebEdit("txtQuantity"),intQuantity,"Quantity"
			 
			 'Select the CurrencyType
			 fSelect objPage.SAPWebExtMenu("wMnuPriceType"),"Other...","Others"
			 Call fChooseValueforSearchField(strCurrencyType,strCurrencyName)
	
			'Enter Price
			fEnterText objPage.WebEdit("txtPrice"),intPrice,"Price"		
			
			Wait MIN_WAIT
			'Click on Update button
			fClick objPage.WebButton("btnUpdateAmount"),"Update Amount"	
			 
			'Verify the Updated Amount
			 intAmountValue = intQuantity*intPrice
			 Wait MIN_WAIT
			 
			 Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")	
			 objPage.Sync
			 'Get the Value from app
			 intAmount = fWebtableGetCelldata(objPage.WebTable("wtbAmount"),11,3,"AmountFields")
			 arrAmount = Split(intAmount," ")
			If strComp(intAmountValue,clng(arrAmount(0))) = 0 Then
				Call fRptWriteReport("Pass", "Verify Updated Amount","Amount field is updated as "&intAmount)
				fCreateNonCatalogItem = True
			Else
				Call fRptWriteReport("Fail","Verify Updated Amount","Amount is not updated")	
			End If
			Wait MIN_WAIT
			
			'Enter Vendor under Supplier information
			fSelect objPage.SAPWebExtMenu("wMnuVendor"),"Search more","SearchMore"
			Call fChooseValueforSearchField(strVendorType,strVendorName)
			 
			'Click on Add to cart
			fClick objPage.WebButton("btnAddtoCart"),"Add to Cart"	
			
			'Verify the Proceed to check window appears
			If fSynUntilObjExists(objPage.WebButton("btnProceedCheckout"),MIN_WAIT) Then
				Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Proceed to Check","Procced to check is verified Successfully")
				'Click on button
				fClick objPage.WebButton("btnProceedCheckout"),"Proceed to Check"	
			Else
				Call fRptWriteReport("Fail", "Verify Proceed to Check","Procced to check is not appeared")		
			End If			
			
			'Verify the Cart Summary
			Call fVerifyCartSummary(arrAmount(0))		
	 		If iItem < iLineItemCount Then
	 			fClick objPage.WebButton("btnContinueShopping"),"ContinueShopping"
	 		End If
	 	End If	
 	Next	
		intRows = objPage.WebTable("wbtLineItemDetails").RowCount	
		
'        To Select Each Item and provide the values of BillTo, CostCenter, ShipTO fields
		For iRows = 2 To intRows
		
			strDesc  = objPage.WebTable("wbtLineItemDetails").GetCellData(iRows,4)
			If Not instr(strDesc, "The specified cell does not exist") > 0 Then
				Set oChkBox = objPage.WebTable("wbtLineItemDetails").ChildItem(iRows,1,"WebCheckBox",0)
				oChkBox.Click
				
				'Select the Actions under LineItems
				fSelect objPage.SAPWebExtMenu("wMnuActionItem"),"Edit","Edit"
				Wait 5
				Set objWshShell = CreateObject("WScript.shell")
				objWshShell.SendKeys "{PGDN}"
				objWshShell.SendKeys "{PGDN}"
						
				'Select BillTo Field
				fSelect objPage.SAPWebExtMenu("wMnuBillTo"),"Search more","SearchMore"
				Call fChooseValueforSearchField(strBillType,strBillName)
				
				Wait 5
				
				objWshShell.SendKeys "{PGDN}"
				objWshShell.SendKeys "{PGDN}"
				'Select CostCenter
				fSelect objPage.SAPWebExtMenu("wMnuCostCenter"),"Search more","SearchMore"
				Call fChooseValueforSearchField(strCostCenter,strCostcenterValue)
				
				Wait 5
				objWshShell.SendKeys "{PGDN}"
				objWshShell.SendKeys "{PGDN}"
				'Enter ShipTo
				fSelect objPage.SAPWebExtMenu("wMnuShipTo"),"Search more","SearchMore"
				Call fChooseValueforSearchField(strShipToType,strShipToValue)				
				
				'Click on OK button
				fClick objPage.WebButton("btnOK"),"OK"					
				Wait 5   				
			End If			
		Next
		
		'Click on Submit button
		fClick objPage.WebButton("btnSubmit"),"Submit"
		
		Wait 5
		
		'Verify the Requistion is submitted 
		strRequisitionText = objPage.WebElement("weViewRequisition").GetROProperty("innertext")
		If Instr(strRequisitionText,strText) > 0 Then
			fCreateNonCatalogItem = True
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify the Requsition is submitted","Requisition is submitted Successfully")
		Else
			Call fRptWriteReport("Fail", "Verify the Requsition is submitted","Failed to submit the Requisition")		
		End If			
 		
 	End If	

 	On error goto 0
 	
 	Set objPage = Nothing
 	
 End Function

'******************************************************************************************************************************************************************************
'    Function Name                        :        fAribaSplitAccounting
'    Objective                            :        Used to perform Split Accounting functionality
'    Input Parameters                     :        ByRef strCostcenterValue_2, ByRef intSplitPerc_1, ByRef intSplitPerc_2
'    Output Parameters                    :        
'    Date Created                         :        05-21-2020
'    UFT Version                          :        UFT 15.0    
'    Pre-requisites                       :        NIL  
'    Created By                           :        Cigniti                         
'    Modification Date                    :           
'******************************************************************************************************************************************************************************
Public Function fAribaSplitAccounting(ByRef strCostcenterValueTwo, ByRef intSplitPercOne, ByRef intSplitPercTwo)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    If fSynUntilObjExists(Browser("brAribaSpendManagement").Page("pgAribaSpendManagement").WebButton("btnSplitAccounting"),MID_WAIT) = True Then       
        Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
        objPage.Sync
       
        'Click on Split ccounting button
        Call fClick(objPage.WebButton("btnSplitAccounting"),"Split Accounting Button")
       
        'Enter Split percentage for 1st cost center
        objPage.WebEdit("txtSplitPercentageOne").Highlight
        Call fEnterText(objPage.WebEdit("txtSplitPercentageOne"),intSplitPercOne,"Split Percentage 1")
        Call fPgDown()
        Wait MIN_WAIT
       
        'Enter 2nd cost center
        Call fClick(objPage.WebElement("weMnuCostCenterTwo"),"Cost Center menu")
        Call fSelect(objPage.SAPWebExtMenu("wMnuCostCenterTwo"),"Search more","SearchMore")
        Call fChooseValueforSearchField("Cost Center",strCostcenterValueTwo)
        Wait MIN_WAIT
       
        'Enter Split percentage for 2nd cost center
        objPage.WebEdit("txtSplitPercentageTwo").Highlight
        Call fEnterText(objPage.WebEdit("txtSplitPercentageTwo"),intSplitPercTwo,"Split Percentage 2")
        Set objWshShell = CreateObject("WScript.shell")
        objWshShell.SendKeys "{TAB}"
        Wait MIN_WAIT
        Call fRptWriteReport("PASSWITHSCREENSHOT", "Add split ccounting screen","Split percentage added to both Cost Centres")           
       
        'Click on OK button
        Call fClick(objPage.WebButton("btnOK"),"OK Button")
    Else
        Call fRptWriteReport("Fail", "Split Accounting functionality","Split Accounting button does NOT exist")       
        Call fRptWriteResultsSummary()
        Exit Function
    End  If
    On error goto 0
End  Function

''******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fCreateNonCatalogItemProjectWBS
'	Objective							:					Used to Create Requisition for Ariba and click on Continue Shopping
'	Input Parameters					:					objDataDict,iRowCountRef
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					15.0
'	QC Version							:		
'	Pre-requisites						:					NILL  
'	Created By							:					
'	Modification Date					:		   
'******************************************************************************************************************************************************************************************************************************************		
Public Function fCreateNonCatalogItemProjectWBS(objDataDict,iRowCountRef)
 	On error resume next
 	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	fCreateNonCatalogItem = False 	
 	strType = objDataDict.Item("TypeOfPurchase" & iRowCountRef)
 	strPurchaseOrg = objDataDict.Item("PurchaseOrg" & iRowCountRef)
' 	strDescription = objDataDict.Item("Description" & iRowCountRef)
 	strCommodityType = objDataDict.Item("CommodityType" & iRowCountRef)
 	strCommodityValue = objDataDict.Item("CommodityValue" & iRowCountRef)
 	strItemCategory = objDataDict.Item("ItemCategory" & iRowCountRef)
 	strCostCenter =  objDataDict.Item("CostCenter" & iRowCountRef)
	strCurrencyType = objDataDict.Item("CurrencyType" & iRowCountRef)
	strCurrencyName = objDataDict.Item("CurrencyName" & iRowCountRef)
	strVendorType = objDataDict.Item("VendorType" & iRowCountRef)
	strVendorName = objDataDict.Item("VendorName" & iRowCountRef)
	strBillType = objDataDict.Item("BillType" & iRowCountRef)
	strBillName = objDataDict.Item("BillName" & iRowCountRef)
	strCostCenter = objDataDict.Item("CostCenter" & iRowCountRef)
	strCostcenterValue = objDataDict.Item("CostCenterValue" & iRowCountRef)
	strShipToType = objDataDict.Item("ShipToType" & iRowCountRef)
	strShipToValue = objDataDict.Item("ShipToValue" & iRowCountRef)
	strText = objDataDict.Item("Text" & iRowCountRef)
	intQuantity = objDataDict.Item("Quantity" & iRowCountRef)
	intPrice = objDataDict.Item("Price" & iRowCountRef)		
	strAccountAssignment = objDataDict.Item("AccountAssignment" & iRowCountRef)
	strCapitalWBS = objDataDict.Item("CapitalProjectWBS" & iRowCountRef)
	strGLAccount = objDataDict.Item("GLAccount" & iRowCountRef)	
	strScenario = objDataDict.Item("Scenario" & iRowCountRef)	
		'Set the object to the page
		Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")	
	 	Wait MIN_WAIT	
		'Click on NonCatolog button
		fClick objPage.WebButton("btnAddNonCatalogItem"),"Add Non-Catolog" 		
		'Verify Create Non-Catolog Screen
	 	If fSynUntilObjExists(objPage.WebElement("weCreateNonCatalogItem"),MID_WAIT) Then 			
			'Selec SearchMore in Purch Org: 
			fSelect objPage.SAPWebExtMenu("wMnuPurchasingOrg"),"Search more","SearchMore"			
			'Select the Value on Purchase Org
			Call fChooseValueforSearchField(strType,strPurchaseOrg)			
			'Enter Description
			fEnterText objPage.WebEdit("txtDescription"),"test","Description"
			objPage.sync					
			'Enter Commodity Code
			fSelect objPage.SAPWebExtMenu("wMnuCommidityCode"),"Search more","SearchMore"			
			'Select the Commodity Type
			Call fChooseValueforSearchField(strCommodityType,strCommodityValue)					
			'Enter Item Category
			fSelect objPage.SAPWebExtList("wlstItemCatogery"),strItemCategory,"ItemCategory"			
	 		'Select the Account Type
			fSelect objPage.SAPWebExtList("wlstCostCenter"),strCostCenter,"CostCenter"				 
			'Enter Quantity
			fEnterText objPage.WebEdit("txtQuantity"),intQuantity,"Quantity"			 
			'Select the CurrencyType
			fSelect objPage.SAPWebExtMenu("wMnuPriceType"),"Other...","Others"
			Call fChooseValueforSearchField(strCurrencyType,strCurrencyName)	
			'Enter Price
			fEnterText objPage.WebEdit("txtPrice"),intPrice,"Price"			
			Wait MIN_WAIT
			'Click on Update button
			fClick objPage.WebButton("btnUpdateAmount"),"Update Amount"				 
			'Verify the Updated Amount
			intAmountValue = intQuantity*intPrice
			Wait MIN_WAIT
			Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")				 
			'Get the Value from app
			intAmount = fWebtableGetCelldata(objPage.WebTable("wtbAmount"),11,3,"AmountFields")
			arrAmount = Split(intAmount," ")
			If strComp(intAmountValue,clng(arrAmount(0))) = 0 Then
				Call fRptWriteReport("Pass", "Verify Updated Amount","Amount field is updated as "&intAmount)
				fCreateNonCatalogItem = True
			Else
				Call fRptWriteReport("Fail","Verify Updated Amount","Amount is not updated")	
			End If
			Wait MIN_WAIT			
			'Enter Vendor under Supplier information
			fSelect objPage.SAPWebExtMenu("wMnuVendor"),"Search more","SearchMore"
			Call fChooseValueforSearchField(strVendorType,strVendorName)			 
			'Click on Add to cart
			fClick objPage.WebButton("btnAddtoCart"),"Add to Cart"				
			'Verify the Proceed to check window appears
			If fSynUntilObjExists(objPage.WebButton("btnProceedCheckout"),MIN_WAIT) Then
				Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Proceed to Check","Procced to check is verified Successfully")
				'Click on button
				fClick objPage.WebButton("btnProceedCheckout"),"Proceed to Check"	
			Else
				Call fRptWriteReport("Fail", "Verify Proceed to Check","Procced to check is not appeared")		
			End If		
			'Verify the Cart Summary
			Call fVerifyCartSummary(arrAmount(0))				
			'Select the Actions under LineItems
			fSelect objPage.SAPWebExtMenu("wMnuActionItem"),"Edit","Edit"
			Wait MIN_WAIT			
			'Select Account Assignment Field	
			objPage.SAPWebExtList("wlstAccAssignment").Highlight	
			objPage.SAPWebExtList("wlstAccAssignment").Select strAccountAssignment
			Wait MIN_WAIT
			'Select BillTo Field
			fSelect objPage.SAPWebExtMenu("wMnuBillTo"),"Search more","SearchMore"
			Call fChooseValueforSearchField(strBillType,strBillName)		
			Wait MIN_WAIT			
			'Enter GL Account
			objPage.WebEdit("txtGLAccount").Highlight
			objPage.WebEdit("txtGLAccount").Set strGLAccount			
			'Select CostCenter
			objPage.WebEdit("txtAccCostCenter").Highlight
			objPage.WebEdit("txtAccCostCenter").Set strCostcenterValue		
			'Enter Project WBS Element			
			If strScenario="ProjectWBS" Then
				objPage.WebEdit("txtProjectWBS").Highlight
				fEnterText objPage.WebEdit("txtProjectWBS"),strCapitalWBS,"ProjectWBS Element"				
				Wait MIN_WAIT
			Else
				objPage.WebEdit("txtProjectWBS").Highlight
				fEnterText objPage.WebEdit("txtProjectWBS"),strCapitalWBS,"ProjectWBS Element"
				objWshShell.SendKeys "{TAB}"
				Wait MIN_WAIT				
			End If			
			'Enter ShipTo
			fSelect objPage.SAPWebExtMenu("wMnuShipTo"),"Search more","SearchMore"
			Call fChooseValueforSearchField(strShipToType,strShipToValue)			
			'Click on OK button
			fClick objPage.WebButton("btnOK"),"OK"
			Wait MIN_WAIT   
			'Click on Submit button
			fClick objPage.WebButton("btnSubmit"),"Submit"			
			Wait MIN_WAIT
			'Verify the Requistion is submitted 
			strRequisitionText = objPage.WebElement("weViewRequisition").GetROProperty("innertext")
			If Instr(strRequisitionText,strText) > 0 Then
				fCreateNonCatalogItem = True
				Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify the Requsition is submitted","Requisition is submitted Successfully")
			Else
				Call fRptWriteReport("Fail", "Verify the Requsition is submitted","Failed to submit the Requisition")		
			End If			
	 	End If 	
 	On error goto 0
 End Function
''******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fAribaCreateRequisitionCapitalWBS
'	Objective							:					Used to Create Requisition for Ariba and click on Continue Shopping
'	Input Parameters					:					objDataDict,iRowCountRef
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					15.0
'	QC Version							:		
'	Pre-requisites						:					NILL  
'	Created By							:					
'	Modification Date					:		   
'********************************************************************************************************************************************************************
Public Function fAribaCreateRequisitionCapitalWBS(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	fAribaCreateRequisition = False	
	'Fetching Data from the Testdata file
	strTitle = objDataDict.Item("Title" & iRowCountRef)
	strOnBehalfOf = objDataDict.Item("BehalfOf" & iRowCountRef)
	intCompanyCode = objDataDict.Item("CompanyCode" & iRowCountRef)
	strDeliverTo = objDataDict.Item("DeliverTo" & iRowCountRef)
	strComments = objDataDict.Item("Commnets" & iRowCountRef)
	strAttachRequ = objDataDict.Item("AttchmentRequired" & iRowCountRef)
	strApprovalFlowdata = objDataDict.Item("ApprovalFlowtext" & iRowCountRef)	
	strType = objDataDict.Item("TypeName" & iRowCountRef)
	strTypeOfCompany = objDataDict.Item("TypeOfCompany" & iRowCountRef)
	strMenuItem = objDataDict.Item("MenuItem" & iRowCountRef)
	strSubItem = objDataDict.Item("SubItem" & iRowCountRef)	
	strSerachScreen = objDataDict.Item("SearchScreen" & iRowCountRef)		
	'Set the object to the page
	Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")		
	'Select the Menu Items to Create Requisition
	Call fselectRecentManageCreateMenu(strMenuItem,strSubItem)		
	'Enter value in title	
	fEnterText objPage.WebEdit("txtTitle"),strTitle,"Title"		
	'Enter value in On Behalf of Field
	fSelect objPage.SAPWebExtMenu("wMnuOnBehalfOf"),"Search more","SearchMore"		
	'Select the Value on Behalf of
	Call fChooseValueforSearchField(strType,strOnBehalfOf)		
	'Enter Date in Date Filed
	dtCalenderDate = fGetCalenderDate()
	Wait MIN_WAIT
	fEnterText objPage.SAPWebExtCalendar("wCalCalenderdate"),dtCalenderDate,"Delay Purchase"		
	'Enter Company Code
	fSelect objPage.SAPWebExtMenu("wMnuCompanyCode"),"Search more","SearchMore"		
	'Select value for Companycode 
	Call fChooseValueforSearchField(strTypeOfCompany,intCompanyCode)		
	'Enter Value in Deliver to
	fEnterText objPage.WebEdit("txtDeliverTo"),strDeliverTo,"DeliverTo"		
	'Enter Value in comments
	fEnterText objPage.WebEdit("txtComments"),strComments,"Comments"		
	'Select the checkbox if it is unchecked
	strCheckStatus = objPage.SAPWebExtCheckBox("wChkVisibleSupplier").GetROProperty("state")
		If strCheckStatus = False Then
			fEnterText objPage.SAPWebExtCheckBox("wChkVisibleSupplier"),strchecked,"VisibleCheck"
		End If		
	'click on ContinueShopping button
	fClick objPage.WebButton("btnContinueShopping"),"ContinueShopping"
		'Verify Catlog Home Page
		If fSynUntilObjExists(objPage.WebElement("weCatalogHome"),MID_WAIT) Then
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Catlog Home Page","Catlog Home page exists Successfully")
			fAribaCreateRequisition = True
		Else
			Call fRptWriteReport("Fail", "Verify Catlog Home Page","Unable to load the Catlog Home Screen")	
		End If
	'Create Non_Catelog	for ProjectCapitalWBS
		Call fCreateNonCatalogItemProjectWBS(objDataDict,iRowCountRef)		
		Wait MID_WAIT		
		'View Requisition
		Call fViewRequisition()		
		'				Get the Requisition Number
		strReqID = fAribaCaptureReqID()		
		'Capture PO Number
		intPONumber = fAribaGeneratePurchaseOrderTwo(objDataDict,iRowCountRef)		
	On error goto 0	
End Function
''******************************************************************************************************************************************************************************************************************************************
'   Function Name		 				:					fSelectMutliLineItems
'	Objective							:					Used to Select multi line items and enter deatils
'	Input Parameters					:					objDataDict,iRowCountRef
'	Output Parameters					:					NIL
'	Date Created						:					
'	UFT Version							:					15.0
'	QC Version							:		
'	Pre-requisites						:					NILL  
'	Created By							:					
'	Modification Date					:		   
'********************************************************************************************************************************************************************

Public Function fSelectMutliLineItems(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	fSelectMutliLineItems = False
	Wait MIN_WAIT
	'Set the object to the page
	Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")	
	'Get the row count for added items
	intRowCount = objPage.WebTable("wbtLineItemDetails").RowCount			
	' To Select Each Item and provide the values of BillTo, CostCenter, ShipTO fields
	 For ivalue = 2 To intRowCount		
		strDesc  = objPage.WebTable("wbtLineItemDetails").GetCellData(ivalue,4)
			If Not instr(strDesc, "The specified cell does not exist") > 0 Then
				Set oChkBox = objPage.WebTable("wbtLineItemDetails").ChildItem(ivalue,1,"WebCheckBox",0)
				oChkBox.Set "ON"
					Wait MIN_WAIT				
				''Choose Edit or copy or Delete 
				  Call fAribaPerformOperationOnLineItem(objDataDict,iRowCountRef)
				'Fill Requisition Accounting Line Item data
				Call fAribaRequisitionAccountingLineItem(objDataDict,iRowCountRef)
				'Fill Requisition Shipping By Line Item data
				Call fAribaRequisitionShippingByLineItem(objDataDict,iRowCountRef)
				
				Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
					If  ivalue  < intRowCount-1 Then
						'Click on OK button
						Call fClick(objPage.WebButton("btnOK"),"OK")
					Else
						Exit for
					End  IF	
			End  IF	
		Next	
End Function

Function fPgDown_withNumber(ByVal ScrollCount)
 	 ' W Shell script for page down
        Set objWshShell = CreateObject("WScript.shell")
        For i = 1 To ScrollCount
        	objWshShell.SendKeys "{PGDN}"
        Next 
        
 End Function
 
 
  
 '******************************************************************************************************************************************************************************
'	Function Name						:		fAribaGetPaymentTermID
'	Objective							:		Get Payment Term ID name / Number
'	Input Parameters					:		Nil
'	Output Parameters					:		NIl
'	Date Created						:		13-Jun-2020
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Automation Team
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Function fAribaGetPaymentTermID()

	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	
	Dim strPaymentTermID
	Dim objHomePage
	Dim strOrderID
	
 	Set objHomePage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")

	 If fVerifyObjectExist(objHomePage.WebTable("wbtOrders")) Then
	 	 	'Click on PO Number
			If objHomePage.WebTable("wbtOrders").ChildItem(2,2,"Link",0).exist(2) Then
				Call objHomePage.WebTable("wbtOrders").ChildItem(2,2,"Link",0).click
				Call fSynUntilObjExists(objHomePage.WebTable("tblPaymentTerms"),MIN_WAIT) 
					'Get PaymentTerms ID
					If fVerifyObjectExist(objHomePage.WebTable("tblPaymentTerms")) Then
		 	 			strPaymentTermID = objHomePage.WebTable("tblPaymentTerms").GetCellData(2,3)
		 	 			Call fPgDown_withNumber(2)
		 	 			Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strPaymentTermID,"TestData","PaymentTerm")
		 	 			Call fRptWriteReport("PassWithScreenshot", "Capture Payment Term ID","Payment Term ID is -" & strPaymentTermID)
		 	 			fAribaGetPaymentTermID = True
		 	 		Else
						Call fRptWriteReport("Fail", "Get Payment Term ID","Payment Term Table not displayed ")
						fAribaGetPaymentTermID = False
						Exit Function	
		 	 		End If
			End If
	 Else
			Call fRptWriteReport("Fail", "Get Order ID","Order ID not displayed in Search results table")
			fAribaGetPaymentTermID = False
			Exit Function
	 End IF
	
	Set objHomePage = Nothing
	On error goto 0
End Function


'******************************************************************************************************************************************************************************
'	Function Name						:		fNavigatePurchaseRequisition
'	Objective							:		Used to Navigate to PurchaseRequisite Page
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Function fNavigatePurchaseRequisition()
	
	strActReqID = fFetchActRequisitionID()
	Set objHomePage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
	strReqID = fGetSingleValue("RequisitionNumber","TestData",Environment("TestName"))
	If strActReqID = strReqID Then
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Requisition ID: "&strReqID&" is Displayed in Search Results","Requisition ID: "&strReqID&" is Displayed in Search Results,Successfully")
		objHomePage.WebTable("wbtReqIDSearchResults").ChildItem(2,4,"Link",0).Click
		Call fSynUntilObjExists(objHomePage.WebTabStrip("tabRequisitionTab"),MID_WAIT)
		If objHomePage.WebTabStrip("tabRequisitionTab").Exist(1) Then
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify navigated to "&strReqID&" - CreatePurchaseRequisition Page","Navigated to "&strReqID&" - CreatePurchaseRequisition Page Successfully")			
		Else
			Call fRptWriteReport("Fail", "Verify navigated to "&strReqID&" - CreatePurchaseRequisition Page","Failed to navigate to "&strReqID&" - CreatePurchaseRequisition Page")
			Call fRptWriteResultsSummary() 
            ExitAction
		End If
	Else
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Requisition ID: "&strReqID&" is Displayed in Search Results","Failed to Display the Requisition ID: "&strReqID&" in Search Results")
		Call fRptWriteResultsSummary() 
        ExitAction
	End If
	
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fClickHistory
'	Objective							:		Used to Navigate to History tab in Requisition Details
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Function fClickHistory()
	
	Set objHomePage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
	strReqID = fGetSingleValue("RequisitionNumber","TestData",Environment("TestName"))	
	strPurchaseOrder = fGetSingleValue("AutoPONumber","TestData",Environment("TestName"))
	
'	Click on History Tab
	objHomePage.WebElement("html tag:=A","innertext:=History").Click		
	Call fSynUntilObjExists(objHomePage.WebTable("wbtRequisitionHistory"),MID_WAIT)
	If objHomePage.WebTable("wbtRequisitionHistory").Exist(1) Then
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify navigated to "&strReqID&" - CreatePurchaseRequisition --> History Tab Page","Navigated to "&strReqID&" - CreatePurchaseRequisition Page --> History Tab Page Successfully")
		If fVerifySummaryInHistoryTab Then
			objHomePage.WebTable("wbtRequisitionHistory").ChildItem(2,5,"Link",0).Click
			Call fSynUntilObjExists(objHomePage.WebTabStrip("tabRequisitionTab"),MID_WAIT)
'				Navigate to Purchase order Detais Page
			If objHomePage.WebTabStrip("tabRequisitionTab").Exist(1) Then
				Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify navigated to "&strPurchaseOrder&" - CreatePurchaseRequisition Details Page","Navigated to "&strPurchaseOrder&" - CreatePurchaseRequisition Details Page Successfully")
			Else
				Call fRptWriteReport("Fail", "Verify navigated to "&strPurchaseOrder&" - CreatePurchaseRequisition Details Page","Failed to Navigate to "&strPurchaseOrder&" - CreatePurchaseRequisition Details Page")
				Call fRptWriteResultsSummary() 
            	ExitAction
			End If			
		End If
	Else
		Call fRptWriteReport("Fail", "Verify navigated to "&strReqID&" - CreatePurchaseRequisition --> History Tab Page Page","Failed to navigated to "&strReqID&" - CreatePurchaseRequisition Page --> History Tab Page")
		Call fRptWriteResultsSummary() 
        ExitAction
	End If	
	
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fVerifySummaryInHistoryTab
'	Objective							:		Used to Verify Summary in History tab in Requisition Details
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Function fVerifySummaryInHistoryTab()	
	
	strPurchaseOrder = fGetSingleValue("AutoPONumber","TestData",Environment("TestName"))
	blgVerifySummary = False
	strReqID = fGetSingleValue("RequisitionNumber","TestData",Environment("TestName"))
	Set objHomePage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
	intRows = objHomePage.WebTable("wbtRequisitionHistory").RowCount
	strActSummary = objHomePage.WebTable("wbtRequisitionHistory").GetCellData(2,5)
	
	If Instr(strReqID,"-V2") > 0 Then
		strExpSummary = "Order "&strPurchaseOrder& " has been canceled. The Cancellation of order "&strPurchaseOrder&" was successfully sent via Ariba Network"
	Else
		strExpSummary = "Order "&strPurchaseOrder& " was successfully sent via Ariba Network"
	End If
	
	If instr(strActSummary,strExpSummary) > 0 Then
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify "&strExpSummary&" is displayed in History","Successfully Displayed "&strExpSummary&" in History")		
		blgVerifySummary = True
	Else
		Call fRptWriteReport("Fail", "Verify "&strExpSummary&" is displayed in History",strExpSummary&" is displayed Successfully in History")
		Call fRptWriteResultsSummary() 
        ExitAction
	End if 
	fVerifySummaryInHistoryTab = blgVerifySummary
	
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fCancelPurchaseOrder
'	Objective							:		Used to Cancel the PurchaseOrder in Ariba
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Function fCancelPurchaseOrder()
	
	strCancelComments = objDataDict.Item("CancelCommnets" & iRowCountRef)	
	Set objHomePage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
	strReqID = fGetSingleValue("RequisitionNumber","TestData",Environment("TestName"))	
	
'	To Click on Cancel Button
	Call fClick(objHomePage.WebButton("btnCancel"), "Cancel")
	
'	To Enter Cancel Comments
	fEnterText objHomePage.WebEdit("txtCancelComments"),strCancelComments,"CancelComments"
	fClick objHomePage.WebButton("btnOK"),"OK"	
	
	Call fSynUntilObjExists(objHomePage.WebElement("weCancellationNotification"),MID_WAIT)
	strActCancelNotification =   objHomePage.WebElement("weCancellationNotification").GetROProperty("innerText")
	strExpCancelNotification = 	"A new version of "&strReqID&" has been generated"
	
	If instr(strActCancelNotification,strExpCancelNotification) > 0 Then
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify A new version of "&strReqID&" has been generated  Successfully","A new version of "&strReqID&" has been generated Successfully")
		Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strReqID&"-V2","TestData","RequisitionNumber")
	Else
		Call fRptWriteReport("Fail", "Verify A new version of "&strReqID&" has been generated Successfully","Failed display A new version of "&strReqID&" has been generated")
		Call fRptWriteResultsSummary() 
        ExitAction
	End If
	
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fFetchActRequisitionID
'	Objective							:		Used to Fetch the Actual Requisition ID in Ariba
'	Input Parameters					:		
'	Output Parameters					:		strActReqID
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Function fFetchActRequisitionID()
	
	Set objHomePage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")
 	strReqID = fGetSingleValue("RequisitionNumber","TestData",Environment("TestName"))
 	strPurchaseOrder = fGetSingleValue("AutoPONumber","TestData",Environment("TestName"))
 	strSearchScreen = fGetSingleValue("SearchItemInScreen","TestData",Environment("TestName"))
	flgSearchTextInAribaPage = fSearchTextInAribaPage (strSearchScreen)
	
	If flgSearchTextInAribaPage Then
		Call fRptWriteReport("Pass", "Verify Clicked on "&strSerachScreen,"Clicked successfully on "&strSerachScreen)
		
		'To Validate Whether the screen is navigated to required Screen    
        flgNavigateScreen = fVerifyProperty(objHomePage.SAPWebExtList("html tag:=DIV","name:=_wh1eo"),"selection",strSearchScreen) 
		
		If flgNavigateScreen Then
	        Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Screen navigation","Navigated to "&strSearchScreen&" Screen Successfully")
	        
'	        To Enter Requisition ID
	        objHomePage.WebEdit("acc_name:=ID:","class:=w-txt","html tag:=INPUT").set strReqID  
	 		objHomePage.WebButton("class:=w-btn w-btn-primary aw7_w-btn-primary","html tag:=BUTTON","name:=Search").Click
	 		Call fSynUntilObjExists(objHomePage.WebTable("wbtOrders"),MID_WAIT)
			
			'To fetch Requesition ID
			objHomePage.WebTable("wbtReqIDSearchResults").Highlight
			strActReqID = objHomePage.WebTable("wbtReqIDSearchResults").GetCellData(2,4)
			fFetchActRequisitionID = strActReqID
		Else
			Call fRptWriteReport("Fail", "Verify Navigated to "&strSerachScreen&" Screen","Failed to Navigated to "&strSerachScreen&" Screen")
            Call fRptWriteResultsSummary() 
            ExitAction
		End If	
	Else
		Call fRptWriteReport("Fail", "Verify Clicked on "&strSerachScreen,"Failed to Click on "&strSerachScreen)
		Call fRptWriteResultsSummary() 
        ExitAction
	End If
	
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fDeleteApprovers
'	Objective							:		Used to Delete the Approvers in Ariba
'	Input Parameters					:		
'	Output Parameters					:		strActReqID
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Function fDeleteApprovers()
	
	Set objPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")   
	Set objCloseChild = Description.Create
    objCloseChild("micclass").Value = "WebElement"
    objCloseChild("class").Value = "a-graph-node-icon w-apv-delete-icon"
    objCloseChild("html tag").Value = "SPAN"
    objCloseChild("title").Value = "Delete"
    objCloseChild("visible").value = True
    Call fPgDown()
    Set objCloseApprover = objPage.WebTable("wbtApprovalFlow").ChildObjects(objCloseChild)
    objCountOfCloseApprover = objCloseApprover.Count                
    
    If objCountOfCloseApprover > 0 Then	            
        Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Approvers are available to Delete","Approvers are available to Delete")	            
        'Click on X-Symbolol on each Approver
        For jCount = 0 To objCountOfCloseApprover-1
            objCloseApprover(jCount).Click	                
            If fSynUntilObjExists(objPage.WebButton("btnDeleteApprover"),MID_WAIT) Then
                fClick objPage.WebButton("btnDeleteApprover"),"Delete Approver --> OK"     
                objPage.Sync 
				Wait MIN_WAIT
				Wait MIN_WAIT
            End If					             
        Next	            
       If objPage.WebTable("wbtApprovalFlow").Exist(2) Then 
			 Set objCloseApprover = objPage.WebTable("wbtApprovalFlow").ChildObjects(objCloseChild)
	        objCountOfCloseApprover = objCloseApprover.Count                
	        
	        If objCountOfCloseApprover > 0 Then	            
	            Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Approvers are available to Delete","Approvers are available to Delete")	            
	            
	            'Click on X-Symbolol on each Approver
	            For jCount = 0 To objCountOfCloseApprover-1
	                objCloseApprover(jCount).Click	                
	                If fSynUntilObjExists(objPage.WebButton("btnDeleteApprover"),MID_WAIT) Then
	                    fClick objPage.WebButton("btnDeleteApprover"),"Delete Approver --> OK"     
	                    objPage.Sync
						Wait MID_WAIT								
	                End If					             
	            Next
	        End  If
	     End  IF   
        If fSynUntilObjExists(objPage.WebButton("btnShowApprovalFlow"),MID_WAIT) Then
            Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Approvers are not available to Delete","Approvers are not available to Delete")
        Else
            Call fRptWriteReport("Fail", "Verify Approvers are not available to Delete","Approvers are available to Delete")
            Call fRptWriteResultsSummary()        
            ExitAction
        End If
        
    Else
        Call fRptWriteReport("Fail", "Verify Approvers are available to Delete","Approvers are not available to Delete.. Please verify Test data Once..")
        Call fRptWriteResultsSummary()        
        ExitAction
    End If	
	
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fVerifyStatusAfterDeleteApprovers
'	Objective							:		Used to Verify Status of the Requisition after deletion of Approvers
'	Input Parameters					:		
'	Output Parameters					:		strActReqID
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Function fVerifyStatusAfterDeleteApprovers()
	
	Set obAribajPage = Browser("brAribaSpendManagement").Page("pgAribaSpendManagement")  	
	Call fSynUntilObjExists(obAribajPage,MIN_WAIT)

	'Get the Status of Requisition
	Call fSynUntilObjExists(obAribajPage.WebElement("weStatus"),MIN_WAIT)
	strReqStatus = obAribajPage.WebElement("weStatus").GetROProperty("innertext")
	
'	Verify if the Status is Approved
	If Instr(strReqStatus,"Approved") > 0 Then
		Call fRptWriteReport("Pass", "Verify Status after Creating Requisition","Status is diaplayed as "&strReqStatus)
	Else
		Call fRptWriteReport("Fail", "Verify Status after Creating Requisition","Status is not diaplayed as "&strReqStatus)	
		Call fRptWriteResultsSummary()        
        ExitAction
	End If		
	
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fAribaSupplierCreateCreditMemoNumber
'	Objective							:		Used to Search PO Number in Ariba Supplier 
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		8-Jun-2020
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti (Shivali)
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fAribaSupplierCreateCreditMemoNumber(objDataDict,iRowCountRef)
	On error resume next
	fAribaSupplierCreateCreditMemoNumber = False
	
	'Fetching Data from the Testdata file
	strInvoiceID 	= fGetSingleValue("AutoInvoiceNumber","TestData",Environment("TestName")) 
	
	Set objANSPage = Browser("brAribaSpendManagement").Page("pgAribaNetworkSupplier")   
	
	'Click on Exit link on Invoice details page
	'If fClick(objANSPage.Link("lnkExit"),"Exit") then
		Call fSynUntilObjExists(objANSPage.Link("lnkSupplierOutbox"),MIN_WAIT) 
		'Click Outbox link on home page
		If fClick(objANSPage.Link("lnkSupplierOutbox"),"Outbox") Then
			Call fSynUntilObjExists(objANSPage.Link("lnkInvoices"),MIN_WAIT) 
			'Click on Invoices option from list
			If fClick(objANSPage.Link("lnkInvoices"),"Invoices") Then
				Call fSynUntilObjExists(objANSPage.WebElement("weSerachFilterExpandObj"),MAX_WAIT)  
				'Click on Expanssion arrow on Search area
				If fClick(objANSPage.WebElement("weSerachFilterExpandObj"),"Expand Search") Then
					Call fSynUntilObjExists(objANSPage.WebEdit("txtInvoiceInputbox"),MIN_WAIT)  
					'Enter Invoice number in input box
					Call fEnterText(objANSPage.WebEdit("txtInvoiceInputbox"),strInvoiceID,"Invoice Input")
					'Select Exact number radio button
					Call fClick(objANSPage.SAPWebExtRadioButton("rbtnExactNumber"),"Exact number")
					Wait MIN_WAIT
					'Click on Search button
					Call fClick(objANSPage.WebButton("btnSearch"),"Search Button")
					'Wait until the filter gets enabled and select the first invioce got filtered by clicking on radio button
					Call fSynUntilObjExists(objANSPage.WebTable("wbtInvoicesList"),MID_WAIT)
					If fClick(objANSPage.SAPWebExtRadioButton("rbtnInvoiceSelection"),"Invoice radio button") Then
						Call fClick(objANSPage.WebButton("btnCreateCreditMemo"),"Create Line Item-Credit Memo button")
						
					Else
						Call fRptWriteReport("Fail","Click on Invoice radio button", "Invoice radio button not been clicked")
						fAribaSupplierCreateCreditMemoNumber = False	
						Call fRptWriteResultsSummary() 
		            	ExitAction
					End  If
					
				Else
					Call fRptWriteReport("Fail","Click on Expand Search", "Expand Search not been clicked")
					fAribaSupplierCreateCreditMemoNumber = False	
					Call fRptWriteResultsSummary() 
	            	ExitAction
				End  If
			Else
				Call fRptWriteReport("Fail","Click on Invoices option", "Invoices option not been clicked")
				fAribaSupplierCreateCreditMemoNumber = False	
				Call fRptWriteResultsSummary() 
            	ExitAction
			End If
		Else
			Call fRptWriteReport("Fail","Click on Outbox", "Outbox not been clicked")
			fAribaSupplierCreateCreditMemoNumber = False	
			Call fRptWriteResultsSummary() 
            ExitAction
		End If
'	Else
'		Call fRptWriteReport("Fail","Click on Exit Link", "Exit Link not been clicked")
'		fAribaSupplierCreateCreditMemoNumber = False	
'		Call fRptWriteResultsSummary() 
'        ExitAction
'	End IF 
	
	'Get Credit Memo number created by entering required details on next page
	strCreditMemoNum = fAribaANSupplierCreateCreditMemoWithDetails(objDataDict,iRowCountRef)
	If strCreditMemoNum <> "" Then
		Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strCreditMemoNum,"TestData","AutoCreditMemoNumber")
'		Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strCreditMemoNum,"TestData","AutoInvoiceNumber")
		Call frptWriteReport("Pass","Create Credit Memo Number","Credit Memo Number :"&strCreditMemoNum&" is generated successfully")
		fAribaSupplierCreateCreditMemoNumber = True
	Else
		Call rptWriteReport("Fail","Create Credit Memo Number","Credit Memo Number is not generated")
	End If
	'On error goto 0	
End Function 


'******************************************************************************************************************************************************************************
'	Function Name						:		fAribaANSupplierCreateCreditMemoEnterDetails
'	Objective							:		Used to Search PO Number in Ariba Supplier 
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		9-Jun-2020
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti (Shivali)
'	Modification Date					:		   
'******************************************************************************************************************************************************************************
Public Function fAribaANSupplierCreateCreditMemoWithDetails(ByRef objDataDict, ByRef iRowCountRef)

	On error resume next
	fAribaANSupplierCreateCreditMemoWithDetails = False
	
	'Fetching Data from the Testdata file
	strCreditMemoReason = objDataDict.Item("ReasonForCreditMemo" & iRowCountRef)
	strCreditMemoQty = objDataDict.Item("CreditMemoQty" & iRowCountRef)
	strAutoPONumber = fGetSingleValue("AutoPONumber","TestData",Environment("TestName")) 	
	strCreditMemo = "CM"&strAutoPONumber
	strFileName = objDataDict.Item("FileName" & iRowCountRef)	
	
	Set objANSPage = Browser("brAribaSpendManagement").Page("pgAribaNetworkSupplier") 
	
	'Enter Credit Memo# in input box
	Call fSynUntilObjExists(objANSPage.WebEdit("txtCreditMemoInput"),MIN_WAIT) 
	Call fEnterText(objANSPage.WebEdit("txtCreditMemoInput"),strCreditMemo,"Credit Memo# Input")
	Call fPgDown_withNumber(2)
	
	'Enter Reason for Credit memo in Input box
	Call fSynUntilObjExists(objANSPage.WebEdit("txtReasonForCreditMemo"),MIN_WAIT) 
	Call fEnterText(objANSPage.WebEdit("txtReasonForCreditMemo"),strCreditMemoReason,"Reason for Credit Memo")
	Call fPgDown_withNumber(1)
	Wait MIN_WAIT
	
 	'Click on Attachmnets in AddHeader listbox
  	fSelect objANSPage.SAPWebExtMenu("wMnuAddAttachments"),"Attachment","AddToHeader-> Attachment"       
  	Call fAribaSupplierAttachFile(strFileName)
  	Wait MIN_WAIT
	Call fPgDown_withNumber(1)
	
	'Enter Qty 
	Call fSynUntilObjExists(objANSPage.WebEdit("txtQtyInput"),MIN_WAIT) 
	If fVerifyObjectExist(objANSPage.WebEdit("txtQtyInput")) Then
		Call fEnterText(objANSPage.WebEdit("txtQtyInput"),strCreditMemoQty,"Credit Memo Qty")
		Wait MIN_WAIT
		Call fPgDown_withNumber(1)
	End If
	
	'Click on Next button
	If fClick(objANSPage.WebButton("btnNext"),"Next button") Then
		Call fSynUntilObjExists(objANSPage.WebButton("btnCMSubmit"),MID_WAIT)
		strCreditMemoNumber = objANSPage.WebElement("weInvoiceNumber").GetROProperty("innertext")
	 	fAribaANSupplierCreateCreditMemoWithDetails = strCreditMemoNumber
		'Click on Submit button
		Call fClick(objANSPage.WebButton("btnCMSubmit"),"Submit button")
		Call fSynUntilObjExists(objANSPage.WebElement("weInvoiceDetail"),MIN_WAIT) 
		If fSynUntilObjExists(objANSPage.WebElement("weInvoiceDetail"),MID_WAIT) Then
			Call fRptWriteReport("PassWithScreenshot", "Create and Capture Credit Memo Number","Credit Memo is Created and Captured Credit Memo Number as "&strCreditMemoNumber)
		Else
			Call fRptWriteReport("Fail", "Create and Capture Credit Memo Number","Unable to Create CreditMemo Number")		
			Call fRptWriteResultsSummary() 
        	ExitAction
		End If	
		
	Else
		Call fRptWriteReport("Fail","Click on Next button", "Next button not been clicked")
		fAribaANSupplierCreateCreditMemoWithDetails = False	
		Call fRptWriteResultsSummary() 
        ExitAction	
	End If
	
	'On error goto 0	
End  Function





