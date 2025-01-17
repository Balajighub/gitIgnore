'***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriLogin
''	Objective						:				Log into Fiori application 
''	Input Parameters				:				strUserRole
''	Output Parameters			    :				Nil
''	Date Created					:				27/April/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		15/May/2020 - Function updated based on User role
'*************************************************************************************************************************************************************************************** 
Public Function fFioriLogin(strUserRole)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Variable declaration
	Dim ObjBrAndPage
	Dim strURL
	Dim strUID
	Dim strPWD
	Dim strLNG
		
    strURL = gstrFioriAnalystAPURL
	strLNG = gstrFioriLanguage
	Select Case Ucase(strUserRole) 
		Case Ucase("APAnalyst")
			strUID = gstrFioriAnalystAPUsername
        	strPWD = gstrFioriAnalystAPPassword
        Case Ucase("APManager")
			strUID = gstrFioriManagerAPUsername
        	strPWD = gstrFioriManagerAPPassword		        
	End Select		   
	   
	    If gstrIEBrowser = "YES" Then
	        SystemUtil.Run "iexplore.exe",gstrFioriAnalystAPURL    
	    ElseIf gstrChromeBrowser = "YES" Then        
	        SystemUtil.Run "Chrome.exe",gstrFioriAnalystAPURL
	    ElseIf gstrFireFoxBrowser = "YES" Then
	        SystemUtil.Run "firefox.exe",gstrFioriAnalystAPURL    
	    End If 
	'Set Browser and Page 
	Set ObjBrAndPage=Browser("brFiori").Page("pgFiori")
	ObjBrAndPage.Sync 
		'Verify Object Exists	
        If fVerifyObjectExist(ObjBrAndPage.WebElement("weCancel")) Then
            Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Fiori Application login page","Fiori home page has been displayed successfully")
        Else
            Call fRptWriteReport("Fail", "Verify Fiori Application login page","Fiori home page has not been displayed")  
            Exit Function
        End  IF
	'Click on Cancel button
	Call fClick(ObjBrAndPage.WebElement("weCancel"), "Cancel")
	' Wait till User Name txt field enabled 
	Call fSynUntilObjExists(ObjBrAndPage.WebEdit("txtUser"),MID_WAIT)
	' Enter data in to User Name txt field 
	Call fEnterText(ObjBrAndPage.WebEdit("txtUser"),strUID,"UserName")
	' Enter data in to Password txt field
	Call fEnterText_SetSecureMode(ObjBrAndPage.WebEdit("txtPassword"),strPWD,"Password")
	'Select data in Language list box
	Call fSelect (ObjBrAndPage.WebList("lstLanguage"),strLNG,"Language")
	'Click on Log on button
	Call fClick(ObjBrAndPage.WebButton("btnLogOn"), "Log On")
	' Wait Home page exist
	Call fSynUntilObjExists(ObjBrAndPage.WebElement("weHome"),MID_WAIT)
		'Verification of Home page
		If fVerifyObjectExist(ObjBrAndPage.WebElement("weHome")) Then
		   	Call fRptWriteReport("PASSWITHSCREENSHOT","Verify Fiori Application login","Fiori Application has been login successfully with user "&strUID)
	   	Else
	   		Call fRptWriteReport("Fail","Verify Fiori Application login","Fiori Application has not been login")
	   		Exit Function
		End  If 
	On error goto 0	
End Function

'***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriLogOut
''	Objective						:				Log out Fiori application 
''	Input Parameters				:				Nil
''	Output Parameters			    :				Nil
''	Date Created					:				27/April/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*************************************************************************************************************************************************************************************** 
Public Function fFioriLogOut()
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Variable declaration
	Dim ObjPgLogOut
	'Set Browser and Page 
	Set ObjPgLogOut=Browser("brFiori").Page("pgFiori")
	Call fFioriLeavePage()
		If fVerifyObjectExist(ObjPgLogOut.WebButton("btnUserIcon")) Then	
			' Wait till Sign Out button enabled 
			Call fSynUntilObjExists(ObjPgLogOut.WebButton("btnUserIcon"),MIN_WAIT)
			'Click on User button
			Call fClick(ObjPgLogOut.WebButton("btnUserIcon"), "User Icon")
			' Wait till Sign Out button enabled 
			Call fSynUntilObjExists(ObjPgLogOut.WebElement("weSignOut"),MIN_WAIT)
			'Click on User button
			Call fClick(ObjPgLogOut.WebElement("weSignOut"), "Sign Out")
			' Wait till OK button appeared in Signout pop up window 
			Call fSynUntilObjExists(ObjPgLogOut.WebButton("btnOK"),MIN_WAIT)
			'Click on OK button in Sing out Pop up window
			Call fClick(ObjPgLogOut.WebButton("btnOK"), "OK")
			'Close Fiori Application
			Call fCloseAllOpenBrowsers("ALL")
				'Verification of LogOut page
				If Not fVerifyObjectExist(Browser("brFiori")) Then
				   	Call fRptWriteReport("Pass","Verify Fiori Application logout","Fiori Application has been logout successfully")
			   	Else
			   		Call fRptWriteReport("Fail","Verify Fiori Application logout","Fiori Application has not been logout")
				End  If
		Else
			Call fRptWriteReport("Fail", "Verify Fiori Application logout","User Icon has not ben found in Fiori Application")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If	
	On error goto 0		
End Function

'***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriSelectSubmenu
''	Objective						:				Used to select submenu
''	Input Parameters				:				objMenu,objSubMenu,strMenu,strSubMenu
''	Output Parameters			    :				Nil
''	Date Created					:				27/04/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'***************************************************************************************************************************************************************************************
Public Function fFioriSelectSubmenu(objMenu,objSubMenu,strMenu,strSubMenu)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Call fSynUntilObjExists(objMenu,MAX_WAIT)
		If fVerifyObjectExist(objMenu) Then	
			'Click on Main menu
			If fClick(objMenu,strMenu) Then
					'Click on Sub menu
					If fClick(objSubMenu,strSubMenu) Then
						Call fRptWriteReport("Pass", "Verify navigation in Fiori Application","Navigated successfully to "&strMenu&"->"&strSubMenu)
					Else
						Call fRptWriteReport("Fail","Verify navigation in Fiori Application","Not able to navigate "&strMenu&"->"&strSubMenu)
						Exit Function
					End If
			Else
				Call fRptWriteReport("Fail","Verify navigation in Fiori Application","Not able to navigate "&strMenu&"->"&strSubMenu)
				Exit Function
			End If
		Else
			Call fRptWriteReport("Fail","Verify navigation in Fiori Application","Not able to navigate "&strMenu)
			Exit Function
		End If
	On error goto 0	
End Function
'***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriFetchAppFromHomePage
''	Objective						:				Used to Search Text In Fiori Home Page 
''	Input Parameters				:				strSearchText
''	Output Parameters			    :				Nil
''	Date Created					:				28/April/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		'05/13/2020 - Read data from excel
'*************************************************************************************************************************************************************************************** 
Public Function fFioriFetchAppFromHomePage(strSearchPageName,strTileName)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Variable declaration
	Dim objHomePage
	'Set Browser and Page 
	Set objHomePage=Browser("brFiori").Page("pgFiori")
'	'Get data from excel
	fFioriFetchAppFromHomePage = False
	' Wait Home page exist
	Call fSynUntilObjExists(objHomePage.WebButton("btnSearch"),MAX_WAIT)
		If fVerifyObjectExist(objHomePage.WebButton("btnSearch")) Then
				'Click on Search button
				If fClick(objHomePage.WebButton("btnSearch"),"Search button") Then
					' Wait till Search txt field exist 
					Call fSynUntilObjExists(objHomePage.WebEdit("txtSearch"),MAX_WAIT)
					Wait(MID_WAIT)'Required
						If strSearchPageName <> "" Then					
						'Enter text in Search txt field
						If fEnterText(objHomePage.WebEdit("txtSearch"),strSearchPageName,"Search field") Then
							' Wait till Search txt field exist 
							Call fSynUntilObjExists(objHomePage.SAPUIButton("btnSearch"),MAX_WAIT)
							'Click on Search txt field
							If fClick(objHomePage.SAPUIButton("btnSearch"),"Search icon") Then
								Wait(MIN_WAIT)'Required
								fFioriFetchAppFromHomePage = True
							End If
						End If	
					Else
						Call fRptWriteReport("Fail","Verify Search button in Fiori Application home page","Search button not found in Fiori Application home page")
						Call fRptWriteResultsSummary()												
	   					Exit Function	
					End If	
				End If
		Else
			Call fRptWriteReport("Fail","Verify Search page","Search page has not been open "&strSearchPageName)
	 	  	Call fRptWriteResultsSummary() 
			Exit Function
		End If
		 'Create Object for Browser   
		Set oDesc = Description.Create()
		oDesc("micclass").value="Browser"
		Set brObj = DeskTop.ChildObjects(oDesc)
		intbrCnt = brObj.Count
		
		If intbrCnt <= 1 Then
			'Create Object for Tile
		    Set objTile = Description.Create
	        objTile("micclass").Value = "WebButton"
	        objTile("html tag").Value = "DIV"
	        objTile("innertext").Value = strTileName
	        objTile("visible").value = True
	        Set objTileButton = objHomePage.ChildObjects(objTile)
	        objCountOfTileButton = objTileButton.Count            ' Get Tile Count  
	             For intTile = 0 to objTileButton.Count -1
	        		objTileButton(intTile).Click
	        		Exit For
	        	 Next
        End IF 
	On error goto 0	
End Function
'***************************************************************************************************************************************************************************************
''    Function Name                   :                fFioriGetGLAccountTableColumnaData
''    Objective                       :                Used to Get GL Account Table Columna Data 
''    Input Parameters                :                strGLAccountTableColumnData
''    Output Parameters               :                arrColData
''    Date Created                    :                29/04/2020
''    UFT/QTP Version                 :                15.0
''    Pre-requisites                  :                NIL  
''    Created By                      :                Cigniti
''    Modification Date               :                   
'*************************************************************************************************************************************************************************************** 
Public Function fFioriGetGLAccountTableColumnaData(strGLAccountTableColumnData)
    On error resume next
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    'Variable declaration
    Dim arrColData()
    Dim strInnerText
    Dim intCount
    Dim objGLFrame
    Dim arrList
    Dim intRowCount
    Dim intColNum
    Dim strColumnName
    'Get data from excel sheet
    arrList = Split(strGLAccountTableColumnData,",")
    intRowCount = arrList(0)
    intColNum = arrList(1)
    strColumnName = arrList(2)
    
    'Set Frame 
    Set ObjPgGLAccount=Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame")
    	If Lcase(strColumnName)<>"stat" Then
			For intItearion = 6 To intRowCount
				If ObjPgGLAccount.WebElement("html id:=M0:46:::"&intItearion&":"&intColNum&"_l").Exist(1) then
					strInnerText = ObjPgGLAccount.WebElement("html id:=M0:46:::"&intItearion&":"&intColNum&"_l").GetROProperty("innertext")
					   If intItearion = 6 Then
					      intCount = 0
					       If Ucase(strInnerText) <> Ucase(strColumnName) Then
					            fFioriGetGLAccountTableColumnaData = False
					            Call fRptWriteReport("Fail","Get GL Account Column Data ","Data is not getting from GL Account table ")
				            	Exit Function
					        End If
					        Else
					        If strInnerText <> False Then 
				                ReDim preserve arrColData(intCount)
				                arrColData(intCount) = strInnerText
				                intCount = intCount + 1
				            End If
					     End If
				    End if     
				Next
			Else
			For intItearion = 6 To intRowCount
				If ObjPgGLAccount.WebElement("html id:=M0:46:::6:"&intColNum&"_l").Exist(1) then 
					If ObjPgGLAccount.WebElement("class:= urSvgAppIconMetric urSvgAppIconColorBase urSvgAppIconVAlign lsAbapList--image","visible:=True","html tag:=svg","Index:="&intItearion-6).Exist(1) then
					    	strInnerText = ObjPgGLAccount.WebElement("class:= urSvgAppIconMetric urSvgAppIconColorBase urSvgAppIconVAlign lsAbapList--image","visible:=True","html tag:=svg","Index:="&intItearion-6).GetROProperty("innerhtml")
					        If intItearion = 0 Then
					            intCount = 0
					            If Ucase(strInnerText) <> Ucase(strColumnName) Then
						            Call fRptWriteReport("Fail","Get GL Account Column Data ","Data is not getting from GL Account table ")
						            fFioriGetGLAccountTableColumnaData = False
					            	Exit Function 
					            End If
					        Else
					            If strInnerText <> False Then ' and strInnerText <> ""
					            	If Instr(1,Lcase(strInnerText),"iconnegative")>0 Then
					            		strInnerText="Open"
					            	ElseIf Instr(1,Lcase(strInnerText),"iconpositive")>0 Then
					            		strInnerText="Cleared"
									Else
										strInnerText="Error"
					            	End If
					                ReDim preserve arrColData(intCount)
					                arrColData(intCount) = strInnerText
					                intCount = intCount + 1
					            End If
					        End If
					   End if  
					  ' Get Last col data based on HTML ID  		
					  If intItearion =intRowCount Then
						   		If ObjPgGLAccount.WebElement("html id:=M0:46:::"&intRowCount&":6_l").Exist(1) then 
						   			If ObjPgGLAccount.WebElement("html id:=M0:46:::"&intRowCount&":6_l").Exist(1) then
						    		strInnerText = ObjPgGLAccount.WebElement("html id:=M0:46:::"&intRowCount&":6_l").GetROProperty("innertext")
						           		If strInnerText <> False Then ' and strInnerText <> "" 
							                ReDim preserve arrColData(intCount)
							                arrColData(intCount) = strInnerText
							                intCount = intCount + 1
						            	End If
						        End If
						   End if     
					   End If
					Else
					   fFioriGetGLAccountTableColumnaData = False
					    Call fRptWriteReport("Fail","Get GL Account Column Data ","Data is not getting from GL Account table ")
					   Call fRptWriteResultsSummary() 
					  Exit Function
					End IF			   
				Next
		End If		
		
		fFioriGetGLAccountTableColumnaData = arrColData
	    'Clear object
	    Set ObjPgGLAccount = Nothing
  	On error goto 0		
End Function


'***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriGLAccountTableLayoutChange
''	Objective						:				Used to change Fiori GL Account Table Layout
''	Input Parameters				:				srtColumnName,intColNum
''	Output Parameters			    :				NIL
''	Date Created					:				30/04/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*************************************************************************************************************************************************************************************** 
Public Function fFioriGLAccountTableLayoutChange(srtColumnName,intColNum)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	
	'Variable declaration
	Dim objGLFrame
	Dim intTableRowCount
	Dim strCellData
	Dim blnFound
	'Set Frame 
	Set objGLFrame = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
		If fnVerifyObjectExist(objGLFrame.WebTable("tblChangeLayout")) Then
			blnFound = False
			'Get table row count
			intTableRowCount = fnGetRoProperty(objGLFrame.WebTable("tblChangeLayout"),"rows","Change Layout")
				For intIteration = 1 To intTableRowCount
					'Get table cell data
					strCellData = trim(fnGetCelldata(objGLFrame.WebTable("tblChangeLayout"),intIteration,1,"Change Layout"))
						If srtColumnName = strCellData Then
								'Click on table cell
								If objGLFrame.WebTable("tblChangeLayout").ChildItem(intIteration,2,"WebList",0).exist(2) Then
									objGLFrame.WebTable("tblChangeLayout").ChildItem(intIteration,2,"WebList",0).click
								End If
							'Enter data in table cell
							Call fnEnterTextInCell(objGLFrame.WebTable("tblChangeLayout"),intIteration,2,"SAPEdit",0,intColNum)
							objGLFrame.SAPEdit("html id:=tbl.*"&intIteration&",2.*c").set intColNum
							blnFound = True
							Exit For
						End If
				Next
			If blnFound Then
				Call fRptWriteReport("Pass", "Verify GL Account Line Item Column","GL Account Line Item Column :"&srtColumnName&" found successfully")
				fFioriGLAccountTableLayoutChange = True
			Else
				Call fRptWriteReport("Fail", "Verify GL Account Line Item Column","GL Account Line Item Column :"&srtColumnName&" not found")
				fFioriGLAccountTableLayoutChange = False
				Exit Function
			End If
		Else
			Call fRptWriteReport("Fail", "Verify GL Account Line Item Column","GL Account Line Item table not found")
			fFioriGLAccountTableLayoutChange = False
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
	On error goto 0	
End Function

'***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriDisplayDocument
''	Objective						:				Used to verify data in Display Document page
''	Input Parameters				:				intDocumentNumber,intCompanyCode,dtFiscalYear
''	Output Parameters			    :				NIL
''	Date Created					:				30/04/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*************************************************************************************************************************************************************************************** 
Public Function fFioriDisplayDocument(intDocumentNumber,intCompanyCode,dtFiscalYear)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Variable declaration
	Dim objDisplayDocumentFrame
	Dim intFioriDocumentNumber
	Dim intFioriCompanyCode
	Dim dtFioriFiscalYear
	 'Set Frame 
    Set objDisplayDocumentFrame = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
	' Wait till SDocument Number txt field exist 
	Call fSynUntilObjExists(objDisplayDocumentFrame.SAPEdit("txtDocumentNumber"),MID_WAIT)
		If fVerifyObjectExist(objDisplayDocumentFrame.SAPEdit("txtDocumentNumber")) Then
			'Enter data in Document Number txt field
			Call fEnterText(objDisplayDocumentFrame.SAPEdit("txtDocumentNumber"),intDocumentNumber,"Document Number")
			'Enter data in Company Code txt field
			Call fEnterText(objDisplayDocumentFrame.SAPEdit("txtCompanyCode"),intCompanyCode,"Company Code")
			'Enter data in Fiscal Year txt field
			Call fEnterText(objDisplayDocumentFrame.SAPEdit("txtFiscalYear"),dtFiscalYear,"Fiscal Year")
			'Click on continue button
			Call fClick(objDisplayDocumentFrame.SAPButton("btnContinue"),"Continue")
			' Wait till SDocument Number txt field exist 
			Call fSynUntilObjExists(objDisplayDocumentFrame.SAPEdit("txtDocumentNumber"),MID_WAIT)
			'Get document number value
			intFioriDocumentNumber = fGetRoProperty(objDisplayDocumentFrame.SAPEdit("txtDocumentNumber"),"value","Document Number")
				If intDocumentNumber = intFioriDocumentNumber Then
					Call fRptWriteReport("Pass", "Verify Document Number in Document page","Document Number value "&intDocumentNumber&" has been displayed in Document page")
				Else
					Call fRptWriteReport("Fail", "Verify Document Number in Document page","Document Number has not been displayed in Document page")
				End If
			'Get Company Code value
			intFioriCompanyCode = fGetRoProperty(objDisplayDocumentFrame.SAPEdit("txtCompanyCode"),"value","Company Code")
				If intCompanyCode = intFioriCompanyCode Then
					Call fRptWriteReport("Pass", "Verify Company Code in Document page","Company Code value "&intCompanyCode&" has been displayed  in Document page")
				Else
					Call fRptWriteReport("Fail", "Verify Company Code in Document page","Company Code has not been displayed  in Document page")
				End If
			'Get Fiscal Year value
			dtFioriFiscalYear = fGetRoProperty(objDisplayDocumentFrame.SAPEdit("txtFiscalYear"),"value","Fiscal Year")
				If dtFiscalYear = dtFioriFiscalYear Then
					Call fRptWriteReport("Pass", "Verify Fiscal Year in Document page","Fiscal Year value "&dtFiscalYear&" has been displayed  in Document page")
				Else
					Call fRptWriteReport("Fail", "Verify Fiscal Year in Document page","Fiscal Year has not been displayed  in Document page")
					Call fRptWriteResultsSummary() 
					Exit Function
				End If
		Else
			Call fRptWriteReport("Fail", "Verify Fiori Display Document Number in Document page","Document Number has noyt been displayed in Document page")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
	On error goto 0	
End Function
'***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriDisplayAsset
''	Objective						:				Used to Verify data in Display Asset page
''	Input Parameters				:				intAsset,strClass,intCompanyCode,intAccountDeterm,intCostCenter,intVendor,arrAreaNumber,arDepreciationArea,dtYear
''	Output Parameters			    :				NIL
''	Date Created					:				01/05/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*************************************************************************************************************************************************************************************** 
Public Function fFioriDisplayAsset(intAsset,intSubNumber,strClass,intCompanyCode,intAccountDeterm,intCostCenter,intVendor,arrAreaNumber,arDepreciationArea,dtYear)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Variable declaration
	Dim objDisplayAssetFrame
	Dim intFioriAsset
	Dim strFioriClass
	Dim intFioriCompanyCode
	Dim strFioriDescription
	Dim intFioriAccountDeterm
	Dim intFioriCostCenter
	Dim intFioriVendor
	Dim intTableRowCount
	Dim arrFioriAreaNumber
	Dim arrDepreciationArea
	Dim intIteration
	Dim blnFound
	Dim intIterator
	Dim strCellData
	Dim strDACellData
	Dim dtFioriFiscalYear
	'Set Display Asset frame
	Set objDisplayAssetFrame = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
	' Wait till Asset txt field exist 
	Call fSynUntilObjExists(objDisplayAssetFrame.SAPEdit("txtAsset"),MID_WAIT)
		If fVerifyObjectExist(objDisplayAssetFrame.SAPEdit("txtAsset")) Then
			'Enter data in Asset txt field
			Call fEnterText(objDisplayAssetFrame.SAPEdit("txtAsset"),intAsset,"Asset")
			'Click on Main Asset Number button
			Call fClick(objDisplayAssetFrame.SAPButton("btnMainAssetNumber"),"Main Asset Number")
			'Select Asset number
			Call fClick(objDisplayAssetFrame.WebTable("tblMainAssetNumber").ChildItem(1,1,"WebButton",0),"Asset Number")
			'Click on Copy button
			Call fClick(objDisplayAssetFrame.SAPButton("btnCopy"),"Copy")
			' Wait till Sub Number txt field exist 
			Call fSynUntilObjExists(objDisplayAssetFrame.SAPEdit("txtSubNumber"),MID_WAIT)
			'Enter data in Sub Number txt field
			Call fEnterText(objDisplayAssetFrame.SAPEdit("txtSubNumber"),intSubNumber,"Sub Number")
			'Enter data in Company Code txt field
			Call fEnterText(objDisplayAssetFrame.SAPEdit("txtCompanyCode"),intCompanyCode,"Company Code")
			'Click on Master Data button
			Call fClick(objDisplayAssetFrame.SAPButton("btnMasterData"),"Master Data")
			' Wait till SDocument Number txt field exist 
			Call fSynUntilObjExists(objDisplayAssetFrame.SAPEdit("txtAsset"),MID_WAIT)
			'Get Asset value 
			intFioriAsset = fGetRoProperty(objDisplayAssetFrame.SAPEdit("txtAsset"),"value","Asset")
				If intAsset = intFioriAsset Then
					Call fRptWriteReport("Pass", "Verify Asset in Display Asset page","Asset value "&intAsset&" has been displayed in General tab of Display Asset page")
				Else
					Call fRptWriteReport("Fail", "Verify Asset in Display Asset page","Asset has not been displayed in General tab of Display Asset page")
				End If
			'Get Class value 
			strFioriClass = fGetRoProperty(objDisplayAssetFrame.SAPEdit("txtClass"),"value","Class")
				If strClass = strFioriClass Then
					Call fRptWriteReport("Pass", "Verify Class in Display Asset page","Class value "&strClass&" has been displayed in General tab of Display Asset page")
				Else
					Call fRptWriteReport("Fail", "Verify Class in Display Asset page","Class has not been displayed in General tab of Display Asset page")
				End If
			'Get Company Code value 
			intFioriCompanyCode = fGetRoProperty(objDisplayAssetFrame.SAPEdit("txtCompanyCode"),"value","Company Code")
				If intCompanyCode = intFioriCompanyCode Then
					Call fRptWriteReport("Pass", "Verify Company Code in Display Asset page","Company Code value "&intCompanyCode&" has been displayed in General tab of Display Asset page")
				Else
					Call fRptWriteReport("Fail", "Verify Company Code in Display Asset page","Company Code has not been displayed in General tab of Display Asset page")
				End If
			'Get Description value 
			strFioriDescription = fGetRoProperty(objDisplayAssetFrame.SAPEdit("txtDescription"),"value","Description")
				If strClass = strFioriDescription Then
					Call fRptWriteReport("Pass", "Verify Description in Display Asset page","Description "&strClass&" has been displayed in General tab of Display Asset page")
				Else
					Call fRptWriteReport("Fail", "Verify Description in Display Asset page","Description has not been displayed in General tab of Display Asset page")
				End If
			'Get Description value 
			intFioriAccountDeterm = fGetRoProperty(objDisplayAssetFrame.SAPEdit("txtAccountDeterm"),"value","Account Determ in General tab")
				If intAccountDeterm = intFioriAccountDeterm Then
					Call fRptWriteReport("Pass", "Verify Account Determ in Display Asset page","Account Determ value "&intAccountDeterm&" has been displayed in General tab of Display Asset page")
				Else
					Call fRptWriteReport("Fail", "Verify Account Determ in Display Asset page","Account Determ has not been displayed of Display Asset page")
				End If
			'Select Time-dependent
			Call fSelect(objDisplayAssetFrame.WebTabStrip("wtsDisplayAsset"),"Time-dependent","Time-dependent")
			' Wait till Cost Center txt field exist 
			Call fSynUntilObjExists(objDisplayAssetFrame.SAPEdit("txtCostCenter"),MID_WAIT)
			'Get Cost Center value 
			intFioriCostCenter = fGetRoProperty(objDisplayAssetFrame.SAPEdit("txtCostCenter"),"value","Cost Center")
				If intCostCenter = intFioriCostCenter Then
					Call fRptWriteReport("Pass", "Verify Cost Center in Display Asset page","Cost Center value "&intCostCenter&" has been displayed in Time-dependent tab of Display Asset page")
				Else
					Call fRptWriteReport("Fail", "Verify Cost Center in Display Asset page","Cost Center has not been displayed in Time-dependent tab of Display Asset page")
				End If
			'Select Origin
			Call fSelect(objDisplayAssetFrame.WebTabStrip("wtsDisplayAsset"),"Origin","Origin")
			' Wait till Cost Center txt field exist 
			Call fSynUntilObjExists(objDisplayAssetFrame.SAPEdit("txtVendor"),MID_WAIT)
			'Get Cost Center value 
			intFioriVendor = fGetRoProperty(objDisplayAssetFrame.SAPEdit("txtVendor"),"value","Vendor")
				If intVendor = intFioriVendor Then
					Call fRptWriteReport("Pass", "Verify Vendor in Display Asset page","Vendor value "&intVendor&" has been displayed in Origin tab of Display Asset page")
				Else
					Call fRptWriteReport("Fail", "Verify Vendor in Display Asset page","Vendor has not been displayed in Origin tab of Display Asset page")
				End If
			'Select Origin
			Call fSelect(objDisplayAssetFrame.WebTabStrip("wtsDisplayAsset"),"Deprec. Areas","Deprec. Areas")
			' Wait till Display Asset table exist 
			Call fSynUntilObjExists(objDisplayAssetFrame.WebTable("tblDisplayAsset"),MID_WAIT)
			'Get Display Asset table row count
			intTableRowCount = fGetRoProperty(objDisplayAssetFrame.WebTable("tblDisplayAsset"),"rows","Display Asset")
			arrFioriAreaNumber = Split(arrAreaNumber,"@")
			arrDepreciationArea = Split(arDepreciationArea,"@")
				For intIteration = 0 To Ubound(arrFioriAreaNumber) - 1
					blnFound = False
						For intIterator = 1 To intTableRowCount
							'Get Area Number cell data
							strCellData = fGetCelldata(objDisplayAssetFrame.WebTable("tblDisplayAsset"),intIterator,2,"Display Asset")
								If strCellData = arrFioriAreaNumber(intIteration) Then
									'Get Depreciation Area cell data
									strDACellData = fGetCelldata(objDisplayAssetFrame.WebTable("tblDisplayAsset"),intIterator,3,"Display Asset")
										If strDACellData = arrDepreciationArea(intIteration) Then
											blnFound = True
											Exit For
										End If'
								End If	
						Next
					If blnFound Then
						Call fRptWriteReport("Pass", "Verify Area Number and Depreciation Area in Display Asset page","Area Number value "&arrFioriAreaNumber(intIteration)&" and Depreciation Area value "&arrDepreciationArea(intIteration)&" has been displayed in Deprec Areas tab of Display Asset page")
					Else
						Call fRptWriteReport("Fail", "Verify Area Number and Depreciation Area in Display Asset page","Area Number and Depreciation Area has not been displayed in Deprec Areas tab of Display Asset page")
					End If
				Next
			Call fClick(objDisplayAssetFrame.SAPButton("txtAssetValues"),"Asset Values")
			' Wait till Display Asset table exist 
			Call fSynUntilObjExists(objDisplayAssetFrame.WebElement("weDepreciationAreas"),MID_WAIT)
			'Click on Depreciation Areas element
			Call fClick(objDisplayAssetFrame.WebElement("weDepreciationAreas"),"Depreciation Areas")
			'Get Company Code value 
			intFioriCompanyCode = fGetRoProperty(objDisplayAssetFrame.SAPEdit("txtCompanyCode"),"value","Company Code")
				If intCompanyCode = intFioriCompanyCode Then
					Call fRptWriteReport("Pass", "Verify Company Code in Display Asset page","Company Code value "&intCompanyCode&" has been displayed in Asset page")
				Else
					Call fRptWriteReport("Fail", "Verify Company Code in Display Asset page","Company Code has not been displayed in Asset page")
				End If
			'Get Asset value 
			intFioriAsset = fGetRoProperty(objDisplayAssetFrame.SAPEdit("txtAsset"),"value","Asset")
				If intAsset = intFioriAsset Then
					Call fRptWriteReport("Pass", "Verify Asset in Display Asset page","Asset value "&intAsset&" has been displayed in Asset page")
				Else
					Call fRptWriteReport("Fail", "Verify Asset in Display Asset page","Asset has not been displayed in Asset page")
				End If
			'Get Fiscal Year value 
			dtFioriFiscalYear = fGetRoProperty(objDisplayAssetFrame.SAPEdit("txtFiscalYear"),"value","Fiscal Year")
				If dtYear = dtFioriFiscalYear Then
					Call fRptWriteReport("Pass", "Verify Fiscal Year in Display Asset page","Fiscal Year value "&dtYear&" has been displayed in Asset page")
				Else
					Call fRptWriteReport("Fail", "Verify Fiscal Year in Display Asset page","Fiscal Year has not been displayed in Asset page")
					Call fRptWriteResultsSummary() 
					Exit Function
				End If
		Else
			Call fRptWriteReport("Fail", "Verify Asset field","Asset field dosn't exist")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
	On error goto 0
End Function
'***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriManagePaymentBlockSupplier
''	Objective						:				Used to  Block Supplier
''	Input Parameters				:				strPaymentBlockReason,strPaymentBlockReasonNotes,strStatus
''	Output Parameters			    :				NIL
''	Date Created					:				04/05/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'***************************************************************************************************************************************************************************************  
 Public Function fFioriManagePaymentBlockSupplier(strPaymentBlockReason,strPaymentBlockReasonNotes,strStatus)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	'Variable declaration
 	Dim objManagePaymentBlocks
 	Dim intTableRowCount
 	Dim blnFoun
 	Dim intIteration
 	'Set Browser and Page
 	Set objManagePaymentBlocks = Browser("brFiori").Page("pgFiori")
 	'Used to Select Manage Payment Block
 	Call fFioriSelectManagePaymentBlock()
		If fVerifyObjectExist(objManagePaymentBlocks.SAPUIButton("btnBlockSupplier")) Then
			'Click on Block Supplier
			Call fClick(objManagePaymentBlocks.SAPUIButton("btnBlockSupplier"),"Block Supplier")
			'Used to Fill Block Payment Details
			Call fFioriFillBlockPaymentDetails(strPaymentBlockReason,strPaymentBlockReasonNotes)
			' Wait till Manage Payment Blocks Reason field exist
			Call fSynUntilObjExists(objManagePaymentBlocks.WebElement("weAccountBlocked"),MAX_WAIT)
				'Verify Supplier Blocked on left side of page
				If fVerifyObjectExist(objManagePaymentBlocks.WebElement("weAccountBlocked")) Then
					Call fRptWriteReport("Pass", "Verify Supplier Blocked on left side of Supplier page","Supplier has been blocked successfully in Supplier page")
				Else
					Call fRptWriteReport("Fail",  "Verify Supplier Blocked on left side of Supplier page","Supplier has not been blocked in Supplier page")
				End If
				'Verify Supplier Blocked on right side of  page
				If fVerifyObjectExist(objManagePaymentBlocks.WebElement("weAccountBlockedOfHeader")) Then
					Call fRptWriteReport("Pass", "Verify Supplier Blocked on right side of Supplier page","Supplier has been blocked successfully in Supplier page")
				Else
					Call fRptWriteReport("Fail",  "Verify Supplier Blocked on right side of Supplier page","Supplier has not been blocked in Supplier page")
				End If
			' Get table row count	
			intTableRowCount = fGetRoProperty(objManagePaymentBlocks.SAPUITable("tblManagePaymentBlocks"),"row count","Manage Payment Blocks")
			blnFoun = False
					'Verify Supplier Blocked on table
					For intIteration = 1 To intTableRowCount
						If strStatus = fGetCelldata(objManagePaymentBlocks.SAPUITable("tblManagePaymentBlocks"),intIteration,2,"Manage Payment Blocks") Then
							blnFoun = True
						Else
							blnFoun = False
							Exit For
						End If
					Next
				If blnFoun Then
					Call fRptWriteReport("Pass", "Verify Supplier Blocked in Supplier page","Supplier has been blocked successfully in Supplier page")
				Else
					Call fRptWriteReport("Fail",  "Verify Supplier Blocked in Supplier page","Supplier has not been blocked in Supplier page")
					Call fRptWriteResultsSummary() 
					Exit Function
				End If
		Else
			Call fRptWriteReport("Fail","Verify Supplier Blocked in Supplier page","Block Supplierd button not displayed")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
	On error goto 0	
 End Function
  '***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriManagePaymentUnBlockSupplier
''	Objective						:				Used to Un-Block Supplier
''	Input Parameters				:				strStatus
''	Output Parameters			    :				NIL
''	Date Created					:				04/05/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'***************************************************************************************************************************************************************************************  
 Public Function fFioriManagePaymentUnBlockSupplier(strStatue)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	'Variable declaration
 	Dim objManagePaymentBlocks
 	Dim intTableRowCount
 	Dim blnFoun
 	Dim intIteration
 	'Set Browser and Page
 	Set objManagePaymentBlocks = Browser("brFiori").Page("pgFiori")
 	'Used to Select Manage Payment Block
 	Call fFioriSelectManagePaymentBlock()
		If fVerifyObjectExist(objManagePaymentBlocks.SAPUIButton("btnUnblockSupplier")) Then
			'Click on un-Block Supplier
			Call fClick(objManagePaymentBlocks.SAPUIButton("btnUnblockSupplier"),"Un-Block Supplier")
			Wait(MIN_WAIT)'Required
				'Verify Supplier Blocked on left side of page
				If not fVerifyObjectExist(objManagePaymentBlocks.WebElement("weAccountBlocked")) Then
					Call fRptWriteReport("Pass", "Verify Supplier un-Blocked on left side of Supplier page","Supplier has been un-blocked successfully in Supplier page")
				Else
					Call fRptWriteReport("Fail",  "Verify Supplier un-Blocked on left side of Supplier page","Supplier has not been un-blocked in Supplier page")
				End If
				'Verify Supplier Blocked on right side of  page
				If not fVerifyObjectExist(objManagePaymentBlocks.WebElement("weAccountBlockedOfHeader")) Then
					Call fRptWriteReport("Pass", "Verify Supplier un-Blocked on right side of Supplier page","Supplier has been un-blocked successfully in Supplier page")
				Else
					Call fRptWriteReport("Fail",  "Verify Supplier un-Blocked on right side of Supplier page","Supplier has not been un-blocked in Supplier page")
				End If
			' Get table row count	
			intTableRowCount = fGetRoProperty(objManagePaymentBlocks.SAPUITable("tblManagePaymentBlocks"),"row count","Manage Payment Blocks")
			blnFoun = False
					'Verify Supplier Blocked on table
					For intIteration = 1 To intTableRowCount
						If strStatue = trim(fGetCelldata(objManagePaymentBlocks.SAPUITable("tblManagePaymentBlocks"),intIteration,2,"Manage Payment Blocks")) Then
							blnFoun = True
						Else
							blnFoun = False
							Exit For
						End If
					Next
				If blnFoun Then
					Call fRptWriteReport("Pass", "Verify Supplier un-Blocked in supplier page","Supplier has been un-blocked successfully in supplier page")
				Else
					Call fRptWriteReport("Fail",  "Verify Supplier un-Blocked in supplier page","Supplier has not been blocked in supplier page")
					Call fRptWriteResultsSummary() 
					Exit Function
				End If
		Else
			Call fRptWriteReport("Fail","Verify un-Block Supplier in supplier page","un-Block Supplierd button not displayed in supplier page")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
	On error goto 0
 End Function
 
'***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriFillBlockPaymentDetails
''	Objective						:				Used to Fill Block Payment Details
''	Input Parameters				:				strPaymentBlockReason,strPaymentBlockReasonNotes
''	Output Parameters			    :				NIL
''	Date Created					:				04/05/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*************************************************************************************************************************************************************************************** 
 Public Function fFioriFillBlockPaymentDetails(strPaymentBlockReason,strPaymentBlockReasonNotes)
 	On error resume next
 	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	'Variable declaration
 	Dim objManagePaymentBlocks
 	'Set Browser and Page
 	Set objManagePaymentBlocks = Browser("brFiori").Page("pgFiori")
 	' Wait till Manage Payment Blocks Reason field exist
	Call fSynUntilObjExists(objManagePaymentBlocks.SAPUIMenu("sumPaymentBlockReason"),MAX_WAIT)
		If fVerifyObjectExist(objManagePaymentBlocks.SAPUIMenu("sumPaymentBlockReason")) Then
			'Select Payment Block Reason
			Call fSelect(objManagePaymentBlocks.SAPUIMenu("sumPaymentBlockReason"),strPaymentBlockReason,"Payment Block Reason")
			'Enter text in Payment Block Reason Notes
			Call fEnterText(objManagePaymentBlocks.SAPUITextEdit("txtPaymentBlockReasonNotes"),strPaymentBlockReasonNotes,"Payment Block Reason Notes")
			'Click on OK button
			Call fClick(objManagePaymentBlocks.SAPUIButton("btnOK"),"OK")
		Else
			Call fRptWriteReport("Fail", "Verify Payment Block Reason in Payment Block page","Payment Block Reason menu not displayed")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
	On error goto 0
 End Function
 
 '***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriSelectManagePaymentBlock
''	Objective						:				Used to Select Manage Payment Block
''	Input Parameters				:				NIL
''	Output Parameters			    :				NIL
''	Date Created					:				04/05/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*************************************************************************************************************************************************************************************** 
  Public Function fFioriSelectManagePaymentBlock()
  	On error resume next
  	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	'Variable declaration
 	Dim objAccountsPayableOverview
 	'Set Browser and Page
 	Set objAccountsPayableOverview = Browser("brFiori").Page("pgFiori")
 	' Wait till Manage Payment Blocks objext exist
	Call fSynUntilObjExists(objAccountsPayableOverview.WebElement("weManagePaymentBlocks"),MAX_WAIT)
		If fVerifyObjectExist(objAccountsPayableOverview.WebElement("weManagePaymentBlocks")) Then
			'Click on manage Payment Blocks
			Call fClick(objAccountsPayableOverview.WebElement("weManagePaymentBlocks"),"Manage Payment Blocks")
			'Wait till Manage Payment Blocks table exist
			Call fSynUntilObjExists(objAccountsPayableOverview.SAPUITable("tblManagePaymentBlocks"),MAX_WAIT)
			'Select checkbox in table
			Call fSelectRowInTable(objAccountsPayableOverview.SAPUITable("tblManagePaymentBlocks"),1,"Manage Payment Blocks")
		Else
			Call fRptWriteReport("Fail", "Verify Manage Payment Blocks in Payment page","Manage Payment Blocks button is not displayed in Payment page")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
	On error goto 0
 End Function


'***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriManagePaymentBlock
''	Objective						:				Used to  Block Payment
''	Input Parameters				:				strPaymentBlockReason,strPaymentBlockReasonNotes,strStatus
''	Output Parameters			    :				NIL
''	Date Created					:				05/05/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'***************************************************************************************************************************************************************************************  
 Public Function fFioriManagePaymentBlock(strPaymentBlockReason,strPaymentBlockReasonNotes,strStatus)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	'Variable declaration
 	Dim objManagePaymentBlocks
 	Dim intIteration
 	'Set Browser and Page
 	Set objManagePaymentBlocks = Browser("brFiori").Page("pgFiori")
 	'Used to Select Manage Payment Block
 	Call fFioriSelectManagePaymentBlock()
		If fVerifyObjectExist(objManagePaymentBlocks.SAPUIButton("btnBlockItem")) Then
			'Click on Block Supplier
			Call fClick(objManagePaymentBlocks.SAPUIButton("btnBlockItem"),"Block Item")
			'Used to Fill Block Payment Details
			Call fFioriFillBlockPaymentDetails(strPaymentBlockReason,strPaymentBlockReasonNotes)
			Wait(2)'Required
				'Verify payment blocked item
				If strStatus = fGetCelldata(objManagePaymentBlocks.SAPUITable("tblManagePaymentBlocks"),1,2,"Manage Payment Blocks") Then
					Call fRptWriteReport("Pass", "Verify Payment Blocked in Payment page","Payment has been blocked successfully in Payment page")
				Else
					Call fRptWriteReport("Fail",  "Verify Payment Blocked in Payment page","Payment has not been blocked  in Payment page")
					Call fRptWriteResultsSummary() 
					Exit Function
				End If
		Else
			Call fRptWriteReport("Fail","Verify Block Supplier in Payment page","Block Item button not displayed  in Payment page")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
	On error goto 0	
 End Function
 
 '***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriManagePaymentUnBlock
''	Objective						:				Used to Un-Block Payment
''	Input Parameters				:				strStatus
''	Output Parameters			    :				NIL
''	Date Created					:				05/05/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'***************************************************************************************************************************************************************************************  
 Public Function fFioriManagePaymentUnBlock(strStatue)
	On error resume next
 	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	'Variable declaration 	
 	Dim objManagePaymentBlocks
 	Dim intTableRowCount
 	Dim blnFoun
 	Dim intIteration
 	'Set Browser and Page
 	Set objManagePaymentBlocks = Browser("brFiori").Page("pgFiori")
 	'Used to Select Manage Payment Block
 	Call fFioriSelectManagePaymentBlock()
		If fVerifyObjectExist(objManagePaymentBlocks.SAPUIButton("btnUnblockItem")) Then
			'Click on un-Block Supplier
			Call fClick(objManagePaymentBlocks.SAPUIButton("btnUnblockItem"),"Unblock Item")
			Wait(MIN_WAIT)'Required
				'Verify Supplier Blocked on table
				If strStatue = trim(fGetCelldata(objManagePaymentBlocks.SAPUITable("tblManagePaymentBlocks"),1,2,"Manage Payment Blocks")) Then
					Call fRptWriteReport("Pass", "Verify Payment un-Blocked in Payment page","Payment has been un-blocked successfully in Payment page")
				Else
					Call fRptWriteReport("Fail",  "Verify Payment un-Blocked in Payment page","Payment has not been blocked in Payment page")
					Call fRptWriteResultsSummary() 
					Exit Function
				End If
		Else
			Call fRptWriteReport("Fail","Verify Payment un-Block in Payment page","Payment un-Block button not displayed in Payment page")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
	On error goto 0	
 End Function 
  '***************************************************************************************************************************************************************************************
''	Function Name					:				fFioriManageSupplierLineItem
''	Objective						:				Used to Manage Supplier LineItem
''	Input Parameters				:				strSupplier
''	Output Parameters			    :				NIL
''	Date Created					:				06/05/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'***************************************************************************************************************************************************************************************  
 Public Function fFioriManageSupplierLineItem(strSupplier)
  	On error resume next
  	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
 	'Variable declaration
 	Dim objAccountsPayableOverview
 	Call fSynUntilObjExists(Browser("brFiori").Page("pgFiori"),MIN_WAIT)
 	'Set Browser and Page
 	Set objAccountsPayableOverview = Browser("brFiori").Page("pgFiori")
 	' Wait till Manage Payment Blocks objext exist
	Call fSynUntilObjExists(objAccountsPayableOverview.WebElement("weManagePaymentBlocks"),MAX_WAIT)
		If fVerifyObjectExist(objAccountsPayableOverview.WebElement("weManagePaymentBlocks")) Then
			'Click on manage Payment Blocks
			Call fClick(objAccountsPayableOverview.WebElement("weManagePaymentBlocks"),"Manage Payment Blocks")
			'Wait till manage Supplier Line Items exist
			Call fSynUntilObjExists(objAccountsPayableOverview.SAPUITable("tblManagePaymentBlocks"),MAX_WAIT)
			'Click on Autodesk Asia Pte Ltd
			Call fClick(objAccountsPayableOverview.WebElement("weAutodeskAsiaPteLtd"),"Autodesk Asia Pte Ltd")
			'Click on GL Account
			objAccountsPayableOverview.SAPUITable("tblManagePaymentBlocks").ChildItem(1,3,"WebElement",0).click
			'Wait till manage Supplier Line Items exist
			Call fSynUntilObjExists(objAccountsPayableOverview.WebButton("btnAutodeskAsiaPteLtd"),MAX_WAIT)
			'Click on Autodesk Asia Pte Ltd
			Call fClick(objAccountsPayableOverview.WebButton("btnAutodeskAsiaPteLtd"),"Autodesk Asia Pte Ltd")
			'Wait till Manage Supplier Line Items exist
			Call fSynUntilObjExists(objAccountsPayableOverview.Link("lnkManageSupplierLineItems"),MAX_WAIT)
			'Click on Manage Supplier Line Items
			Call fClick(objAccountsPayableOverview.Link("lnkManageSupplierLineItems"),"Manage Supplier Line Items")
			'Wait till Manage Supplier Line Items exist
			Call fSynUntilObjExists(objAccountsPayableOverview.WebTable("tblManageSupplierLineItem"),MAX_WAIT)
			Wait(2)'Required
				'Verify payment blocked item
				If strSupplier = fGetCelldata(objAccountsPayableOverview.WebTable("tblManageSupplierLineItem"),2,1,"Manage Supplier Line Item") Then
					Call fRptWriteReport("Pass", "Verify Supplier in Supplier page","Supplier "&strSupplier&" has been displayed successfully in Supplier page")
				Else
					Call fRptWriteReport("Fail",  "Verify Supplier in Supplier page","Supplier "&strSupplier&" has not been displayed in Supplier page")
					Call fRptWriteResultsSummary() 
					Exit Function
				End If
		Else
			Call fRptWriteReport("Fail", "Verify Manage Payment Blocks button in Supplier page","Manage Payment Blocks button not displayed  in Supplier page")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
	On error goto 0	
 End Function

'******************************************************************************************************************************************************************************************************************************************
''	Function Name					:				fFioriNavigationFromHomePage
''	Objective						:				Navigation from home page
''	Input Parameters				:				objDataDict,iRowCountRef
''	Output Parameters			    :				Nil
''	Date Created					:				27/April/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*************************************************************************************************************************************************************************************** 
Public Function fFioriNavigationFromHomePage(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim ObjPgHome
	Dim ObjPgGLAccount
	Dim strInfoMsg
	Dim ObjPg
	Dim strTitleIcon
	Dim strTitlePageName
	'Input data
		strTitleIcon = objDataDict.Item("HomePageTitleIcon" & iRowCountRef)
		strTitlePageName = objDataDict.Item("PageTitleName" & iRowCountRef)
		
		Call fSynUntilObjExists(Browser("brFIORI").Page("pgFiori"),MIN_WAIT)
		Set ObjPgHome=Browser("brFIORI").Page("pgFiori")
		Set ObjPgGLAccount=Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame")
		Set ObjPg=Browser("brFioriAutoDesk").Page("pgGLAccountLineItem")
		'Click on G/L Account Line Item under General display section
		Call fClick(ObjPgHome.WebButton(strTitleIcon),strTitlePageName)
		'Get text and Compare
		strInfoMsg = Lcase(fGetText(ObjPg.WebElement("wePageHeaderName"),"innertext","Page Title"))
			If instr(1,strInfoMsg,Lcase(strTitlePageName)) Then
				Call fRptWriteReport("Pass" ,"Verify opened Page",Ucase(strInfoMsg) &" -Page opened successfully")	 
			Else
				Call fRptWriteReport("Fail" ,"Verify opened Page",Ucase(strInfoMsg) &" -Page is not opened")
				Call fRptWriteResultsSummary() 
				Exit Function
			End If
		'Clear object details
		Set ObjPgHome = Nothing
		Set ObjPgGLAccount = Nothing
		Set ObjPg = Nothing
	
On error goto 0	
End Function
'***************************************************************************************************************************************
''	Function Name					:				fFioriFillGLAccountSelection
''	Objective						:				Fill details in G/L Account Selection
''	Input Parameters				:				strGLAccountFrom,strGLAccountTo,strCompanyCodeFrom,strCompanyCodeTo
''	Output Parameters			    :				Nil
''	Date Created					:				28/April/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fFioriFillGLAccountSelection(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim ObjPgGLAccount
	Dim strGLAccountFrom
	Dim strGLAccountTo
	Dim strCompanyCodeFrom
	Dim strCompanyCodeTo
		'Get Test data
		strGLAccountFrom =  objDataDict.Item("GLAccountFrom" & iRowCountRef)
		strGLAccountTo =  objDataDict.Item("GLAccountTo" & iRowCountRef)
		strCompanyCodeFrom =  objDataDict.Item("CompanyCodeFrom" & iRowCountRef)
		strCompanyCodeTo = objDataDict.Item("CompanyCodeTo" & iRowCountRef)
		Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame"),MIN_WAIT)
		'Set Object
		Set ObjPgGLAccount=Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame")
		'Enter G/L account from text field 
			If Lcase(strGLAccountFrom)<> "n/a" and Lcase(strGLAccountFrom)<> "null" and Lcase(strGLAccountFrom)<> "no" Then
				Call fEnterText(ObjPgGLAccount.WebEdit("txtGLAccountFrom"),strGLAccountFrom,"G/L account from")
			End If
			'Enter G/L account to text field 
			If Lcase(strGLAccountTo)<> "n/a" and Lcase(strGLAccountTo)<> "null" and Lcase(strGLAccountTo)<> "no" Then
				Call fEnterText(ObjPgGLAccount.WebEdit("txtGLAccountTo"),strGLAccountTo,"Company Code To")
			End If	
			'Enter Company Code From
			If Lcase(strCompanyCodeFrom)<> "n/a" and Lcase(strCompanyCodeFrom)<> "null" and Lcase(strCompanyCodeFrom)<> "no" Then
				Call fEnterText(ObjPgGLAccount.WebEdit("txtCompanyCodeFrom"),strCompanyCodeFrom,"Company Code From")
			End If
			'Enter Company Code To
			If Lcase(strCompanyCodeTo)<> "n/a" and Lcase(strCompanyCodeTo)<> "null" and Lcase(strCompanyCodeTo)<> "no" Then
				Call fEnterText(ObjPgGLAccount.WebEdit("txtCompanyCodeTo"),strCompanyCodeTo,"Company Code To")
			End If	
		'Clear object
		Set ObjPgGLAccount= Nothing
	On error goto 0			
End Function

'***************************************************************************************************************************************
''	Function Name					:				fFioriFillLineItemSelection
''	Objective						:				Fill details in Line Item Selection
''	Input Parameters				:				strLineItemType,strOpenAtKeyDate_OpenItems,strClearingDateFrom
													'strClearedDateTo,strOpenAtKeyDate_ClosedItems,strPostingDateFrom,strPostingDateTo
''	Output Parameters			    :				Nil
''	Date Created					:				28/April/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fFioriFillLineItemSelection(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim ObjPgGLAccount
	Dim strLineItemType
	Dim strOpenAtKeyDate_OpenItems
	Dim strClearingDateFrom
	Dim strClearedDateTo
	Dim strOpenAtKeyDate_ClosedItems
	Dim strPostingDateFrom
	Dim strPostingDateTo
		'Get test data	
		strLineItemType = objDataDict.Item("LineItemType" & iRowCountRef)
		strOpenAtKeyDate_OpenItems = objDataDict.Item("OpenAtKeyDate_OpenItems" & iRowCountRef)
		strClearingDateFrom = objDataDict.Item("ClearingDateFrom" & iRowCountRef)
		strClearedDateTo = objDataDict.Item("ClearedDateTo" & iRowCountRef)
		strOpenAtKeyDate_ClosedItems = objDataDict.Item("OpenAtKeyDate_ClosedItems" & iRowCountRef)
		strPostingDateFrom = objDataDict.Item("PostingDateFrom" & iRowCountRef)
		strPostingDateTo = objDataDict.Item("PostingDateTo" & iRowCountRef)
		
		Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame"),MIN_WAIT)
		Set ObjPgGLAccount=Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame")		
		Select Case strLineItemType
			Case Lcase("openitems")
				'Select Open Items option
				Call fSelect(ObjPgGLAccount.SAPRadioGroup("rbLineItemStatus"),"Open Items","Open Items")
				strOpenAtKeyDate_OpenItems=fGetFutureDateAdd(strOpenAtKeyDate_OpenItems)
			 	Call fEnterText(ObjPgGLAccount.WebEdit("txtOpenAtKeyDate_OpenItems"),strOpenAtKeyDate_OpenItems,"Open At Key Date")
				
			Case Lcase("cleareditems")
				'Select Cleared Items option
				Call fSelect(ObjPgGLAccount.SAPRadioGroup("rbLineItemStatus"),"Cleared Items","Cleared Items")
				'Enter Clearing Date From
				If Lcase(strClearingDateFrom)<> "n/a" and Lcase(strClearingDateFrom)<> "null" and Lcase(strClearingDateFrom)<> "no" Then
					strClearingDateFrom=fGetFutureDateAdd(strClearingDateFrom)
					Call fEnterText(ObjPgGLAccount.WebEdit("txtClearingDateFrom"),strClearingDateFrom,"Clearing Date From")
				End If
				'Enter Clearing Date To
				If Lcase(strClearedDateTo)<> "n/a" and Lcase(strClearedDateTo)<> "null" and Lcase(strClearedDateTo)<> "no" Then
					strClearedDateTo=fGetFutureDateAdd(strClearedDateTo)
					Call fEnterText(ObjPgGLAccount.WebEdit("txtClearedDateTo"),strClearedDateTo,"Cleared Date To")
				End If
				
				'Enter Open At Key Date
				If Lcase(strOpenAtKeyDate_ClosedItems)<> "n/a" and Lcase(strOpenAtKeyDate_ClosedItems)<> "null" and Lcase(strOpenAtKeyDate_ClosedItems)<> "no" Then
					strOpenAtKeyDate_ClosedItems=fGetFutureDateAdd(strOpenAtKeyDate_ClosedItems)
					Call fEnterText(ObjPgGLAccount.WebEdit("txtOpenAtKeyDate_ClosedItems"),strOpenAtKeyDate_ClosedItems,"Open At Key Date")
				End  IF
				
			Case Lcase("allitems")
				'Select All Items option
				Call fSelect(ObjPgGLAccount.SAPRadioGroup("rbLineItemStatus"),"All Items","All Items")
				'Enter Posting Date From
				If Lcase(strPostingDateFrom)<> "n/a" and Lcase(strPostingDateFrom)<> "null" and Lcase(strPostingDateFrom)<> "no" Then
					strPostingDateFrom=fGetFutureDateAdd(strPostingDateFrom)
					Call fEnterText(ObjPgGLAccount.WebEdit("txtPostingDateFrom"),strPostingDateFrom,"Posting Date From")
				End If	
				
				'Enter Posting Date To
				If Lcase(strPostingDateTo)<> "n/a" and Lcase(strPostingDateTo)<> "null" and Lcase(strPostingDateTo)<> "no" Then
					strPostingDateTo=fGetFutureDateAdd(strPostingDateTo)
					Call fEnterText(ObjPgGLAccount.WebEdit("txtPostingDateTo"),strPostingDateTo,"Posting Date To")
				End If	
		End Select
		'Clear object
		Set ObjPgGLAccount = Nothing
	On error goto 0				
End Function	
'***************************************************************************************************************************************
''	Function Name					:				fFioriFillListOutput
''	Objective						:				Fill details in List output Selection
''	Input Parameters				:				strLayout,strMaxiumNoOfItems
''	Output Parameters			    :				Nil
''	Date Created					:				28/April/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fFioriFillListOutput(objDataDict,iRowCountRef)
		On error resume next
		'Verify if Step Failed, If yes, it will not run the function
	    If Environment("StepFailed") = "YES" Then
			Exit Function
		End If	
		'Declarations
		Dim ObjPgGLAccount
		Dim strLayout
		Dim strMaxiumNoOfItems
		'Get Test data from excel sheet
		strLayout = objDataDict.Item("Layout" & iRowCountRef)
		strMaxiumNoOfItems = objDataDict.Item("MaxiumNoOfItems" & iRowCountRef)
		Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame"),MIN_WAIT)
		Set ObjPgGLAccount=Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame")
		'Enter Layout and Maximum Number of items
	 	If fEnterText(ObjPgGLAccount.WebEdit("txtLayout"),strLayout,"Layout") Then 
	 		If fEnterText(ObjPgGLAccount.WebEdit("txtMaximumNumberOfItems"),strMaxiumNoOfItems,"Maximum Number of Items") Then
	 			Call fRptWriteReport("PASSWITHSCREENSHOT","Fill List Output details","List output details are filled")
	 			fFioriFillListOutput = True
	 		Else
	 			Call fRptWriteReport("Fail","Fill Max Number Of Items ",strMaxiumNoOfItems & "- Max Number Of Items not Entered Properly")
	 			fFioriFillListOutput = False
	 			Call fRptWriteResultsSummary() 
	 			Exit Function
	 		End If
	 	Else
	 		Call fRptWriteReport("Fail","Fill LayOut Name ",strLayout & "- Layout Name not Entered Properly")
	 		fFioriFillListOutput = False
	 		Call fRptWriteResultsSummary() 
	 		Exit Function
	 	End  IF 
		'Clear Object
		Set ObjPgGLAccount = Nothing
	On error goto 0			
End Function	
'***************************************************************************************************************************************
''	Function Name					:				fFioriPerformExecute
''	Objective						:				Click on Execute button and Verify information message
''	Input Parameters				:				Nil
''	Output Parameters			    :				Nil
''	Date Created					:				28/April/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fFioriPerformExecute(objDataDict,iRowCountRef)
		On error resume next
		'Verify if Step Failed, If yes, it will not run the function
	    If Environment("StepFailed") = "YES" Then
			Exit Function
		End If	
		'Declarations
		Dim ObjPgGLAccount
		Dim strMsg
		fFioriPerformExecute = False
		'Get test data from excel
		strMsg = objDataDict.Item("Message" & iRowCountRef)
		Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame"),MID_WAIT)
		Set ObjPgGLAccount=Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame")
		'Click Execute button
		If fClick(ObjPgGLAccount.WebButton("btnExecute"),"Execute") Then
			strInfoMsg = Lcase(fGetText (ObjPgGLAccount.WebElement("wePopUpMsg"),"innertext","Information Message"))
				If instr(1,strInfoMsg,Lcase(strMsg)) Then
					Call fRptWriteReport("PASSWITHSCREENSHOT" ,"Verify message details in Popup window",strInfoMsg &" - message appeared in Popup window")	
						 If fClick(ObjPgGLAccount.WebButton("btnContinue"),"Continue") Then
							Wait (MIN_WAIT)
							fFioriPerformExecute = TRUE
						 Else
							fFioriPerformExecute = False
							Call fRptWriteResultsSummary() 
							Exit Function
						 End If 
				Else
					Call fRptWriteReport("Fail" ,"Verify message details in Popup window","Expected message details are not appeared in Popup window")	
					fFioriPerformExecute = False
					Call fRptWriteResultsSummary() 
					Exit Function					
				End If
		Else
			Call fRptWriteReport("Fail","Click on Execute","Execute Button not clicked or Not Exist")
			fFioriPerformExecute = False
			Call fRptWriteResultsSummary() 
			Exit Function
		End If 
		'Clear Object
		Set ObjPgGLAccount = Nothing
	On error goto 0		
End Function
'***************************************************************************************************************************************
''	Function Name					:				fFioriVerifyDetailsInGLAccountLineItem
''	Objective						:				Verify Details in G/L Account Line Item display page
''	Input Parameters				:				Nil
''	Output Parameters			    :				Nil
''	Date Created					:				28/April/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		29/April/2020 - Added 'On error resume next and Reporting
'**************************************************************************************************************************************
Public Function fFioriVerifyDetailsInGLAccountLineItem()
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim ObjPgGLAccount
	Dim intGLAccountNumber
	Dim intCompanyCodeID
		Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame"),MID_WAIT)
		'Create Object
		Set ObjPgGLAccount=Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame")
		'Verify G/L Account label	
		If fVerifyObjectExist(ObjPgGLAccount.WebElement("weGLAccount")) Then
			'Verify G/L Account Number
			If fVerifyObjectExist(ObjPgGLAccount.WebElement("weGLAccountNo")) then
					'Get G/LAccount Number
					intGLAccountNumber=fGetRoProperty(ObjPgGLAccount.WebElement("weGLAccountNo"),"innertext","G/L Account Number")
					Environment.Value("StepName") = "Verify G/L Account Label and Account Number in GL Account page"	
					Call fRptWriteReport("Pass",Environment.Value("StepName"), "G/L Account label displayed and G/L Account Number is " & intGLAccountNumber&"  in GL Account page")
						'Verify CompanyCode label
						If fVerifyObjectExist(ObjPgGLAccount.WebElement("weCompanyCode")) then
								'Verify CompanyCode ID
								If fVerifyObjectExist(ObjPgGLAccount.WebElement("weCompanyCodeID")) Then
									'Get G/LAccount Number
									intCompanyCodeID=fGetRoProperty(ObjPgGLAccount.WebElement("weCompanyCodeID"),"innertext","Company Code ID")
									Environment.Value("StepName") = "Verify Company Code Label and Company Code ID in GL Account page"	
									Call fRptWriteReport("Pass",Environment.Value("StepName"), "Company Code label displayed and Company Code is " & intCompanyCodeID&" in GL Account page")
								Else
									Call fRptWriteReport("Fail","Verify and Get Company Code ID in GL Account page", "Company Code ID not displayed  in GL Account page")
									Call fRptWriteResultsSummary() 
									Exit Function	
								End If
						Else
							Call fRptWriteReport("Fail","Verify Company Code Label in GL Account page", "Company Code label not displayed in GL Account page")	
							Call fRptWriteResultsSummary() 
							Exit Function
						End If
			Else
				Call fRptWriteReport("Fail","Verify and Get G/L Account Number in GL Account page", "G/L Account Number not displayed in GL Account page")	
				Call fRptWriteResultsSummary() 
				Exit Function
			End  IF
		Else
			Call fRptWriteReport("Fail","Verify G/L Account label in GL Account page", "G/L Account label not displayed in GL Account page")	
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
		'Clear object
		Set ObjPgGLAccount	= Nothing
	On error goto 0			
End Function

'***************************************************************************************************************************************
''	Function Name					:				fFioriGLAccountVerifyColumnName
''	Objective						:				Verify Column names in G/L Account Line item display page
''	Input Parameters				:				Nil
''	Output Parameters			    :				Nil
''	Date Created					:				28/April/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		29/April/2020 - Added 'On error resume next and Reporting
'**************************************************************************************************************************************
Public Function fFioriGLAccountVerifyColumnName(objDataDict,iRowCountRef)	
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    Dim arrColData
    Dim strInnerText
    Dim intCount
	Dim ObjPgGLAccount
	Dim intList
	Dim intRowCount
	Dim intColNum
	Dim stColumnName
	Dim strColumnDetails
		'Get Column details for verification 
		strColumnDetails = objDataDict.Item("ColumnDetails" & iRowCountRef)
		'Set Object
		Set ObjPgGLAccount=Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmGLFrame")
		arrIndividualColData = Split(strColumnDetails,";")
		' Continue loop
			For intList = 0 To Ubound(arrIndividualColData)
				arrColData = Split(arrIndividualColData(intList),",")
				intRowCount = arrColData(0)
				intColNum = arrColData(1)
				stColumnName = arrColData(2)
				' Read table data
					For intItearion = 6 To intRowCount
						If ObjPgGLAccount.WebElement("html id:=M0:46:::"&intItearion&":"&intColNum&"_l").Exist(1) then
						    	strInnerText = ObjPgGLAccount.WebElement("html id:=M0:46:::"&intItearion&":"&intColNum&"_l").GetROProperty("innertext")
						        If intItearion = 6 Then
						            intCount = 0
							            If strcomp(lcase(strInnerText), Lcase(stColumnName),1)=0 Then
							            	fFioriGLAccountVerifyColumnName = True
							            	Exit For
							            End If
						        Else
						           	fFioriGLAccountVerifyColumnName = False
						        End If
						 End if     
					Next		
				    If 	fFioriGLAccountVerifyColumnName = False	 Then
				    	Call fRptWriteReport("Fail","Verify GL Account '"& stColumnName +"' Column in GL Account page", "Column "&stColumnName&" not displayed in GL Account page")
				    	Call fRptWriteResultsSummary() 
						Exit Function	
				    Else
						Call fRptWriteReport("Pass","Verify GL Account '"& stColumnName +"' Column in GL Account page", "Column "&stColumnName&" displayed in GL Account page")	    
				    End If
		    Next
	    'clear object
	    Set ObjPgGLAccount = Nothing
	On error goto 0	    	   
End Function
'***************************************************************************************************************************************************************************************
''    Function Name                    :                fReadAndWriteGLAccountLineItemDataIntoExcel
''    Objective                        :                Read and Write G/L Account Line Item data in Excel
''    Input Parameters                 :                NIL
''    Output Parameters                :                NIL
''    Date Created                     :                02/May/2020
''    UFT/QTP Version                  :                15.0
''    Pre-requisites                   :                NIL  
''    Created By                       :                Cigniti
''    Modification Date                :                   
'*************************************************************************************************************************************************************************************** 
Public Function fReadAndWriteGLAccountLineItemDataIntoExcel(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim GLAccountTableColumnData_1
	Dim GLColumnName_1
	Dim GLAccountTableColumnData_2
	Dim GLColumnName_2
	Dim GLAccountTableColumnData_3
	Dim GLColumnName_3
	Dim GLAccountTableColumnData_4
	Dim GLColumnName_4
	Dim GLAccountTableColumnData_5
	Dim GLColumnName_5
	Dim GLAccountTableColumnData_6
	Dim GLColumnName_6
	Dim GLAccountTableColumnData_7
	Dim GLColumnName_7
	Dim GLAccountTableColumnData_8
	Dim GLColumnName_8
	Dim GLAccountTableColumnData_9
	Dim GLColumnName_9
	Dim GLAccountTableColumnData_10
	Dim GLColumnName_10
	Dim GLAccountTableColumnData_11
	Dim GLColumnName_11
	Dim GLAccountTableColumnData_12
	Dim GLColumnName_12
	Dim GLAccountTableColumnData_13
	Dim GLColumnName_13
	Dim XLStorePath
	Dim strSheetName

	GLAccountTableColumnData = objDataDict.Item("GLAccountTableColumnData" & iRowCountRef)
	GLColumnName = objDataDict.Item("GLColumnName" & iRowCountRef)
	GLAccountTableColumnData_1 = objDataDict.Item("GLAccountTableColumnData_1" & iRowCountRef)
	GLColumnName_1 = objDataDict.Item("GLColumnName_1" & iRowCountRef)
	GLAccountTableColumnData_2 = objDataDict.Item("GLAccountTableColumnData_2" & iRowCountRef)
	GLColumnName_2 = objDataDict.Item("GLColumnName_2" & iRowCountRef)
	GLAccountTableColumnData_3 = objDataDict.Item("GLAccountTableColumnData_3" & iRowCountRef)
	GLColumnName_3 = objDataDict.Item("GLColumnName_3" & iRowCountRef)
	GLAccountTableColumnData_4 = objDataDict.Item("GLAccountTableColumnData_4" & iRowCountRef)
	GLColumnName_4 = objDataDict.Item("GLColumnName_4" & iRowCountRef)
	GLAccountTableColumnData_5 = objDataDict.Item("GLAccountTableColumnData_5" & iRowCountRef)
	GLColumnName_5 = objDataDict.Item("GLColumnName_5" & iRowCountRef)
	GLAccountTableColumnData_6 = objDataDict.Item("GLAccountTableColumnData_6" & iRowCountRef)
	GLColumnName_6 = objDataDict.Item("GLColumnName_6" & iRowCountRef)
	GLAccountTableColumnData_7 = objDataDict.Item("GLAccountTableColumnData_7" & iRowCountRef)
	GLColumnName_7 = objDataDict.Item("GLColumnName_7" & iRowCountRef)
	GLAccountTableColumnData_8 = objDataDict.Item("GLAccountTableColumnData_8" & iRowCountRef)
	GLColumnName_8 = objDataDict.Item("GLColumnName_8" & iRowCountRef)
	GLAccountTableColumnData_9 = objDataDict.Item("GLAccountTableColumnData_9" & iRowCountRef)
	GLColumnName_9 = objDataDict.Item("GLColumnName_9" & iRowCountRef)
	GLAccountTableColumnData_10 = objDataDict.Item("GLAccountTableColumnData_10" & iRowCountRef)
	GLColumnName_10 = objDataDict.Item("GLColumnName_10" & iRowCountRef)
	GLAccountTableColumnData_11 = objDataDict.Item("GLAccountTableColumnData_11" & iRowCountRef)
	GLColumnName_11 = objDataDict.Item("GLColumnName_11" & iRowCountRef)
	GLAccountTableColumnData_12 = objDataDict.Item("GLAccountTableColumnData_12" & iRowCountRef)
	GLColumnName_12 = objDataDict.Item("GLColumnName_12" & iRowCountRef)
	GLAccountTableColumnData_13 = objDataDict.Item("GLAccountTableColumnData_13" & iRowCountRef)
	GLColumnName_13 = objDataDict.Item("GLColumnName_13" & iRowCountRef)
	XLStorePath =  objDataDict.Item("XLStorePath" & iRowCountRef)
	strSheetName = objDataDict.Item("SheetName" & iRowCountRef)

	' Read and store Status Column data
	arrStatus = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName,arrStatus)
	' Read and store Assignment Column data
	arrAssignment = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_1)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_1,arrAssignment)
	' Read and store Pstng Date Column data
	arrPstngDate = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_2)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_2,arrPstngDate)
	' Read and store DocumentNo Column data
	arrDocumentNo = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_3)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_3,arrDocumentNo)
	' Read and store DocumentNo Column data
	arrBusA = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_4)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_4,arrBusA)
	' Read and store Type Column data
	 arrType = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_5)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_5,arrType)
	' Read and store Doc. Date Column data
	arrDocDate = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_6)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_6,arrDocDate)
	' Read and store PK Column data
	arrPK = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_7)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_7,arrPK)
	' Read and store DC Amount Column data
	arrDCAmount = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_8)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_8,arrDCAmount)
	' Read and store Curr Column data
	arrCurr = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_9)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_9,arrCurr)
	' Read and store LC Amount" Column data
	arrLCAmount = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_10)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_10,arrLCAmount)
	' Read and store LCurr Column data
	arrLCurr = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_11)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_11,arrLCurr)
	' Read and store Group Currency Column data
	arrGroupCurrency = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_12)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_12,arrGroupCurrency)
	' Read and store CurrGroup Column data
	arrCurrGroup = fFioriGetGLAccountTableColumnaData(GLAccountTableColumnData_13)	
	Call fFioriAddDataIntoDataTableBasedOnColumn(GLColumnName_13,arrCurrGroup)
	'Export data in excel sheet
	Call fDataTableExportSheet(XLStorePath,strSheetName)
	On error goto 0		
End  Function

''*******************************************************************************************************
''	Function Name					:				    fnGetTableHeaderColumnNumber
''	Objective						:					Get Header column number based on column name
''	Input Parameters				:					sObjectName, intHeaderColumnNumber, intHeaderRowNumber, strColumnNameEN
''	Output Parameters			    :					Column number 
''	Date Created					:					05/May/2020
''	QTP Version						:					15
''	Pre-requisites					:					NIL  
''	Created By						:					Cigniti
''	Modification Date		        :		   			
'********************************************************************************************************  
Public Function fGetTableHeaderColumnNumber (Byval sObjectName,Byval intHeaderColumnNumber,Byval intHeaderRowNumber,Byval strColumnNameEN)
		On error resume next
		'Verify if Step Failed, If yes, it will not run the function
	    If Environment("StepFailed") = "YES" Then
			Exit Function
		End If	
			'Variable Declaration / Initialization
		Dim found
		Dim intColCount
		Dim intCount
		Dim strColumnName		
		blnfound = False
		If strColumnNameEN <> "" Then		
			intColCount= sObjectName.ColumnCount(intHeaderColumnNumber)
				For intCount=1 to intColCount
					strColumnNameAN=sObjectName.GetCellData(intHeaderRowNumber,intCount)
						If  Trim(Ucase(strColumnNameAN)) = Trim(Ucase(strColumnNameEN)) Then
							blnfound = True	
							Exit for 							
						End If
				Next
		End If
		'Report			
		If  blnfound then 
			Call fRptWriteReport("Pass", "Search "&ColumnName& " column name" ,ColumnName & " Column name  was found")
		Else
			Call fRptWriteReport("Fail", "Search "&ColumnName& " column name" , ColumnName & " Column name  was not found")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
		fGetTableHeaderColumnNumber = intCount
	On error goto 0		
  End Function

'********************************************************************************************************************************************************************************************
''    Function Name                        :                fDisplaySupplierBalances
''    Objective                            :                Display Supplier Balances
''    Input Parameters                     :                strSearchText,strSupplierSearch
''    Output Parameters                    :                Nil
''    Date Created                         :                05/May/2020
''    UFT/QTP Version                      :                15.0
''    Pre-requisites                       :                NIL  
''    Created By                           :               Cigniti
''    Modification Date                    :                 
''********************************************************************************************************************************************************************************************
Public Function fDisplaySupplierBalances(strSearchText,strSupplierSearch)
'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
Dim objBrwAndPg,strSupplierValue,strSupplier
	Call fFioriFetchAppFromHomePage(strSearchText)
	Call fSynUntilObjExists(Browser("brFiori").Page("pgFiori"),MIN_WAIT)	
	Set objBrwAndPg = Browser("brFiori").Page("pgFiori")
		If fVerifyObjectExist(objBrwAndPg.WebElement("weAppAccountsPayableOverview")) Then
			Call fClick(objBrwAndPg.WebElement("weAppAccountsPayableOverview"),"App Accounts Payable Overview")
		End  If
	Call fSynUntilObjExists(objBrwAndPg.WebElement("weAccountsPayableOverview"),MID_WAIT)
		If fVerifyObjectExist(objBrwAndPg.WebElement("weAccountsPayableOverview")) Then
			Call fClick(objBrwAndPg.WebElement("weAccountsPayableOverview"),"Accounts Payable Overview")
		End  If
	Call fSynUntilObjExists(objBrwAndPg.WebElement("weManagePaymentBlocks"),MID_WAIT)
		If fVerifyObjectExist(objBrwAndPg.WebElement("weManagePaymentBlocks")) Then
			Call fClick(objBrwAndPg.WebElement("weManagePaymentBlocks"),"Manage Payment Blocks")
		End If
	Call fSynUntilObjExists(objBrwAndPg.WebEdit("txtSearch"),MID_WAIT)
		If fVerifyObjectExist(objBrwAndPg.WebEdit("txtSearch")) Then
			Call fEnterText(objHomePage.WebEdit("txtSearch"),strSupplierSearch,"Supplier Search field")
			Call fClick(objBrwAndPg.WebElement("weSearchIcon"),"Search Icon")
				If fVerifyObjectExist(objBrwAndPg.WebElement("weAutodeskAsiaPteLtd")) Then
					Call fClick(objBrwAndPg.WebElement("weAutodeskAsiaPteLtd"),"Autodesk Asia Pte Ltd Supplier Accounts")
					Call fClick(objBrwAndPg.WebElement("weG/LAccountDocument"),"G/L Account Document")
				End  If 
				If fVerifyObjectExist(objBrwAndPg.WebButton("btnAutodeskAsiaPteLtd")) Then
					strSupplier = fGetRoProperty(objBrwAndPg.WebButton("btnAutodeskAsiaPteLtd"),"innertext","Supplier Value")
					Call fClick(objBrwAndPg.WebButton("btnAutodeskAsiaPteLtd"),"Autodesk Asia Pte Ltd")
						If fVerifyObjectExist(objBrwAndPg.Link("lnkDisplaySupplierBalances")) Then
							Call fClick(objBrwAndPg.Link("lnkDisplaySupplierBalances"),"Display Supplier Balances")
						End If
				End If
			Call fSynUntilObjExists(objBrwAndPg.WebEdit("txtSupplier"),MID_WAIT)
				If fVerifyObjectExist(objBrwAndPg.WebEdit("txtSupplier")) Then
					strSupplierValue =  fSplitFor(2,"(",strSupplier)
					blnCompareSupplierValue = fVerifyProperty(objBrwAndPg.WebElement("weSupplierToken"),"innertext","="&Replace(strSupplierValue,")",""))
				End  If
				If fVerifyObjectExist(objBrwAndPg.WebEdit("txtSupplier"))  AND blnCompareSupplierValue = "True" Then
					Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Display Supplier Balances page","Display Supplier Balances page is displayed successfully")
				Else
					Call fRptWriteReport("Fail", "Verify Display Supplier Balances page","Display Supplier Balances page is not displayed")
					Exit Function
				End If
		End  If
		Set objBrwAndPg = Nothing	
End Function
'********************************************************************************************************************************************************************************************
''    Function Name                       :                fFioriCustomerLineItemDisplay
''    Objective                           :                Customer Line Item Display
''    Input Parameters                    :                strCustomerNumber,intCompanyCode,strLineItemSelectionStatus,strDropDownSelected,strSubDropDownSelected
''    Output Parameters                   :                Nil
''    Date Created                        :                01/May/2020
''    UFT/QTP Version                     :                15.0
''    Pre-requisites                      :                NIL  
''    Created By                          :               Cigniti
''    Modification Date                   :                 
''********************************************************************************************************************************************************************************************
Public Function fFioriCustomerLineItemDisplay(strCustomerNumber,intCompanyCode,strLineItemSelectionStatus,strDropDownSelected,strSubDropDownSelected)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objBrwAndPg
	
	Call fSynUntilObjExists(Browser("brFiori").Page("pgFiori"),MID_WAIT)
	Set objBrwAndPg = Browser("brFiori").Page("pgFiori")
	Call fSynUntilObjExists(objBrwAndPg.WebElement("weManageCus­tomerLine"),MID_WAIT)
		If fVerifyObjectExist(objBrwAndPg.WebElement("weManageCus­tomerLine")) Then
			Call fClick(objBrwAndPg.WebElement("weManageCus­tomerLine"),"Man­age Cus­tomer Line Items FBL5N")
		End If
	Set objBrwAndPg = Browser("brFioriAutodesk").Page("pgFioriAutodesk").SAPFrame("sfCustomerLineItemDisplay")
	Call fSynUntilObjExists(objBrwAndPg.SAPEdit("txtCompanycode"),MID_WAIT)
		If fVerifyObjectExist(objBrwAndPg.SAPEdit("txtCompanycode")) Then
			Call fEnterText(objBrwAndPg.SAPEdit("txtCustomeraccount"),strCustomerNumber,"Customer Number")
			Call fEnterText(objBrwAndPg.SAPEdit("txtCompanycode"),intCompanyCode,"Company Code")
			Call fSelect(objBrwAndPg.SAPRadioGroup("rbSAPRadioGroup"),strLineItemSelectionStatus,"Line Item Selection")
			Call fClick(objBrwAndPg.SAPButton("btnExecute"),"Execute")
		End  If
	Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnDisplayDocument"),MID_WAIT)
		If fVerifyObjectExist(objBrwAndPg.SAPButton("btnDisplayDocument")) Then
			Call fClick(objBrwAndPg.WebElement("weOpenItemsSymbol"),"Open Item Symbol")
			Call fClick(objBrwAndPg.SAPButton("btnDisplayDocument"),"Display Document")
		End  If	
	Set objBrwAndPg = Browser("brFioriAutodesk").Page("pgFioriAutodesk").SAPFrame("sfDisplayDocumentLine")
	Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnMore"),MID_WAIT)
		If fVerifyObjectExist(objBrwAndPg.SAPButton("btnMore")) Then
			Call fClick(objBrwAndPg.SAPButton("btnMore"),"More")
			Call fSelect(objBrwAndPg.SAPDropDownMenu("lstDisplayDocumentHeader"),strDropDownSelected,"Goto ")
			Call fSelect(objBrwAndPg.SAPDropDownMenu("lstDisplayDocumentHeader"),strSubDropDownSelected,"Document Overview (F9)")
		End  If
	Set objBrwAndPg = Browser("brFioriAutodesk").Page("pgFioriAutodesk")
	Call fSynUntilObjExists(objBrwAndPg.WebElement("weDisplayDocumentData"),MID_WAIT)
		If fVerifyObjectExist(objBrwAndPg.WebElement("weDisplayDocumentData")) Then
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Display Document Data page","Display Document Data page is displayed Successfully")
	   	Else
		  	Call fRptWriteReport("Fail", "Verify Display Document Data page","Display Document Data page is not displayed")
			Exit Function
		End  If 
	Set objBrwAndPg = Nothing
	On error goto 0
End Function
'***************************************************************************************************************************************************************************************
''    Function Name                    :                fFioriAddDataIntoDataTableBasedOnColumn
''    Objective                        :                Write O/P data in to data table
''    Input Parameters                 :                strExcelPath,strExcelSheetName
''    Output Parameters                :                NIL
''    Date Created                     :                02/May/2020
''    UFT/QTP Version                  :                15.0
''    Pre-requisites                   :                NIL  
''    Created By                       :                Cigniti
''    Modification Date                :                   
'*************************************************************************************************************************************************************************************** 
Public Function fFioriAddDataIntoDataTableBasedOnColumn(strColumnName,strColumnData)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim intVal
	Dim strVal
	
	'arrColumnData=split(strColumnData,";")
	fFioriAddDataIntoDataTableBasedOnColumn = False
	DataTable.AddSheet("Sheet1")
	DataTable.GetSheet("Sheet1").AddParameter strColumnName,""
		For intVal = 0 To Ubound(strColumnData)
			strVal=strColumnData(intVal)
			DataTable.SetCurrentRow(intVal+1)	
			DataTable.GetSheet("Sheet1").GetParameter(strColumnName).Value=strVal
			DataTable.SetNextRow
			fFioriAddDataIntoDataTableBasedOnColumn = True
		Next
		If  fFioriAddDataIntoDataTableBasedOnColumn then 
			Call fRptWriteReport("Pass", "Import -" &strColumnName& " data" ,"Import -" &strColumnName& " data from application to Data Table")
		Else
			Call fRptWriteReport("Fail", "Import -" &strColumnName& " data" ,strColumnName& " column data not read from application or not written in Data Table")
			fFioriAddDataIntoDataTableBasedOnColumn = False
			Exit Function
		End If
   On error goto 0
End  Function
'***************************************************************************************************************************************************************************************
''    Function Name                    :                fDataTableExportSheet
''    Objective                        :                Export Data table data in to Excel sheet
''    Input Parameters                 :                strExcelPath,strExcelSheetName
''    Output Parameters                :                NIL
''    Date Created                     :                02/May/2020
''    UFT/QTP Version                  :                15.0
''    Pre-requisites                   :                NIL  
''    Created By                       :                Cigniti
''    Modification Date                :                   
'*************************************************************************************************************************************************************************************** 
Public Function fDataTableExportSheet(strExcelPath,strExcelSheetName)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Exports the current sheet to the specifed path.
	dtDateAndTime= Replace(Replace(Replace(now(),":",""),"/","")," ","")
		Call fCreateExcelFile(strExcelPath+"\"+strExcelSheetName,dtDateAndTime)
	DataTable.ExportSheet strExcelPath+"\"+strExcelSheetName+dtDateAndTime+".xlsx","Sheet1"
	On error goto 0	
End  Function
'***************************************************************************************************************************************************************************************
''    Function Name                    :                fFioriGLAccountTableLayoutChange
''    Objective                        :                Used to change Fiori GL Account Table Layout
''    Input Parameters                 :                srtColumnName,intColNum,intColLength
''    Output Parameters                :                NIL
''    Date Created                     :                30/04/2020
''    UFT/QTP Version                  :                15.0
''    Pre-requisites                   :                NIL  
''    Created By                       :                Cigniti
''    Modification Date                :                   
'*************************************************************************************************************************************************************************************** 
Public Function fFioriGLAccountTableLayoutChange(srtColumnName,intColNum,intColLength)
    On error resume next
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    'Variable declaration
    Dim objGLFrame
    Dim intTableRowCount
    Dim strCellData
    Dim blnFound
    Dim strCellData1  
    Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk"),MID_WAIT)
    'Set Frame 
    Set objGLFrame = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
        If fVerifyObjectExist(objGLFrame.WebTable("tblChangeLayout")) Then
            blnFound = False
            'Get table row count
            intTableRowCount = fGetRoProperty(objGLFrame.WebTable("tblChangeLayout"),"rows","Change Layout")
                For intIteration = 1 To intTableRowCount
                    'Get table cell data
                    strCellData = fGetCelldata(objGLFrame.WebTable("tblChangeLayout"),intIteration,1,"Change Layout")
                    	' Enter Position details
                        If srtColumnName = strCellData Then
                                'Click on table cell
                                If objGLFrame.WebTable("tblChangeLayout").ChildItem(intIteration,2,"WebList",0).exist(2) Then
                                    objGLFrame.WebTable("tblChangeLayout").ChildItem(intIteration,2,"WebList",0).click
                                End If
                            'Enter data in table cell
                            Call fEnterTextInCell(objGLFrame.WebTable("tblChangeLayout"),intIteration,2,"SAPEdit",0,intColNum)                           
                            objGLFrame.SAPEdit("html id:=tbl.*"&intIteration&",2.*c").set intColNum
                         End If
                        'Enter Length details
                        If srtColumnName = strCellData Then
                                'Click on table cell
                                If objGLFrame.WebTable("tblChangeLayout").ChildItem(intIteration,3,"WebList",0).exist(2) Then
                                    objGLFrame.WebTable("tblChangeLayout").ChildItem(intIteration,3,"WebList",0).click
                                End If
                            'Enter data in table cell
                            Call fEnterTextInCell(objGLFrame.WebTable("tblChangeLayout"),intIteration,3,"SAPEdit",0,intColLength)                            
                            objGLFrame.SAPEdit("html id:=tbl.*"&intIteration&",3.*c").set intColLength
                            blnFound = True
                            Exit For
                        End If
                Next
            If blnFound Then
                Call fRptWriteReport("Pass", "Verify GL Account Line Item Column in GL Account page","GL Account Line Item Column :"&srtColumnName&" displayed successfully in GL Account page")
            Else
                Call fRptWriteReport("Fail", "Verify GL Account Line Item Column in GL Account page","GL Account Line Item Column :"&srtColumnName&" not displayed in GL Account page")
                Call fRptWriteResultsSummary() 
                Exit Function
            End If
        Else            
            Call fRptWriteReport("Fail", "Verify GL Account Line Item Column in GL Account page","GL Account Line Item table not displayed")
            Call fRptWriteResultsSummary() 
            Exit Function
        End If  
    Set objGLFrame = Nothing
   On error goto 0	 
End Function


'************************************************************************************************************************************************
'Function Name                                :                    fGetRowNumberInTableBasedonColumnData
'Objective                                    :                     Get Row Number Based on expected data
'Input Parameters                            :                    sObjectName,intColumnNo,strColExpValue
'Output Parameters                            :                    Row Number
'Date Created                                :                    05/May/2020
'QTP Version                                :                    14.50
'Pre-requisites                                :                    NIL  
'Created By                                    :                    Cigniti
'Modification Date                            :           
'************************************************************************************************************************************************
Public Function fGetRowNumberInTableBasedonColumnData (Byval sObjectName,Byval intColumnNo,ByVal strColExpValue)                
    On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	    
    Dim blnfound
    Dim intRowcount
    Dim strCurrVal
    Dim intRowNo    
    blnfound=False        
    If intColumnNo >= 0  then
        intRowcount = sObjectName.Rowcount
        On error resume next
        For intRowNo = 1 to intRowcount
            strCurrVal=sObjectName.getcelldata(intRowNo,intColumnNo)
                If Lcase(trim(strColExpValue)) =Lcase(trim(strCurrVal))  Then
                    sObjectName.ChildItem(intRowNo,intColumnNo,"WebElement",1).Highlight                    
                    blnfound=True
                    fGetRowNumberInTableBasedonColumnData=intRowNo                        
                    Exit for 
                End If
        Next
    End  If
        'Report
        If  blnfound then 
            Call fRptWriteReport("Pass", "Fetch Row Number from webtable",strColExpValue & " related Row Number should be fetched based on Expected data","Fetch the Row Number Based on '"&strColExpValue & "' data")
        Else
            Call fRptWriteReport("Fail", "Fetch Row Number from webtable",strColExpValue & " related Row Number should be fetched based on Expected data","Unable to fetch the Row Number Based on '"&strColExpValue & "' data")
            Call fRptWriteResultsSummary() 
            Exit Function
        End If
    On error goto 0	
End function


'***************************************************************************************************************************************
''	Function Name					:				fSwitchWorkView
''	Objective						:				Select Switch work view Option
''	Input Parameters				:				SwitchWorkViewOption (ex: Personal View/ All Users View/...)
''	Output Parameters			    :				Nil
''	Date Created					:				12/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fSwitchWorkView(objDataDict,iRowCountRef)

		'Verify if Step Failed, If yes, it will not run the function
	    If Environment("StepFailed") = "YES" Then
			Exit Function
		End If	
		Dim objGLFrame
		Dim strSwitchWorkViewOption
		fSwitchWorkView = False
		
		'Get data from excel sheet
		strSwitchWorkViewOption = objDataDict.Item("SwitchWorkViewOption" & iRowCountRef)
		Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmVIM"),MID_WAIT)
		'Set Frame 
    	Set objGLFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmVIM")
		If fClick( objGLFrame.WebButton("btnSwitchWorkView"),"Switch Work View") Then
			'Wait till Switch Work View Options exist
			Call fSynUntilObjExists(objGLFrame.SAPRadioGroup("rgSwitchWorkViewOptions"),MAX_WAIT)
				'Select Switch Work View option based on user input
				Select Case Lcase(strSwitchWorkViewOption)
						Case Lcase("Personal View")
							'Select Personal View" option
							If fSelect(objGLFrame.SAPRadioGroup("rgSwitchWorkViewOptions"),"Personal View","Personal View") Then
								Call fSynUntilObjExists(objGLFrame.WebButton("btnContinue"),MIN_WAIT)
									If fClick(objGLFrame.WebButton("btnContinue"),"Continue") Then
										fSwitchWorkView = True
										Call fRptWriteReport("PASSWITHSCREENSHOT","View only Personal View Items","Personal list displayed Sucessfully")
									Else
										fSwitchWorkView = False
										Call fRptWriteReport("Fail","Click on Continue ","Continue button not not clicked in popup window")
										Call fRptWriteResultsSummary() 
										Exit Function
									End If
							Else
								fSwitchWorkView = False
								Call fRptWriteReport("Fail","Select Personal View Option","Personal View Option not selected Sucessfully")
								Call fRptWriteResultsSummary() 
								Exit Function
							End  IF
						Case Lcase("Other User's View")
							If fSelect(objGLFrame.SAPRadioGroup("rgSwitchWorkViewOptions"),"Other User's View","Other User's View") Then
								Call fSynUntilObjExists(objGLFrame.WebButton("btnContinue"),MIN_WAIT)
									If fClick(objGLFrame.WebButton("btnContinue"),"Continue") Then
										fSwitchWorkView = True
										Call fRptWriteReport("PASSWITHSCREENSHOT","View Other User's Items","Other User's list displayed Sucessfully")
									Else
										fSwitchWorkView = False
										Call fRptWriteReport("Fail","Click on Continue ","Continue button not not clicked in popup window")
										Call fRptWriteResultsSummary() 
										Exit Function
									End If
							Else
								fSwitchWorkView = False
								Call fRptWriteReport("Fail","Select Other User's list Option","Other User's list Option not selected Sucessfully")
								Call fRptWriteResultsSummary() 
								Exit Function
							End  IF
							
						Case Lcase("All Users View")
								'Select All Users View option
								If fSelect(objGLFrame.SAPRadioGroup("rgSwitchWorkViewOptions"),"All Users View","All Users View") Then
									Call fSynUntilObjExists(objGLFrame.WebButton("btnContinue"),MIN_WAIT)
										If fClick(objGLFrame.WebButton("btnContinue"),"Continue") Then
											fSwitchWorkView = True
											Call fRptWriteReport("PASSWITHSCREENSHOT","View All Users list"," All Users list displayed Sucessfully")
										Else
											fSwitchWorkView = False
											Call fRptWriteReport("Fail","Click on Continue ","Continue button not not clicked in popup window")
											Call fRptWriteResultsSummary() 
											Exit Function
										End If
								Else
									fSwitchWorkView = False
									Call fRptWriteReport("Fail","Select All Users list Option","All Users list Option not selected Sucessfully")
									Call fRptWriteResultsSummary() 
									Exit Function
								End  IF
					 End Select	
			Else
			 	Call fRptWriteReport("Fail","Click on Switch Work View", "Switch Work View button not Clicked")
			 	fSwitchWorkView = False
			 	Call fRptWriteResultsSummary() 
			 	Exit Function
		 	End  If
		 	
			'CLick on Continue
			Call fClick( objGLFrame.WebButton("btnContinue"),"Continue")
			Set objGLFrame = Nothing
			 
	
End Function
'***************************************************************************************************************************************
''	Function Name					:				fFillAutomaticPaymentTransactionsStatus
''	Objective						:				Fill Automatic Payment Transactions: Status details
''	Input Parameters				:				intRunDate,strIdentification
''	Output Parameters			    :				Nil
''	Date Created					:				12/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		05/28/2020 - Reporting details are updated
'**************************************************************************************************************************************
Public Function fFillAutomaticPaymentTransactionsStatus(objDataDict,iRowCountRef)
		On error resume next
		'Verify if Step Failed, If yes, it will not run the function
	    If Environment("StepFailed") = "YES" Then
			Exit Function
		End If	
		'Declarations
		Dim objGLFrame
		Dim intRunDate
		Dim strIdentification
		
		fFillAutomaticPaymentTransactionsStatus = False
		'Read data from excel
		intRunDate = fGetSingleValue("RunDate","TestData",Environment("TestName"))
		'Get Identification from testdata sheet
		strIdentification = fGetSingleValue("Identification","TestData",Environment("TestName")) 
		Set objGLFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions")
		objGLFrame.Highlight
		Call fSynUntilObjExists(objGLFrame.SAPEdit("txtRunDate"),MID_WAIT)
		Call fSynUntilObjExists(objGLFrame.SAPEdit("txtIdentification"),MID_WAIT) 
		Call fSynUntilObjExists(objGLFrame.SAPEdit("txtIdentification"),MID_WAIT)
		
		'Enter Run date
		If Lcase(intRunDate)<> "n/a" and Lcase(intRunDate)<> "null" and Lcase(intRunDate)<> "no" Then
			intRunDate=fGetFutureDateAdd(intRunDate)
			If fEnterText(objGLFrame.SAPEdit("txtRunDate"),intRunDate,"Run Date") Then
				fFillAutomaticPaymentTransactionsStatus = True
			Else
				fFillAutomaticPaymentTransactionsStatus = False
				Call fRptWriteReport("Fail","Fill Automatic Payment Transaction status data","Run Date details are not filled")
				Call fRptWriteResultsSummary() 
				Exit Function
			End  IF
		End If
		'Enter Identification
		If Lcase(strIdentification)<> "n/a" and Lcase(strIdentification)<> "null" and Lcase(strIdentification)<> "no" Then
			If fEnterText(objGLFrame.SAPEdit("txtIdentification"),strIdentification,"Identification") Then
				fFillAutomaticPaymentTransactionsStatus = True
			Else
				fFillAutomaticPaymentTransactionsStatus = False
				Call fRptWriteReport("Fail","Fill Automatic Payment Transaction status data","Run Date details are not filled")
				Call fRptWriteResultsSummary() 
				Exit Function
			End  IF
		End If

		'Clear object
		Call fRptWriteReport("PASSWITHSCREENSHOT","Fill Automatic Payment Transaction status data","Run Date and Identification details are filled")
		Set objGLFrame = Nothing
		On error goto 0			
End Function	

'***************************************************************************************************************************************
''	Function Name					:				fFillAutomaticPaymentTransactionsParameters
''	Objective						:				Detailed filled in Automatic Payment Transactions Parameters page
''	Input Parameters				:				CompanyCode,PmtMeths,NextPstDate,VendorFrom,VendorTO,CustomerFrom,CustomerTO
''	Output Parameters			    :				Nil
''	Date Created					:				13/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fFillAutomaticPaymentTransactionsParameters(objDataDict,iRowCountRef)
		On error resume next
		'Verify if Step Failed, If yes, it will not run the function
	    If Environment("StepFailed") = "YES" Then
			Exit Function
		End If	
		'Declarations
		Dim objGLFrame
		Dim strEnterCompanyCode
		Dim strPmtMeths
		Dim strNextPstDate
		Dim intVendorFrom
		Dim intVendorTO
		Dim intCustomerFrom
		Dim intCustomerTO
		Dim intNextPstDate
		Dim intPstDate
		
		fFillAutomaticPaymentTransactionsParameters = False
		'Read data from excel
		strEnterCompanyCode = objDataDict.Item("CompanyCode" & iRowCountRef)
		strPmtMeths = objDataDict.Item("PmtMeths" & iRowCountRef)
		strtPstDate = objDataDict.Item("RunDate" & iRowCountRef)
		intPstDate = fGetFutureDateAdd(strNextPstDate)
		strNextPstDate = objDataDict.Item("NextPstDate" & iRowCountRef)
		intNextPstDate = fGetFutureDateAdd(strNextPstDate)
		intVendorFrom = objDataDict.Item("VendorFrom" & iRowCountRef)
		
		Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions"),MIN_WAIT)
		Set objGLFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions")
		'Click Parameter
		If fClick(objGLFrame.WebElement("weParameter"),"Parameter") Then
			Call fSynUntilObjExists(objGLFrame.WebTable("tblPaymentsControl"),MIN_WAIT)
			'Enter Posting Date
			intPstDate = fGetFutureDateAdd(strtPstDate)
				If fEnterText(objGLFrame.SAPEdit("txtPostingDate"),intPstDate,"Posting Date") Then
						If fClick(objGLFrame.WebTable("tblPaymentsControl").ChildItem(1,1,"WebList",0),"Company Code text field") Then
							'Enter CompanyCode
							If fEnterTextInCell(objGLFrame.WebTable("tblPaymentsControl"),1,1,"SAPEdit",0,strEnterCompanyCode) Then
								If  fClick(objGLFrame.WebTable("tblPaymentsControl").ChildItem(1,2,"WebList",0),"Pmt Meths text field") Then
									'Enter Pmt Meths
									If fEnterTextInCell(objGLFrame.WebTable("tblPaymentsControl"),1,2,"SAPEdit",0,strPmtMeths) Then
										'Enter Next PstDate
										If fClick(objGLFrame.WebTable("tblPaymentsControl").ChildItem(1,3,"WebList",0),"Next Pst Date") Then
											If fEnterTextInCell(objGLFrame.WebTable("tblPaymentsControl"),1,3,"SAPEdit",0,intNextPstDate) Then
												'Enter Vendor from
												If Lcase(intVendorFrom)<> "n/a" and Lcase(intVendorFrom)<> "null" and Lcase(intVendorFrom)<> "no" Then
													 If fEnterText(objGLFrame.SAPEdit("txtVendorFrom"),intVendorFrom,"Vendors From") Then
													 	Call fRptWriteReport("PASSWITHSCREENSHOT","Fill Automatic Payment Transactions Parameters","Automatic Payment Transaction details are filled properly")
													 	fFillAutomaticPaymentTransactionsParameters = TRUE
													 Else
														fFillAutomaticPaymentTransactionsParameters = False
														Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Parameters","Automatic Payment Transaction details are not filled properly")
														Call fRptWriteResultsSummary() 
														Exit Function
													End If
												Else
													fFillAutomaticPaymentTransactionsParameters = False
													Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Parameters","Automatic Payment Transaction details are not filled properly")
													Call fRptWriteResultsSummary() 
													Exit Function
												End If
											Else
												fFillAutomaticPaymentTransactionsParameters = False
												Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Parameters","Automatic Payment Transaction details are not filled properly")
												Call fRptWriteResultsSummary() 
												Exit Function
											End If
										Else
											fFillAutomaticPaymentTransactionsParameters = False
											Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Parameters","Automatic Payment Transaction details are not filled properly")
											Call fRptWriteResultsSummary() 
											Exit Function
										End If
									Else
										fFillAutomaticPaymentTransactionsParameters = False
										Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Parameters","Automatic Payment Transaction details are not filled properly")
										Call fRptWriteResultsSummary() 
										Exit Function
									End If
								Else
									fFillAutomaticPaymentTransactionsParameters = False
									Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Parameters","Automatic Payment Transaction details are not filled properly")
									Call fRptWriteResultsSummary() 
									Exit Function
								End If
							Else
								fFillAutomaticPaymentTransactionsParameters = False
								Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Parameters","Automatic Payment Transaction details are not filled properly")
								Call fRptWriteResultsSummary() 
								Exit Function
							End If
						Else
							fFillAutomaticPaymentTransactionsParameters = False
							Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Parameters","Automatic Payment Transaction details are not filled properly")
							Call fRptWriteResultsSummary() 
							Exit Function
						End If
				Else
					fFillAutomaticPaymentTransactionsParameters = False
					Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Parameters","Automatic Payment Transaction details are not filled properly")
					Call fRptWriteResultsSummary() 
					Exit Function
				End  IF
		Else
			fFillAutomaticPaymentTransactionsParameters = False
			Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Parameters","Automatic Payment Transaction details are not filled properly")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If 
		
		'For Future purpouse ############################################################################################################################
'		intVendorTO = objDataDict.Item("VendorTO" & iRowCountRef)
'		intCustomerFrom = objDataDict.Item("CustomerFrom" & iRowCountRef)
'		intCustomerTO = objDataDict.Item("CustomerTO" & iRowCountRef)
'		'Enter Vendor To
'		If Lcase(intVendorTO)<> "n/a" and Lcase(intVendorTO)<> "null" and Lcase(intVendorTO)<> "no" Then
'			Call fEnterText(objGLFrame.SAPEdit("txtVendorTO"),intVendorFrom,"Vendors To")
'		End If
'		'Enter Customer from
'		If Lcase(intCustomerFrom)<> "n/a" and Lcase(intCustomerFrom)<> "null" and Lcase(intCustomerFrom)<> "no" Then
'			Call fEnterText(objGLFrame.SAPEdit("txtCustomerFrom"),intVendorFrom,"Customer From")
'		End If
'		'Enter Customer To
'		If Lcase(intCustomerTO)<> "n/a" and Lcase(intCustomerTO)<> "null" and Lcase(intCustomerTO)<> "no" Then
'			Call fEnterText(objGLFrame.SAPEdit("txtCustomerTO"),intCustomerTO,"Customer To")
'		End If	
		'################################################################################################################################################
		'Clear object
		Set objGLFrame = Nothing
	On error goto 0			

End Function	
'***************************************************************************************************************************************
''	Function Name					:				fFillAutomaticPaymentTransactionsAdditionalLog
''	Objective						:				Detailed filled in Automatic Payment Transactions Additional Log
''	Input Parameters				:				DueDateCheck,PaymentMethodAllCases,PmntMethodNotSuccessful,LineItemsPaymentDocuments,VendorFrom
''	Output Parameters			    :				Nil
''	Date Created					:				13/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fFillAutomaticPaymentTransactionsAdditionalLog(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim objGLFrame
	Dim blnDueDateCheck
	Dim blnPaymentMethodAllCases
	Dim blnPmntMethodNotSuccessful
	Dim blnLineItemsPaymentDocuments
	Dim intVendorsFrom
		
		'Read data from excel
		blnDueDateCheck = objDataDict.Item("DueDateCheck" & iRowCountRef)
		blnPaymentMethodAllCases = objDataDict.Item("PaymentMethodAllCases" & iRowCountRef)
		blnPmntMethodNotSuccessful = objDataDict.Item("PmntMethodNotSuccessful" & iRowCountRef)
		blnLineItemsPaymentDocuments = objDataDict.Item("LineItemsPaymentDocuments" & iRowCountRef)
		intVendorsFrom = objDataDict.Item("VendorFrom" & iRowCountRef)
		fFillAutomaticPaymentTransactionsAdditionalLog = False
		Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions"),MIN_WAIT)
		Set objGLFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions")
		'Fill details
		If fClick(objGLFrame.WebElement("weAdditionalLog"),"Additional Log") Then		
			Call fSynUntilObjExists(objGLFrame.SAPCheckBox("chkDueDateCheck"),MIN_WAIT)
				If fSelect(objGLFrame.SAPCheckBox("chkDueDateCheck"),blnDueDateCheck,"Due Date Check") Then
					If fSelect(objGLFrame.SAPCheckBox("chkPaymentMethodAllCases"),blnPaymentMethodAllCases,"Payment Method Selection in All Cases") Then
						If fSelect(objGLFrame.SAPCheckBox("chkPmntMethodNotSuccessful"),blnPmntMethodNotSuccessful,"Pmnt Method Selection If Not Successful") Then
							If fSelect(objGLFrame.SAPCheckBox("chkLineItemsPaymentDocuments"),blnLineItemsPaymentDocuments,"Line Items of the Payment Documents") Then
								If fSelect(objGLFrame.SAPCheckBox("chkLineItemsPaymentDocuments"),blnLineItemsPaymentDocuments,"Line Items of the Payment Documents") Then
									If fEnterText(objGLFrame.SAPEdit("txtVendorsFrom"),intVendorsFrom,"Vendor Number") Then
										fFillAutomaticPaymentTransactionsAdditionalLog = TRUE
										Call fRptWriteReport("PASSWITHSCREENSHOT","Fill Automatic Payment Transactions Additional Log data","Automatic Payment Transactions Additional Log details are filled Properly")
									Else
										fFillAutomaticPaymentTransactionsAdditionalLog = False
										Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Additional Log data","Automatic Payment Transactions Additional Log details are not filled Properly")
										Call fRptWriteResultsSummary() 
										Exit Function
									End  IF
								Else
									fFillAutomaticPaymentTransactionsAdditionalLog = False
									Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Additional Log data","Automatic Payment Transactions Additional Log details are not filled Properly")
									Call fRptWriteResultsSummary() 
									Exit Function
								End  IF
							Else
								fFillAutomaticPaymentTransactionsAdditionalLog = False
								Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Additional Log data","Automatic Payment Transactions Additional Log details are not filled Properly")
								Call fRptWriteResultsSummary() 
								Exit Function
							End  IF
						Else
							fFillAutomaticPaymentTransactionsAdditionalLog = False
							Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Additional Log data","Automatic Payment Transactions Additional Log details are not filled Properly")
							Call fRptWriteResultsSummary() 
							Exit Function
						End  IF 
					Else
						fFillAutomaticPaymentTransactionsAdditionalLog = False
						Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Additional Log data","Automatic Payment Transactions Additional Log details are not filled Properly")
						Call fRptWriteResultsSummary() 
						Exit Function
					End  IF 
				Else
					fFillAutomaticPaymentTransactionsAdditionalLog = False
					Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Additional Log data","Automatic Payment Transactions Additional Log details are not filled Properly")
					Call fRptWriteResultsSummary() 
					Exit Function
				End  IF 
		Else
			fFillAutomaticPaymentTransactionsAdditionalLog = False
			Call fRptWriteReport("Fail","Fill Automatic Payment Transactions Additional Log data","Automatic Payment Transactions Additional Log details are not filled Properly")
			Call fRptWriteResultsSummary() 
			Exit Function
		End  IF 	
		'Clear object
		Set objGLFrame = Nothing
	On error goto 0				
End Function	

'***************************************************************************************************************************************
''	Function Name					:				fSaveAutomaticPaymentTransactionDetails
''	Objective						:				Save Automatic Payment Transaction Details
''	Input Parameters				:				objDataDict,iRowCountRef
''	Output Parameters			    :				Nil
''	Date Created					:				13/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fSaveAutomaticPaymentTransactionDetails(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim objGLFrame
	Dim strMsg
	Dim SaveParametersMsg
		'Read data from excel
		strMsg =  objDataDict.Item("SaveDataMsg" & iRowCountRef)
		strSaveParametersMsg = objDataDict.Item("SaveParametersMsg" & iRowCountRef)
		Set objGLFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions")
		Call fClick(objGLFrame.WebElement("weStatus"),"Status Tab")
		strInfoMsg = Lcase(fGetText(objGLFrame.WebElement("weSavedata"),"innertext","Information Message"))
			If Instr(1,strInfoMsg,Lcase(strMsg)) Then
				Call fClick(objGLFrame.SAPButton("btnSaveDataYes"),"Yes")
			Else
				Call fRptWriteReport("Fail" ,"Verify message details in Popup window",strMsg&" - Expected message details are not appeared in Popup window")		
			End If
		'Verify Message details in Status section
		Call fVerifyStatusInAutomaticPaymentTransactionStatusTable(strSaveParametersMsg)
		'Clear Object
		Set objGLFrame = Nothing
	On error goto 0	
End Function



'***************************************************************************************************************************************
''	Function Name					:				fScheduleProposalAutomaticPaymentTransaction
''	Objective						:				Schedule Proposal Automatic Payment Transaction
''	Input Parameters				:				objDataDict,iRowCountRef
''	Output Parameters			    :				Nil
''	Date Created					:				13/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fScheduleProposalAutomaticPaymentTransaction(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim objGLFrame
	Dim blnStartImmediately
	Dim strMsg
	'Read data from excel
	blnStartImmediately = objDataDict.Item("StartImmediately" & iRowCountRef)
	strMsg = objDataDict.Item("PaymentProMsg" & iRowCountRef)
	Set objGLFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions")
		If fVerifyObjectExist(objGLFrame.SAPButton("btnProposal")) Then
			Call fClick(objGLFrame.SAPButton("btnProposal"),"Proposal Tab")
			'Wait till Popup exist and Start Immediately check box enabled 
			Call fSynUntilObjExists(objGLFrame.SAPCheckBox("chkStartImmediately"),MAX_WAIT)
			'Check Start Immediately check box 
			Call fSelect(objGLFrame.SAPCheckBox("chkStartImmediately"),blnStartImmediately,"Start Immediately")
			'Click Proposal 
			Call fClick(objGLFrame.SAPButton("btnSchedule"),"Proposal Tab")
			'Verify Message details in Status section
			Call fVerifyStatusInAutomaticPaymentTransactionStatusTable(strMsg)
		Else
			Call fRptWriteReport("Fail" ,"Verify Automatic Payment Transaction","Proposal button not displayed in  Automatic Payment Transaction")	
			Call fRptWriteResultsSummary() 
			Exit Function
		End If		
	'Clear Object
	Set objGLFrame = Nothing
	On error goto 0	
End Function
'***************************************************************************************************************************************
''	Function Name					:				fPaymentRunAutomaticPaymentTransaction
''	Objective						:				Carried Payment
''	Input Parameters				:				objDataDict,iRowCountRef
''	Output Parameters			    :				Nil
''	Date Created					:				13/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fPaymentRunAutomaticPaymentTransaction(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim objGLFrame
	Dim blnStartImmediately
	Dim strMsg
			'Read data from excel
			blnStartImmediately = objDataDict.Item("StartImmediately" & iRowCountRef)
			strMsg = objDataDict.Item("PaymentRunMsg" & iRowCountRef)
			Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions"),MIN_WAIT)
			Set objGLFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions")
				If fVerifyObjectExist(objGLFrame.WebElement("weParameter")) Then 
					'Click Parameter and  Status - to refresh page
					Call fClick(objGLFrame.WebElement("weParameter"),"Parameter Tab")
					Call fClick(objGLFrame.WebElement("weStatus"),"Status Tab")
					'Wait Payment Run button exist
					Call fSynUntilObjExists(objGLFrame.SAPButton("btnPaymentRun"),MIN_WAIT)
					Call fClick(objGLFrame.SAPButton("btnPaymentRun"),"Payment Run")
					'Wait till Popup exist and Start Immediately check box enabled 
					Call fSynUntilObjExists(objGLFrame.SAPCheckBox("chkStartImmediately"),MIN_WAIT)
					'Check Start Immediately check box 
					Call fSelect(objGLFrame.SAPCheckBox("chkStartImmediately"),blnStartImmediately,"Start Immediately")
					'Click Proposal 
					Call fClick(objGLFrame.SAPButton("btnSchedule"),"Schedule")
					'Verify Message details in Status section
					Call fVerifyStatusInAutomaticPaymentTransactionStatusTable(strMsg)
				Else
					Call fRptWriteReport("Fail" ,"Verify Payment Transaction status","parameter not displayed in Payment page")	
					Call fRptWriteResultsSummary() 
					Exit Function
				End If
		'Clear Object
		Set objGLFrame = Nothing
	On error goto 0		
End Function

'***************************************************************************************************************************************
''	Function Name					:				fVerifyStatusInAutomaticPaymentTransactionStatusTable
''	Objective						:				Verify Status in Automatic Payment Transaction Status section
''	Input Parameters				:				strExpectedMessage
''	Output Parameters			    :				Nil
''	Date Created					:				13/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fVerifyStatusInAutomaticPaymentTransactionStatusTable(strExpectedMessage)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim objGLFrame
	Dim strMessageEV
		'Get data from Excel sheet / Input from function
		strMessageEV=strExpectedMessage
		Set objGLFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions")
		'Click Status& Parameter button
		Call fClick(objGLFrame.WebElement("weParameter"),"Parameter")
		objGLFrame.Sync
		Call fClick(objGLFrame.WebElement("weStatus"),"Status Tab")
		objGLFrame.Sync
		'Wait Payment Run button exist
		Call fSynUntilObjExists(objGLFrame.WebTable("tblFinalStatusAutoPayTras"),MIN_WAIT)
		'Get data from table and compare Actual value and expected value
		strMessageAV=fWebtableGetCelldata(objGLFrame.WebTable("tblFinalStatusAutoPayTras"),1,1,"Automatic Payment Transactions Status")
			If instr(1,Lcase(strMessageAV),Lcase(strMessageEV)) > 0 Then
				Call fRptWriteReport("PASSWITHSCREENSHOT" ,"Verify message details in Status section",strMessageEV&"  message details are displayed")
			Else
				Call fRptWriteReport("Fail" ,"Verify message details in Status section",strMessageEV&"  message details are not displayed")	
			End If
		'Clear Object
		Set objGLFrame = Nothing
	On error goto 0		
End Function
'**************************************************************************************************************************************
''	Function Name					:				    fFillAutomaticPaymentTransactionsFreeSelection
''	Objective						:					Fill Automatic Payment Transactions-Free Selection details
''	Input Parameters				:					objDataDict,iRowCountRef
''	Output Parameters			    :					NIl
''	Date Created					:					14/May/2020
''	QTP/UFT Version					:					15
''	Pre-requisites					:					NIL  
''	Created By						:					Cigniti
''	Modification Date		        :		   			05/28/2020 -
'**************************************************************************************************************************************
Public Function fFillAutomaticPaymentTransactionsFreeSelection(objDataDict,iRowCountRef)
		On error resume next
		'Verify if Step Failed, If yes, it will not run the function
	    If Environment("StepFailed") = "YES" Then
			Exit Function
		End If	
		'Declarations
		Dim objFrame
		Dim strFieldLabelName
		Dim strInvoiceDocumentNumber
		Dim strCreditMemoDocumentNumber
		Dim intColumnNo
		Dim intRowNoDocNo
		'Read data from excel
		strFieldLabelName = objDataDict.Item("FieldLabelName" & iRowCountRef)
		strInvoiceDocumentNumber  = fGetSingleValue("InvoiceDocumentNumber","TestData",Environment("TestName")) 
		strCreditMemoDocumentNumber  = fGetSingleValue("CreditMemoDocumentNumber","TestData",Environment("TestName"))
		Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions"),MIN_WAIT)
		'Object for till frame
		Set objFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions")
		Call fClick(objFrame.WebElement("weFreeSelection"),"Free Selection Tab")
		'Wait till object exist
		Call fSynUntilObjExists(objFrame.SAPEdit("sFreeSelectionFiledName"),MIN_WAIT)
			If fVerifyObjectExist(objFrame.SAPEdit("sFreeSelectionFiledName")) Then	
				Call fClick(objFrame.SAPEdit("sFreeSelectionFiledName"),"Free Selection Filed Name")
				'Wait till object exist
				Call fSynUntilObjExists(objFrame.SAPButton("sFieldNameButton"),MIN_WAIT)
				Call fClick(objFrame.SAPButton("sFieldNameButton"),"Field Name")
				'Wait till object exist
				Call fSynUntilObjExists(objFrame.WebTable("tblResttictionsData"),MIN_WAIT)
				'Get Column number
				intColumnNo = fGetTableHeaderColumnNumber(objFrame.WebTable("tblRestrictionsHeader"),1,1,"Field Label")
				'Get Row number
				intRowNoDocNo = fGetRowNumberInTableBasedonColumnData (objFrame.WebTable("tblResttictionsData"),intColumnNo,strFieldLabelName) 
				Call fClick(objFrame.WebTable("tblResttictionsData").ChildItem(intRowNoDocNo,intColumnNo,"WebElement",0),strFieldLabelName)
				Call fClick(objFrame.SAPButton("btnFreeSelectionCopy"),"Copy") ' Click COPY
				'Wait till object exist
				Call fSynUntilObjExists(objFrame.SAPEdit("sFieldNameValues"),MAX_WAIT)
				'Enter Value 1 and Value 2
				strInvoiceDocumentNumber = Replace(strInvoiceDocumentNumber,"@",",")
				
				Call fEnterText(objFrame.SAPEdit("sFieldNameValues"),strInvoiceDocumentNumber,"Field Name Value1")
				If strCreditMemoDocumentNumber <> "" and  Not ISempty(strCreditMemoDocumentNumber) and  UCASE(strCreditMemoDocumentNumber) <> "EMPTY"  Then
					Call fEnterText(objFrame.SAPEdit("sFieldNameValues2"),strCreditMemoDocumentNumber,"Field Name Value2")
				End If
				
				Call fRptWriteReport("PASSWITHSCREENSHOT","Fill Automatic Payment Transactions Free Selection","Details are filled in Free Selection page")
			Else
				Call fRptWriteReport("Fail", "Verify Payment Transactions Free Selection","Free Selection Filed Name not displayed")
			End If
		Set objFrame = Nothing
		On error goto 0	
End Function 		

'***************************************************************************************************************************************
''	Function Name					:				fFillPaymentProposalWorkflowSelectionValues
''	Objective						:				Fill Enter Selection Values
''	Input Parameters				:				objDataDict,iRowCountRef
''	Output Parameters			    :				Nil
''	Date Created					:				14/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fFillPaymentProposalWorkflowSelectionValues(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Declarations
	Dim objFrame
	Dim intRunDateFrom
	Dim intRunDateTo
	Dim strIdentificationFrom
	Dim strIdentificationTo
		'Read data from excel
		intRunDateFrom = objDataDict.Item("PaymentPropWorkflowRunDateFrom" & iRowCountRef)
		strIdentificationFrom = fGetSingleValue("Identification","TestData",Environment("TestName")) 
		Set objPage = Browser("brFiori").Page("pgFiori")
		Set objFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmPaymentProposalWorkflow")
	
'		Call fSynUntilObjExists(objPage.WebButton("btnPaymentProposalWorkflow"),MID_WAIT) 
'		Call fSynUntilObjExists(objPage.WebButton("btnPaymentProposalWorkflow"),MID_WAIT)
'		Call fClick(objPage.WebButton("btnPaymentProposalWorkflow"),"Payment Proposal Workflow")
		Call fSynUntilObjExists(objFrame.SAPEdit("txtRunDate"),MID_WAIT)
		Call fSynUntilObjExists(objFrame.SAPEdit("txtIdentificationTo"),MID_WAIT)
				'Enter Run date
			If Lcase(intRunDateFrom)<> "n/a" and Lcase(intRunDateFrom)<> "null" and Lcase(intRunDateFrom)<> "no" Then
				intRunDateFrom=fGetFutureDateAdd(intRunDateFrom)
				Call fEnterText(objFrame.SAPEdit("txtRunDate"),intRunDateFrom,"Run Date From")
				Wait(2)
			End If
			'Enter Identification
			If Lcase(strIdentificationFrom)<> "n/a" and Lcase(strIdentificationFrom)<> "null" and Lcase(strIdentificationFrom)<> "no" Then
			
				Call fEnterText(objFrame.SAPEdit("txtIdentification"),strIdentificationFrom,"Identification From")
				Wait(2)
			End If
			' For Future Purpouse #########################################################################################################
'			intRunDateTo = objDataDict.Item("PaymentPropWorkflowRunDateTO" & iRowCountRef)
'			strIdentificationTo = objDataDict.Item("PaymentPropWorkflowIdentificationTO" & iRowCountRef)
'			If Lcase(intRunDateTo)<> "n/a" and Lcase(intRunDateTo)<> "null" and Lcase(intRunDateTo)<> "no" Then
'				intRunDateTo=fGetFutureDateAdd(intRunDateTo)
'				Call fEnterText(objFrame.SAPEdit("txtRunDateTo"),intRunDateTo,"Run Date To")
'			End If
			
'			If Lcase(strIdentificationTo)<> "n/a" and Lcase(strIdentificationTo)<> "null" and Lcase(strIdentificationTo)<> "no" Then
'				Call fEnterText(objFrame.SAPEdit("txtIdentificationTo"),strIdentificationTo,"Identification To")
'			End If
			'################################################################################################################################
		'CLick Execute
		Call fClick(objFrame.SAPButton("Execute"),"Execute")
		'Wait till object exist
		Call fSynUntilObjExists(objFrame.WebTable("tblPayProWorkflow"),MAX_WAIT)
		intPayPropCount = objFrame.WebTable("tblPayProWorkflow").RowCount
		For iCount = 1 To intPayPropCount
			stridentificatioValue = objFrame.WebTable("tblPayProWorkflow").GetCellData(iCount,4)
			If Trim(stridentificatioValue) = Trim(strIdentificationFrom) Then
				objFrame.WebTable("tblPayProWorkflow").ChildItem(iCount,1,"SAPCheckBox",0).set "ON"
				Exit For
			End If			
		Next
		Call fClick(objFrame.SAPButton("btnSetStatusREADY"),"Set Status READY")
		wait (MIN_WAIT)
		blnStatus = False
		For iCount = 1 To intPayPropCount
			strStatus = objFrame.WebTable("tblPayProWorkflow").ChildItem(iCount,5,"WebList",0).getroproperty("innerhtml")		
			If  strStatus = "Ready" Then
				blnStatus = True			
				Exit For
			End If			
		Next	
		
		Call fSynUntilObjExists(objFrame.WebTable("tblPayProWorkflow"),MIN_WAIT)
			If blnStatus Then
				Call fRptWriteReport("PASSWITHSCREENSHOT", "Verification of Payment Proposal Workflow","Payment Proposal Workflow has been changed to Ready")
			Else
				Call fRptWriteReport("Fail",  "Verification of Payment Proposal Workflow","Payment Proposal Workflow has not been changed to Ready")
				Call fRptWriteResultsSummary() 
			Exit Function
			End If
		wait 5
		Call fFioriLeavePage()
		'Clear object
		Set objFrame = Nothing
	On error goto 0			
End Function	

'***************************************************************************************************************************************
''	Function Name					:				fFioriVIMWorkplaceForProcessPOProcessing
''	Objective						:				VIM Work place For Process PO Processing
''	Input Parameters				:				objDataDict,iRowCountRef
''	Output Parameters			    :				Nil
''	Date Created					:				14/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fFioriVIMWorkplaceForProcessPOProcessing(objDataDict,iRowCountRef,strInvoiceOrCreditMemo)
	On error resume next
	
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim strSwitchWorkView
	Dim strComment
	Dim intBalance
	Dim strManagerApproval
	Dim objBrwAndPg
	Dim blnExecute
	Dim iRowC
	
	strSwitchWorkView = objDataDict.Item("SwitchWorkView" & iRowCountRef)
	strComment = objDataDict.Item("BypassComment" & iRowCountRef)
	intBalance = objDataDict.Item("Balance" & iRowCountRef)
	strTaxCode = objDataDict.Item("TaxCode" & iRowCountRef)
	strUpdatedPrice = objDataDict.Item("Credit Memo Amount" & iRowCountRef)
    strupdateQuanity= fGetSingleValue("Credit Memo Quantity","TestData",Environment("TestName"))    
    strMappingInvoice = fGetSingleValue("AutoInvoiceNumber","TestData",Environment("TestName"))
	
	
	Set objBrwAndPg = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
'	Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk"),MID_WAIT)
	objBrwAndPg.sync
'	objBrwAndPg.Highlight

	
	Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnSwitchWorkView"),MIN_WAIT)
'	Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnSwitchWorkView"),MIN_WAIT)
'	Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnSwitchWorkView"),MIN_WAIT)
	Call fClick(objBrwAndPg.SAPButton("btnSwitchWorkView"),"Switch Work View")
	Call fSynUntilObjExists(objBrwAndPg.SAPRadioGroup("rbSwitchWorkView"),MID_WAIT)
	Call fSelect(objBrwAndPg.SAPRadioGroup("rbSwitchWorkView"),strSwitchWorkView,"Switch Work View")
	Call fClick(objBrwAndPg.SAPButton("btnContinue"),"Continue Switch Work View")
	blnExecute =  fFioriExecteInvoice(objBrwAndPg,strInvoiceOrCreditMemo)

	If blnExecute <> "" Then
			
			If strInvoiceOrCreditMemo = "creditmemo" Then
			strMappingInvoice = fGetSingleValue("AutoInvoiceNumber","TestData",Environment("TestName"))
    		Call fEnterText(objBrwAndPg.SAPFrame("SAPFrame").SAPEdit("txtCMReferenceNumber"),strMappingInvoice,"CM Reference Number")
    		End If
    		
    			Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnSimulateRules"),2)
			If fVerifyObjectExist(objBrwAndPg.SAPButton("btnSimulateRules")) Then
    			Call fClick(objBrwAndPg.SAPButton("btnSimulateRules"),"Simulate Rules")
				' Simulate Business Rules
				Call fFioriExceptionValidationForSimulateBusinessRules(objDataDict,iRowCountRef,objBrwAndPg,strComment)
			End If
			
			Call fSynUntilObjExists(objBrwAndPg.WebElement("weLineItems"),2)
			If fVerifyObjectExist(objBrwAndPg.WebElement("weLineItems")) Then
			Call fClick(objBrwAndPg.WebElement("weLineItems"),"Line Items")
			Call fSynUntilObjExists(objBrwAndPg.WebTable("tblLineItems"),MIN_WAIT)
			
			intPODoc = fFioriPurchasingDocID()
			intLineItemsRows = objBrwAndPg.WebTable("tblLineItems").RowCount()
				' Update Line item details for more than 1 line item	
				set objBrwAndPg =  Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
				intLineItemsRows =  objBrwAndPg.WebTable("tblLineItems").RowCount
				For iRowC = 1 To intLineItemsRows
                        intPurchasingDoc = objBrwAndPg.WebTable("tblLineItems").ChildItem(iRowC,5,"SAPEdit",0).getroproperty("value")
                        intDocNo    =    objBrwAndPg.WebTable("tblLineItems").GetCellData(iRowC,1)
                            If intPurchasingDoc = intPODoc and intDocNo <> "" Then
                            
                            	If strInvoiceOrCreditMemo = "invoice" and strTaxCode <> "" Then		 
	                                objBrwAndPg.WebTable("tblLineItems").ChildItem(iRowC,10,"WebList",0).highlight
	                                Wait(2)
	                                objBrwAndPg.WebTable("tblLineItems").ChildItem(iRowC,14,"WebList",0).highlight
	                               	wait(2)
	                                objBrwAndPg.WebTable("tblLineItems").ChildItem(iRowC,14,"WebList",0).click
	                                wait(4)
	                                
	                                objBrwAndPg.WebTable("tblLineItems").ChildItem(iRowC,14,"SAPList",0).Select strTaxCode
	                                Wait(5)
	                                'Exit For
	                                If objBrwAndPg.WebTable("tblLineItems").GetCellData(iRowC+1,1) = "" Then
	                                	Exit For
	                                End If
	                             ElseIf strInvoiceOrCreditMemo = "creditmemo" and strupdateQuanity <> "" Then
						                             
	                               objBrwAndPg.WebTable("tblLineItems").ChildItem(iRowC,9,"WebList",0).click
	                            	wait 2
	                            	objBrwAndPg.WebTable("tblLineItems").ChildItem(iRowC,9,"SAPEdit",0).set strUpdatedPrice
	                            	wait 2
	                               objBrwAndPg.WebTable("tblLineItems").ChildItem(iRowC,10,"WebList",0).Click  
	                            	wait 2
	                            	objBrwAndPg.WebTable("tblLineItems").ChildItem(iRowC,10,"SAPEdit",0).set strupdateQuanity
                            		wait 2 
                            		 'Exit For
	                                If objBrwAndPg.WebTable("tblLineItems").GetCellData(iRowC+1,1) = "" Then
	                                	Exit For
	                                End If
                                End If
                                
                            End If
                    Next
                    
                    Call fClick(objBrwAndPg.SAPButton("btnSave"),"Save ")                 
				
				Call fSynUntilObjExists(objBrwAndPg.SAPEdit("txtDocumentBalance"),MID_WAIT)
				Wait(3)
				If "0.00" <> fGetRoProperty(objBrwAndPg.SAPEdit("txtDocumentBalance"),"value","Document Balance")Then
					
					Call fUpdateBasicDataTaxCode(objDataDict,iRowCountRef)
				End  If



	            End  If
			'*************************************************************************************************
			Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnApplyRules"),2)
			If fVerifyObjectExist(objBrwAndPg.SAPButton("btnApplyRules")) Then
				Call fClick(objBrwAndPg.SAPButton("btnApplyRules"),"Apply Rules")
				Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnReset"),2)
			If fVerifyObjectExist(objBrwAndPg.SAPButton("btnReset")) Then
				Call fClick(objBrwAndPg.SAPButton("btnReset"),"Reset")
				'strManagerApproval = objDataDict.Item("APAnalystApproval" & iRowCountRef)
				' Need approval from Analyst or Manager
				strManagerApproval = fGetSingleValue("APAnalystApproval","TestData",Environment("TestName")) ' Get Data from excel sheet 
				If Ucase(strManagerApproval) ="YES" Then
					Call fFioriApproveInvoice(objDataDict,iRowCountRef,strInvoiceOrCreditMemo)
				End If
			End If
			
			strManualCheckNonUSANInvoices= fGetSingleValue("ManualCheckNonUSANInvoices","TestData",Environment("TestName")) ' Get Data from excel sheet
				If Ucase(strManualCheckNonUSANInvoices) ="YES" Then
					Call fFioriProcessInvoiceManually(strInvoiceOrCreditMemo)
				End If
			
			End If
			
'			If fVerifyObjectExist(objBrwAndPg.SAPButton("btnApprove")) Then
'				Call fFioriProcessInvoiceManually(objDataDict,iRowCountRef,strInvoiceOrCreditMemo)
'			End If

			' VIM Process completed
			'strInvoiceOrCreditMemo
			If strInvoiceOrCreditMemo ="creditmemo"  Then
				Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"DONE","TestData","ProcessCreditMemoVIM")
			ElseIf strInvoiceOrCreditMemo ="invoice" Then
				Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"DONE","TestData","ProcessInvoiceVIM")
			End If

			Else
				Call fCLick(objBrwAndPg.WebElement("weTabAllCompleted"),"All Completed")
				Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnReset"),MID_WAIT)
				Call fClick(objBrwAndPg.SAPButton("btnReset"),"Reset")
				Call fSynUntilObjExists(objBrwAndPg.SAPEdit("txtReference"),MID_WAIT)
				Call fFioriExecteInvoice(objBrwAndPg,strInvoiceOrCreditMemo)
				If UCase(fFioriGetInvoiceORCreditMemoDocumentStatus) = "POSTED" Then
					If strInvoiceOrCreditMemo ="creditmemo"  Then
						Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"DONE","TestData","ProcessCreditMemoVIM")
					ElseIf strInvoiceOrCreditMemo ="invoice" Then
						Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"DONE","TestData","ProcessInvoiceVIM")
					End If
				End If

			End If
			'Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"DONE","TestData","ProcessInvoiceVIM")
	
	On error goto 0	
End Function

'***************************************************************************************************************************************
''	Function Name					:				fFioriApproveInvoice
''	Objective						:				Approve Invoice
''	Input Parameters				:				objBrwAndPg
''	Output Parameters			    :				Nil
''	Date Created					:				14/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fFioriApproveInvoice(objDataDict,iRowCountRef,strInvoiceOrCreditMemo)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim strInvoiceID
	'Get the Invoice Number from testdata sheet
	Set objBrwAndPg = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
	strInvoiceID = fGetSingleValue("AutoInvoiceNumber","TestData",Environment("TestName")) 
	
		If fVerifyObjectExist(objBrwAndPg.SAPEdit("txtReference")) Then
			Call fFioriExecteInvoice(objBrwAndPg,strInvoiceOrCreditMemo)
		End If
		
		Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnApprove"),MAX_WAIT)
			If fVerifyObjectExist(objBrwAndPg.SAPButton("btnApprove")) Then
				
				Call fClick(objBrwAndPg.SAPButton("btnApprove"),"Approve")
				
				Call fSynUntilObjExists(objBrwAndPg.WebTable("tblApproval"),MID_WAIT)
				blnManager = False
				If objBrwAndPg.WebTable("tblApproval").Exist(1) Then
					
					If objBrwAndPg.WebTable("tblApproval").GetCellData(1,1) <> "" Then
						blnManager = True				
					End If 
					
				End If
				
				If blnManager Then
					Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"YES","TestData","APAnalystApproval")
				Else
					Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"NO","TestData","APAnalystApproval")
				End If
				

'					Call fSynUntilObjExists(objBrwAndPg.WebElement("weApproveComment"),MAX_WAIT)
'					Call fClick(objBrwAndPg.WebElement("weApproveComment"),"Approve Comment")
				Call fSynUntilObjExists(objBrwAndPg.SAPEdit("txtApproveComment"),MAX_WAIT)
				If Len(strComment) <2 Then 
					strComment ="Approve"
				End If
				Call fEnterText(objBrwAndPg.SAPEdit("txtApproveComment"),strComment,"Approve Comment")
				Call fClick(objBrwAndPg.SAPButton("btnApprove"),"Approve")
				Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnSwitchWorkView"),MIN_WAIT)
				' Verify Document Status '06/16/2020 - Ratnakar Eda

				strManagerApproval = fGetSingleValue("APAnalystApproval","TestData",Environment("TestName")) ' Get 
			If Ucase(strManagerApproval) ="YES" Then
				
				' Approve Invoice details with manager role
				 'Logout
				Call fFioriLogOut()     

				' Login as manager and Approve 
				Call fFioriLogin("APManager")
				'Page Navigation 
				strSearchPageName = objDataDict.Item("WorkSpace" & iRowCountRef)
				If Ucase(strSearchPageName) = Ucase("VIM workplace") Then
				Else
					strSearchPageName = "VIM workplace"
				End If
				strTileName = objDataDict.Item("VIMTileName" & iRowCountRef)
				Call fFioriFetchAppFromHomePage(strSearchPageName,strTileName)
				strInvoiceID = fGetSingleValue("AutoInvoiceNumber","TestData",Environment("TestName")) 
				Call fFioriExecteInvoice(objBrwAndPg,strInvoiceOrCreditMemo)
				
				Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnApprove"),MAX_WAIT)
				If fVerifyObjectExist(objBrwAndPg.SAPButton("btnApprove")) Then
					Call fClick(objBrwAndPg.SAPButton("btnApprove"),"Approve")
	'				Call fSynUntilObjExists(objBrwAndPg.WebElement("weApproveComment"),MAX_WAIT)
	'				Call fClick(objBrwAndPg.WebElement("weApproveComment"),"Approve Comment")
					Call fSynUntilObjExists(objBrwAndPg.SAPEdit("txtApproveComment"),MAX_WAIT)
						If Len(strComment) <2 Then
							strComment ="Approve"
						End If
							Call fEnterText(objBrwAndPg.SAPEdit("txtApproveComment"),strComment,"Approve Comment")
							Call fClick(objBrwAndPg.SAPButton("btnApprove"),"Approve")
							wait MIN_WAIT
'							Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnSwitchWorkView"),MID_WAIT)
							'Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"DONE","TestData","ProcessInvoiceVIM")
				Else
					Call fRptWriteReport("Fail","Verification of invoice approval", "APAnalyst - Approval button not displayed")
					 
					Exit Function
				End If
			Else		
				'Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"DONE","TestData","ProcessInvoiceVIM")
				 
				Exit Function
			End If	
			
		End If		
End Function
'***************************************************************************************************************************************
''	Function Name					:				fFioriExecteInvoice
''	Objective						:				Execte Invoice
''	Input Parameters				:				objBrwAndPg
''	Output Parameters			    :				Nil
''	Date Created					:				14/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'**************************************************************************************************************************************
Public Function fFioriExecteInvoice(objBrwAndPg,strInvoiceOrCreditMemo)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim strInvoiceID
	Dim intCount
	Dim intExecute
	'set frame object
	If strInvoiceOrCreditMemo = "invoice" Then
		strInvoiceID = fGetSingleValue("AutoInvoiceNumber","TestData",Environment("TestName")) 
	Else
		strInvoiceID = fGetSingleValue("AutoCreditMemoNumber","TestData",Environment("TestName")) 	
	End If
	
	Call fSynUntilObjExists(objBrwAndPg.SAPEdit("txtReference"),MID_WAIT)
		If fVerifyObjectExist(objBrwAndPg.SAPEdit("txtReference")) Then
			Call fSynUntilObjExists(objBrwAndPg.SAPEdit("txtReference"),MID_WAIT)
			Wait(5)
			Call fEnterText(objBrwAndPg.SAPEdit("txtReference"),strInvoiceID,"Invoice Reference")
			Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnApply"),MID_WAIT)
			Call fClick(objBrwAndPg.SAPButton("btnApply"),"Apply")	
			Call fSynUntilObjExists(objBrwAndPg.WebTable("tblAllUsersViewExecute"),MID_WAIT)
			Wait(2)
			intCount = 0
				Do while  objBrwAndPg.WebTable("tblAllUsersViewExecute").GetROProperty("innertext") = ""
					Wait(MAX_WAIT)
					Call fClick(objBrwAndPg.SAPButton("btnReset"),"Reset")
					Call fSynUntilObjExists(objBrwAndPg.SAPEdit("txtReference"),MID_WAIT)
					Wait(MIN_WAIT)
					Call fEnterText(objBrwAndPg.SAPEdit("txtReference"),strInvoiceID,"Invoice Reference")
					Call fClick(objBrwAndPg.SAPButton("btnApply"),"Apply")
					Wait(MIN_WAIT)
						If intCount = 5 Then
							Exit do
						End If
					intCount = intCount + 1
				loop
			intExecute = objBrwAndPg.WebTable("tblAllUsersViewExecute").RowCount()
			'fFioriExecteInvoice = False 
			fFioriExecteInvoice = objBrwAndPg.WebTable("tblAllUsersViewExecute").GetROProperty("innertext")
				For intIter  = 1 To intExecute
					strExecuteBtn = objBrwAndPg.WebTable("tblAllUsersViewExecute").ChildItem(intIter,3,"WebElement",0).GetRoProperty("innertext")
						If strExecuteBtn = "Execute Workitem" Then
							objBrwAndPg.WebTable("tblAllUsersViewExecute").ChildItem(intIter,3,"WebElement",0).Click
							'fFioriExecteInvoice = True
							Exit For
						End If
				Next
'				If not fFioriExecteInvoice Then
'					Call fRptWriteReport("Fail", "Verify invoice record in vim page","Invoice record not displayed in vim page")
'					Call fRptWriteResultsSummary() 
'					Exit Function
'				End If
		End If
		On error goto 0	
End Function

'****************************************************************************************************************
''    Function Name                    :                fFioriExceptionValidationForSimulateBusinessRules
''    Objective                        :                Exception Validation For Simulate Business Rules
''    Input Parameters                 :                 sObjectName,strComment
''    Output Parameters                :                Nil
''    Date Created                     :                15/May/2020
''    UFT/QTP Version                  :                15.0
''    Pre-requisites                   :              NIL  
''    Created By                       :              Cigniti
''    Modification Date                :                   
'*****************************************************************************************************************
Public Function fFioriExceptionValidationForSimulateBusinessRules(objDataDict,iRowCountRef,sObjectName,strComment)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	

    Dim intSimulateRules
    Dim intIter
    Dim intGSTRegistrationNumber
    fFioriExceptionValidationForSimulateBusinessRules = False
    
    intGSTRegistrationNumber = objDataDict.Item("GST Registration Number" & iRowCountRef)
    If intGSTRegistrationNumber <> "" Then
    	Call fFioriUpdateGSTRegistrationNumber(objDataDict,iRowCountRef)
    End If   
    
    Call fSynUntilObjExists(sObjectName.WebTable("tblException"),MID_WAIT)
        If fVerifyObjectExist(sObjectName.WebTable("tblException")) Then
            Call fSynUntilObjExists(sObjectName.WebTable("tblException"),MID_WAIT)
            intSimulateRules = sObjectName.WebTable("tblException").RowCount()
                For intIter = 2 To intSimulateRules
              
                'Manual Check on Non-US Ariba Network Invoices (PO) - Out OF USA '06/19/2020- Ratnakar Eda 
                 If Ucase(Trim("Manual Check on Non-US Ariba Network Invoices (PO)")) = Ucase(Trim(sObjectName.WebTable("tblException").GetCellData(intIter,2))) Then 
                  	If "Exception occured" = sObjectName.WebTable("tblException").ChildItem(intIter,3,"WebElement",0).GetRoProperty("innertext") Then    
                        If "Bypass" = Trim(sObjectName.WebTable("tblException").ChildItem(intIter,5,"SAPButton",0).GetROProperty("name"))  Then
		                 	'Write Value in Test data sheet
		     				Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"YES","TestData","ManualCheckNonUSANInvoices")
     					End  IF
     				End  IF	
                 End  IF
                 
                   ' Need Approval Required from Analyst or Manager '05/30/2020- Ratnakar Eda 
                   If Ucase(Trim("Approval Required (PO)")) = Ucase(Trim(sObjectName.WebTable("tblException").GetCellData(intIter,2))) Then 
                  	If "Exception occured" = sObjectName.WebTable("tblException").ChildItem(intIter,3,"WebElement",0).GetRoProperty("innertext") Then    
                        If "Bypass" <> Trim(sObjectName.WebTable("tblException").ChildItem(intIter,5,"SAPButton",0).GetROProperty("name"))  Then
		                 	'Write Value in Test data sheet
		     				Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,"YES","TestData","APAnalystApproval")
     					End  IF
     				End  IF	
                 End  IF
                 
                 
                    If "Exception occured" = sObjectName.WebTable("tblException").ChildItem(intIter,3,"WebElement",0).GetRoProperty("innertext") Then    
                        If "Bypass" = Trim(sObjectName.WebTable("tblException").ChildItem(intIter,5,"SAPButton",0).GetROProperty("name"))  Then
                            sObjectName.WebTable("tblException").ChildItem(intIter,5,"SAPButton",0).Click 
                            Call fSynUntilObjExists(sObjectName.SAPButton("btnYes"),MID_WAIT)
                            Call fClick(sObjectName.SAPButton("btnYes"),"Yes")
                            Call fSynUntilObjExists(sObjectName.WebElement("weCommentSection"),MID_WAIT)
                            Call fClick(sObjectName.WebElement("weCommentSection"),"Comment Section")
                            Call fSynUntilObjExists(sObjectName.SAPEdit("txtSimulateBusinessRules"),MID_WAIT)
                            Call fEnterText(sObjectName.SAPEdit("txtSimulateBusinessRules"),strComment,"Simulate Business Rules")
                            Call fClick(sObjectName.SAPButton("btnPopSave"),"Save")
                            Wait(5)'Required
                            fFioriExceptionValidationForSimulateBusinessRules = TRUE
                        End If
                    End If
                Next
            Call fClick(sObjectName.SAPButton("btnExit"),"Exit")
            fFioriExceptionValidationForSimulateBusinessRules = TRUE
        Else        
            Call fRptWriteReport("Fail", "Verify Simulpate business rules","Simulate rules page not displayed")
            Call fRptWriteResultsSummary() 
            Exit Function
        End If
End Function


'****************************************************************************************************************
''	Function Name					:				fFioriPurchasingDocID
''	Objective						:				Get PO numbers based on invoice number
''	Input Parameters				:				Nil
''	Output Parameters			    :				Nil
''	Date Created					:				15/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************
Public Function fFioriPurchasingDocID()
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim strInvoiceID
	'Get the Invoice Number from testdata sheet
	strInvoiceID = fGetSingleValue("AutoInvoiceNumber","TestData",Environment("TestName")) 
		For intIter = 1 To Len(strInvoiceID)
			intNum = Mid(strInvoiceID,intIter,1)
				If IsNumeric(intNum) Then 
					fFioriPurchasingDocID = fFioriPurchasingDocID & intNum
				End If
		Next
End Function



'****************************************************************************************************************
''    Function Name                    :                fFioriFillDocumentList
''    Objective                        :                Get document numbers based on Reference number
''    Input Parameters                :                objDataDict,iRowCountRef
''    Output Parameters                :                Nil
''    Date Created                    :                15/May/2020
''    UFT/QTP Version                 :                15.0
''    Pre-requisites                    :                NIL  
''    Created By                        :                Cigniti
''    Modification Date                :                   
'*****************************************************************************************************************
Public Function fFioriFillDocumentList_OLD(objDataDict,iRowCountRef)
    On error resume next
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    Dim intCompanyCode
    Dim objBrowAndPage
    Dim objBrowAndPageAndFrm
    Dim intColumnNo
    Dim arrDocumentNumbers()
        'Get the Invoice Number from testdata sheet
        intInvoice = fGetSingleValue("AutoInvoiceNumber","TestData",Environment("TestName")) 
        intCompanyCode = objDataDict.Item("CompanyCode" & iRowCountRef)
        Call fSynUntilObjExists(Browser("brFiori").Page("pgFiori"),MID_WAIT)
        Set objBrowAndPage = Browser("brFiori").Page("pgFiori")
        Set objBrowPage = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk")
        Set objBrowAndPageAndFrm = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
        objBrowAndPage.Highlight
        Call fSynUntilObjExists(objBrowAndPageAndFrm.SAPButton("btnDocumentList"),MID_WAIT)
            If fVerifyObjectExist(objBrowAndPage.WebElement("weDisplayDocument")) Then
'                Call fClick(objBrowAndPage.WebElement("weDisplayDocument"),"Display Document")
'                Call fSynUntilObjExists(objBrowAndPageAndFrm.SAPButton("btnDocumentList"),MID_WAIT)
                Call fClick(objBrowAndPageAndFrm.SAPButton("btnDocumentList"),"Document List")
                Call fSynUntilObjExists(objBrowAndPageAndFrm.SAPEdit("txtCompanyCode"),MID_WAIT)
                Call fSynUntilObjExists(objBrowAndPageAndFrm.SAPEdit("txtCompanyCode"),MIN_WAIT)
                Call fEnterText(objBrowAndPageAndFrm.SAPEdit("txtCompanyCode"),intCompanyCode,"Company Code")
                Call fEnterText(objBrowAndPageAndFrm.SAPEdit("txtReferenceNumber"),intInvoice,"Reference Number")
                Call fClick(objBrowAndPageAndFrm.SAPButton("btnExecute"),"Execute")
                Call fSynUntilObjExists(objBrowAndPageAndFrm.WebTable("tblDocumentListHeader"),MID_WAIT)
                intcount = 0
                    Do while objBrowAndPageAndFrm.WebTable("tblDocumentListData").RowCount() <> 2
                        Wait(MAX_WAIT)
                        Call fClick(objBrowPage.WebButton("btnBack"),"Back")
                        Wait(MID_WAIT)
                        Call fClick(objBrowAndPageAndFrm.SAPButton("btnExecute"),"Execute")
                        Wait(MID_WAIT)
                         If intcount = 10 Then
                             Exit do
                         End If
                        intcount = intcount + 1
                    loop
                'Get Column number based on Column name
                intColumnNo = fGetTableHeaderColumnNumber(objBrowAndPageAndFrm.WebTable("tblDocumentListHeader"),1,1,"DocumentNo")
                'Get Row number based on Column number
                intRowCount = fGetRoProperty(objBrowAndPageAndFrm.WebTable("tblDocumentListData"),"rows","Document List Data")
                    For intIteator = 1 To intRowCount
                        ReDim preserve arrDocumentNumbers(intIteator-1)
                        arrDocumentNumbers(intIteator-1) = fGetCelldata(objBrowAndPageAndFrm.WebTable("tblDocumentListData"),intIteator,intColumnNo,"Document Numbers")
                    Next
                'Write data in excel
                Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,arrDocumentNumbers(0),"TestData","InvoiceDocumentNumber")
                Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,arrDocumentNumbers(1),"TestData","CreditMemoDocumentNumber")
                Call fFioriLeavePage()
                fFioriFillDocumentList = arrDocumentNumbers
            Else
                Call fRptWriteReport("Fail", "Verify Display Document Icon in Searched page","Display document is not displayed in Searched results page")
                Call fRptWriteResultsSummary() 
                Exit Function
            End If
        Set objBrowAndPage = Nothing
        Set objBrowAndPageAndFrm = Nothing
    On error goto 0	
End Function

Public Function fFioriFillDocumentList(objDataDict,iRowCountRef,strInvoiceType)
    On error resume next
    Dim intCompanyCode
    Dim objBrowAndPage
    Dim objBrowAndPageAndFrm
    Dim intColumnNo
    Dim arrDocumentNumbers()
        'Get the Invoice Number from testdata sheet
     If strInvoiceType = "invoice" Then
		intInvoice = fGetSingleValue("AutoInvoiceNumber","TestData",Environment("TestName")) 
	Else
		intInvoice = fGetSingleValue("AutoCreditMemoNumber","TestData",Environment("TestName")) 	
	End If

        intCompanyCode = objDataDict.Item("CompanyCode" & iRowCountRef)
        Call fSynUntilObjExists(Browser("brFiori").Page("pgFiori"),MID_WAIT)
        Set objBrowAndPage = Browser("brFiori").Page("pgFiori")
        Set objBrowPage = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk")
        Set objBrowAndPageAndFrm = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
        'objBrowAndPage.Highlight ' Ratnakar Eda -05/28/2020 - Updated
'        Call fSynUntilObjExists(objBrowAndPage.WebElement("weDisplayDocument"),MAX_WAIT)
 		Call fSynUntilObjExists(objBrowAndPageAndFrm.SAPButton("btnDocumentList"),MID_WAIT)
            If fVerifyObjectExist(objBrowAndPageAndFrm.SAPButton("btnDocumentList")) Then
                'Call fClick(objBrowAndPage.WebElement("weDisplayDocument"),"Display Document")
'                Call fSynUntilObjExists(objBrowAndPageAndFrm.SAPButton("btnDocumentList"),MID_WAIT)
                Call fClick(objBrowAndPageAndFrm.SAPButton("btnDocumentList"),"Document List")
                Call fSynUntilObjExists(objBrowAndPageAndFrm.SAPEdit("txtCompanyCode"),MID_WAIT)
                Call fEnterText(objBrowAndPageAndFrm.SAPEdit("txtCompanyCode"),intCompanyCode,"Company Code")
                Call fEnterText(objBrowAndPageAndFrm.SAPEdit("txtReferenceNumber"),intInvoice,"Reference Number")
                Call fClick(objBrowAndPageAndFrm.SAPButton("btnExecute"),"Execute")
                Call fSynUntilObjExists(objBrowAndPageAndFrm.WebTable("tblDocumentListHeader"),MID_WAIT)
                If lcase(strInvoiceType) = "invoice" and instr(1,Ucase(fGetSingleValue("PaymentTerm","TestData",Environment("TestName"))),"DD")>0 Then
					intDocumentCount = 2
				ElseIf lcase(strInvoiceType) = "invoice" and instr(1,Ucase(fGetSingleValue("PaymentTerm","TestData",Environment("TestName"))),"A0")>0 Then
					intDocumentCount = 1
				ElseIF lcase(strInvoiceType) = "creditmemo" Then
					intDocumentCount = 1
				End If	
                
'                intcount = 0
'                    Do while objBrowAndPageAndFrm.WebTable("tblDocumentListData").RowCount() <>  intDocumentCount
'                        Wait(MAX_WAIT)
'                        Call fClick(objBrowPage.WebButton("btnBack"),"Back")
'                        Wait(MID_WAIT)
'                        Call fClick(objBrowAndPageAndFrm.SAPButton("btnExecute"),"Execute")
'                        Wait(MID_WAIT)
'                         If intcount = 5 Then
'                             Exit do
'                         End If
'                        intcount = intcount + 1
'                    loop
                'Get Column number based on Column name
                intColumnNo = fGetTableHeaderColumnNumber(objBrowAndPageAndFrm.WebTable("tblDocumentListHeader"),1,1,"DocumentNo")
                'Get Row number based on Column number
                intRowCount = fGetRoProperty(objBrowAndPageAndFrm.WebTable("tblDocumentListData"),"rows","Document List Data")
                    For intIteator = 1 To intRowCount
                    	If intIteator = 1 Then
                    	    DocNoList = fGetCelldata(objBrowAndPageAndFrm.WebTable("tblDocumentListData"),intIteator,intColumnNo,"Document Numbers") 	
                    	Else
                    		DocNoList = DocNoList & "@" & fGetCelldata(objBrowAndPageAndFrm.WebTable("tblDocumentListData"),intIteator,intColumnNo,"Document Numbers") 
                    	End If
'                        ReDim preserve arrDocumentNumbers(intIteator-1)
'                        arrDocumentNumbers(intIteator-1) = fGetCelldata(objBrowAndPageAndFrm.WebTable("tblDocumentListData"),intIteator,intColumnNo,"Document Numbers")
                    Next
                'Write data in excel
'                Call fWriteOutputValueInExcel(Environment("TestName"),1,arrDocumentNumbers(0),"TestData","InvoiceDocumentNumber")
'                Call fWriteOutputValueInExcel(Environment("TestName"),1,arrDocumentNumbers(1),"TestData","CreditMemoDocumentNumber")
'                Ratnakar Eda -05/28/2020 - Updated
				If lcase(strInvoiceType) = "invoice" Then
					Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,DocNoList,"TestData","InvoiceDocumentNumber")
				ElseIf lcase(strInvoiceType) = "creditmemo" Then
					Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,DocNoList,"TestData","CreditMemoDocumentNumber")
				End If			
                
                fFioriFillDocumentList = arrDocumentNumbers
            Else
                Call fRptWriteReport("Fail", "Verify Display Document Icon in Searched page","Display document is not displayed in Searched results page")
                Call fRptWriteResultsSummary() 
                'ExitAction
            End If
            Call fFioriLeavePage()
        Set objBrowAndPage = Nothing
        Set objBrowAndPageAndFrm = Nothing
    
End Function

'****************************************************************************************************************
''	Function Name					:				fFioriLeavePage
''	Objective						:				Leave Page
''	Input Parameters				:				Nil
''	Output Parameters			    :				Nil
''	Date Created					:				15/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************
Public Function fFioriLeavePage()
    On error resume next
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    Dim objBr
    Dim objBrowAndPage
    Dim oDesc
    Dim brObj
    Dim intbrCnt
    Set objBr = Browser("brFioriAutoDesk")
    Set objBrowAndPage = Browser("brFiori").Page("pgFiori")
    'Need to add Object by using Insight
	    'Close Browser
	    If objBr.Exist(2) Then
	        objBr.Close
	        wait (2)
	        Set oDesc = Description.Create()
			oDesc("micclass").value="Browser"
			Set brObj = DeskTop.ChildObjects(oDesc)
			intbrCnt = brObj.Count
			If intbrCnt > 1 Then
	            'Click Leave button in pop up window
	            If objBr.InsightObject("btnLeave").Exist(2) Then
	                Call fClick(objBr.InsightObject("btnLeave"),"Leave")
	                Call fSynUntilObjExists(objBrowAndPage.WebButton("btnHome"),MID_WAIT)
	            End If
	        End IF    
	    End If
	    'Click on Autodesk icon -Navigate to Home page
	    If objBrowAndPage.WebButton("btnHome").Exist(1) Then
	    	Call fClick(objBrowAndPage.WebButton("btnHome"),"Home")
	    End If
    Call fSynUntilObjExists(objBrowAndPage.WebElement("weHome"),MID_WAIT)
    Set objBrowAndPage = Nothing
    Set objBr = Nothing
    Set brObj = Nothing
	Set oDesc = Nothing
    On error goto 0	
End Function

'****************************************************************************************************************
''	Function Name					:				fFioriManagerApproval
''	Objective						:				Manager Approval for invoice
''	Input Parameters				:				objDataDict,iRowCountRef
''	Output Parameters			    :				Nil
''	Date Created					:				15/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************	
Public Function fFioriManagerApproval()
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim strInvoiceID
	Dim objBrwAndPg
	'Get the Invoice Number from testdata sheet
	strInvoiceID = fGetSingleValue("AutoInvoiceNumber","TestData",Environment("TestName")) 
	Set objBrwAndPg = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
	Call fFioriApproveInvoice(objBrwAndPg)
	Call fSelect(objBrwAndPg.WebTabStrip("wtsMyTask"),"My Completed","My Task")
	Call fSynUntilObjExists(objBrwAndPg.SAPEdit("txtReference"),MID_WAIT)
	Call fEnterText(objBrwAndPg.SAPEdit("txtReference"),strInvoiceID,"Invoice Reference")
	Call fClick(objBrwAndPg.SAPButton("btnApply"),"Apply")	
	Call fSynUntilObjExists(objBrwAndPg.WebTable("tblLineItemsHeader"),MID_WAIT)
		If "Posted" = objBrwAndPg.WebTable("tblMyCompleted").ChildItem(1,15,"WebList",0).getroproperty("innerhtml") Then
			Call fRptWriteReport("PASS", "Verify Invoice approval of manager","Manager has been approved the invoice")
		Else
			Call fRptWriteReport("Fail", "Verify Invoice approval of manager","Manager has not been approved the invoice")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If		
End Function	
	
'****************************************************************************************************************
''	Function Name					:				fFioriApprovePaymentProposal
''	Objective						:				Approve Payment Proposal
''	Input Parameters				:				objDataDict,iRowCountRef
''	Output Parameters			    :				Nil
''	Date Created					:				15/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************	
Public Function fFioriApprovePaymentProposal(objDataDict,iRowCountRef)
    On error resume next
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    Dim strIdentification
    Dim objPage
    Dim objBroAndPag
    Dim objFrm
	    strIdentification = fGetSingleValue("Identification","TestData",Environment("TestName")) 
	    Set objPage = Browser("brFiori").Page("pgFiori")
	    set objBroAndPag = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk")
	    Set objFrm = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
	    objPage.Highlight
		    If fVerifyObjectExist(objPage.WebElement("html tag:=LI","innertext:=Edit Payment Proposal.*"&strIdentification&".*")) Then
'			    Call fSynUntilObjExists(objPage.WebButton("btnMyInboxAllItems"),MID_WAIT)
'			    Call fSynUntilObjExists(objPage.WebButton("btnMyInboxAllItems"),MID_WAIT)
'			    Call fClick(objPage.WebButton("btnMyInboxAllItems"),"My Inbox AllItems")
			    Call fSynUntilObjExists(objPage.WebElement("html tag:=LI","innertext:=Edit Payment Proposal.*"&strIdentification&".*"),MAX_WAIT)
			    wait 5
			    Call fClick(objPage.WebElement("html tag:=LI","innertext:=Edit Payment Proposal.*"&strIdentification&".*"),"Proposal")
			    Call fSynUntilObjExists(objPage.SAPUIButton("btnOpenTask"),MID_WAIT)
			    Call fClick(objPage.SAPUIButton("btnOpenTask"),"Open Task")
			    Call fSynUntilObjExists(objBroAndPag.SAPButton("btnBack"),MID_WAIT)
			    Call fClick(objBroAndPag.SAPButton("btnBack"),"Back")
			    Call fSynUntilObjExists(objBroAndPag.SAPButton("btnYes"),MID_WAIT)
			    Call fClick(objBroAndPag.SAPButton("btnYes"),"Yes")
		    Else
			     Call fRptWriteReport("Fail","Verify Approve Payment Proposal","My inbox all items are not displayed")
			     Call fRptWriteResultsSummary() 
			     Exit Function
			End If
    On error goto 0	
End Function
''****************************************************************************************************************
'''	Function Name					:				fScheduleProposalAutomaticPaymentRunTransaction
'''	Objective						:				Schedule Proposal Automatic Payment Run Transaction
'''	Input Parameters				:				objDataDict,iRowCountRef
'''	Output Parameters			    :				Nil
'''	Date Created					:				15/May/2020
'''	UFT/QTP Version 				:				15.0
'''	Pre-requisites					:				NIL  
'''	Created By						:				Cigniti
'''	Modification Date		        :		   		
''*****************************************************************************************************************

Public Function fScheduleProposalAutomaticPaymentRunTransaction(objDataDict,iRowCountRef)
    On error resume next
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    'Declarations
    Dim objGLFrame
    Dim blnStartImmediately
    Dim strMsg
    Dim objFrame
        'Read data from excel
        blnStartImmediately = objDataDict.Item("StartImmediately" & iRowCountRef)
        strMsg = objDataDict.Item("PaymentRunProMsg" & iRowCountRef)
        Call fSynUntilObjExists(Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions"),MIN_WAIT)
        Set objGLFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions")
        Set objFrame = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
        'Click Parameter and  Status - to refresh page
        Call fClick(objGLFrame.WebElement("weParameter"),"Payments Tab")
        Call fSynUntilObjExists(objGLFrame.SAPButton("btnPaymentRun"),MIN_WAIT)
        Call fClick(objGLFrame.SAPButton("btnPaymentRun"),"Payment Run")
        'Wait till Popup exist and Start Immediately check box enabled 
        Call fSynUntilObjExists(objGLFrame.SAPCheckBox("chkStartImmediately"),MAX_WAIT)
        'Check Start Immediately check box 
        Call fSelect(objGLFrame.SAPCheckBox("chkStartImmediately"),blnStartImmediately,"Start Immediately")
        'Click Proposal 
        Call fClick(objGLFrame.SAPButton("btnSchedule"),"Schedule")
        Call fSynUntilObjExists(objGLFrame.WebElement("weParameter"),MIN_WAIT)
        Call fClick(objGLFrame.WebElement("weParameter"),"Payments Tab")
        'Verify Message details in Status section
        Call fVerifyStatusInAutomaticPaymentTransactionStatusTable(strMsg)
'        Call fClick(objFrame.SAPButton("btnPaymentAutoPayTransa"),"Payments Log")
'        Call fSynUntilObjExists(objFrame.WebElement("weClearingDocument"),MID_WAIT)
'            If fVerifyObjectExist(objFrame.WebElement("weClearingDocument")) Then
'                intDocNumber = split(fGetRoProperty(objFrame.WebElement("weClearingDocument"),"innertext","Clear Document")," ")(2)
'                If intDocNumber<> "" and Isnumeric(intDocNumber) and len(intDocNumber) >6 Then
'                    'Call fRptWriteReport("PASS", "Verification of Clearing document #","Clearing document "&intDocNumber&" has been created")
'                    Call fRptWriteReport("PASSWITHSCREENSHOT","Verification of Clearing document #","Clearing document "&intDocNumber&" has been created")
'                    Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,intDocNumber,"TestData","ClearingDocumentNo")
'                Else
'                    Call fRptWriteReport("Fail","Verification of Clearing document #","Clearing document has not been created or Not displayed in Payments screen")
'                End If
'        Else
'                Call fRptWriteReport("Fail","Verification of Clearing document #","Clearing document has not been created")
'                Call fRptWriteResultsSummary() 
'                Exit Function
'            End If            
        'Clear Object
        Set objGLFrame = Nothing
    On error goto 0	    
End Function



'****************************************************************************************************************
''    Function Name                    :                fFioriScheduleAutomaticPayableWithInLimit
''    Objective                        :                Complete Schedule Automatic Payment Proposal
''    Input Parameters                 :                Nil
''    Output Parameters                :                Nil
''    Date Created                     :                15/May/2020
''    UFT/QTP Version                  :                15.0
''    Pre-requisites                   :                NIL  
''    Created By                       :                Cigniti
''    Modification Date                :                   
'*************************************************************************************'*************************************************************************************
Public Function fFioriScheduleAutomaticPayableWithInLimit(objDataDict,iRowCountRef)
    On error resume next
    'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
    Dim strIdentification
        'Page Navigation 
        strSearchPageName = objDataDict.Item("WorkSpace" & iRowCountRef)
        Call fFioriFetchAppFromHomePage(strSearchPageName)
        'Function call
        Call fFioriVIMWorkplaceForProcessPOProcessing(objDataDict,iRowCountRef)
        'Logout
        Call fFioriLogOut()
        Call fCloseAllOpenBrowsers("CHROME") 
        'Generate Random number for identification
        strIdentification = "Y"&right(fRemoveSpecialCharectersFromString(" |/|:|[A-Z]",now,""),4)
        Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strIdentification,"TestData","Identification")
        ' lOGIN AS APAnalyst
        Call fFioriLogin("APAnalyst")
        'Page Navigation 
        strSearchPageName = objDataDict.Item("FB03PageName" & iRowCountRef)
        Call fFioriFetchAppFromHomePage (strSearchPageName)
        'Get Document Numbers
        arrDocuNumbers = fFioriFillDocumentList(objDataDict,iRowCountRef)
        ''Page Navigation 
        strSearchPageName = objDataDict.Item("SearchPageName" & iRowCountRef)
        Call fFioriFetchAppFromHomePage (strSearchPageName)
        'Fill Status details 
        Call fFillAutomaticPaymentTransactionsStatus (objDataDict,iRowCountRef)
        'Fill Parameter details 
        Call fFillAutomaticPaymentTransactionsParameters(objDataDict,iRowCountRef)
        'Fill Free Selection details 
        Call fFillAutomaticPaymentTransactionsFreeSelection(objDataDict,iRowCountRef)
        'Fill Additional log details 
        Call fFillAutomaticPaymentTransactionsAdditionalLog(objDataDict,iRowCountRef)
        'Save detail and verify Conformation message 
        Call fSaveAutomaticPaymentTransactionDetails(objDataDict,iRowCountRef)
        'Payment Schedule Proposal
        Call fScheduleProposalAutomaticPaymentTransaction(objDataDict,iRowCountRef)
        Call fFioriLogOut()
        ' lOGIN AS AP Manager
        Call fFioriLogin("APManager")
        'Page Navigation
        strSearchPageName = objDataDict.Item("SearchPageName_Mngr" & iRowCountRef)
        Call fFioriFetchAppFromHomePage (strSearchPageName)
        'Payment Proposal workflow
        Call fFillPaymentProposalWorkflowSelectionValues(objDataDict,iRowCountRef)
        Call fFioriLogOut()
        ' lOGIN AS Manager
        Call fFioriLogin("APManager")
        strSearchPageName = objDataDict.Item("MyInbox" & iRowCountRef)
        Call fFioriFetchAppFromHomePage (strSearchPageName)
        Call fFioriApprovePaymentProposal(objDataDict,iRowCountRef)
        Call fFioriLogOut()
        Call fFioriLogin("APAnalyst")
        'Page Navigation
        strSearchPageName = objDataDict.Item("SearchPageName" & iRowCountRef)
        Call fFioriFetchAppFromHomePage (strSearchPageName)
        Call fFillAutomaticPaymentTransactionsStatus(objDataDict,iRowCountRef)
        'Automatic Payment Run
        Call fScheduleProposalAutomaticPaymentRunTransaction(objDataDict,iRowCountRef)
    On error goto 0	
End Function

'****************************************************************************************************************
''	Function Name					:				fFioriScheduleAutomaticPayable
''	Objective						:				Scedule Automatic Payable
''	Input Parameters				:				objDataDict,iRowCountRef
''	Output Parameters			    :				Nil
''	Date Created					:				15/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************
Public Function fFioriScheduleAutomaticPayable(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim strIdentification
		strSearchPageName = objDataDict.Item("WorkSpace" & iRowCountRef)
		strManagerApproval = objDataDict.Item("ManagerApproval" & iRowCountRef)
		Call fFioriFetchAppFromHomePage(strSearchPageName)
		Call fFioriVIMWorkplaceForProcessPOProcessing(objDataDict,iRowCountRef)
		Call fFioriLogOut()
		Call fCloseAllOpenBrowsers("CHROME") 
    	strIdentification = "Y"&right(fRemoveSpecialCharectersFromString(" |/|:|[A-Z]",now,""),4)
		Call fWriteOutputValueInExcel(Environment("TestName"),1,strIdentification,"TestData","Identification")
		Call fFioriLogin("APAnalyst")
		strSearchPageName = objDataDict.Item("FB03PageName" & iRowCountRef)
		Call fFioriFetchAppFromHomePage (strSearchPageName)
		arrDocuNumbers = fFioriFillDocumentList(objDataDict,iRowCountRef)
		strSearchPageName = objDataDict.Item("SearchPageName" & iRowCountRef)
		Call fFioriFetchAppFromHomePage (strSearchPageName)
		Call fFillAutomaticPaymentTransactionsStatus (objDataDict,iRowCountRef)
		Call fFillAutomaticPaymentTransactionsParameters(objDataDict,iRowCountRef)
		Call fFillAutomaticPaymentTransactionsFreeSelection(objDataDict,iRowCountRef)
		Call fFillAutomaticPaymentTransactionsAdditionalLog(objDataDict,iRowCountRef)
		Call fSaveAutomaticPaymentTransactionDetails(objDataDict,iRowCountRef)
		Call fScheduleProposalAutomaticPaymentTransaction(objDataDict,iRowCountRef)
		Call fFioriLogOut()
		Call fFioriLogin("APManager")
		strSearchPageName = objDataDict.Item("SearchPageName_Mngr" & iRowCountRef)
		Call fFioriFetchAppFromHomePage (strSearchPageName)
		Call fFillPaymentProposalWorkflowSelectionValues(objDataDict,iRowCountRef)
		Wait MIN_WAIT
		strSearchPageName = objDataDict.Item("MyInbox" & iRowCountRef)
		Call fFioriFetchAppFromHomePage (strSearchPageName)
		Call fFioriApprovePaymentProposal(objDataDict,iRowCountRef)
		Call fFioriLogOut()
		Call fFioriLogin("APAnalyst")
		strSearchPageName = objDataDict.Item("SearchPageName" & iRowCountRef)
		Call fFioriFetchAppFromHomePage (strSearchPageName)
		Call fFillAutomaticPaymentTransactionsStatus(objDataDict,iRowCountRef)
    	Call fScheduleProposalAutomaticPaymentRunTransaction(objDataDict,iRowCountRef)
	On error goto 0
End Function

'****************************************************************************************************************
''	Function Name					:				fFioriScheduleAutomaticPayable
''	Objective						:				Scedule Automatic Payable
''	Input Parameters				:				objDataDict,iRowCountRef
''	Output Parameters			    :				Nil
''	Date Created					:				15/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************
Public Function fFioriScheduleAutomaticPayableTwo(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim strIdentification
		strSearchPageName = objDataDict.Item("WorkSpace" & iRowCountRef)
		strManagerApproval = objDataDict.Item("ManagerApproval" & iRowCountRef)
		Call fFioriFetchAppFromHomePage(strSearchPageName)
		Call fFioriVIMWorkplaceForProcessPOProcessingTwo(objDataDict,iRowCountRef)
		Call fFioriLogOut()
		Call fCloseAllOpenBrowsers("CHROME") 
			If Ucase(strManagerApproval) = "YES" Then				
				Call fFioriLogin("APManager")
				strSearchPageName = objDataDict.Item("WorkSpace" & iRowCountRef)
				Call fFioriFetchAppFromHomePage(strSearchPageName)
				Call fFioriManagerApproval()
				Call fFioriLogOut()
			End IF		
		strIdentification = "Y"&right(fRemoveSpecialCharectersFromString(" |/|:|[A-Z]",now,""),4)
		Call fWriteOutputValueInExcel(Environment("TestName"),1,strIdentification,"TestData","Identification")
		Call fFioriLogin("APAnalyst")
		strSearchPageName = objDataDict.Item("FB03PageName" & iRowCountRef)
		Call fFioriFetchAppFromHomePage (strSearchPageName)
		arrDocuNumbers = fFioriFillDocumentList(objDataDict,iRowCountRef)
		strSearchPageName = objDataDict.Item("SearchPageName" & iRowCountRef)
		Call fFioriFetchAppFromHomePage (strSearchPageName)
		Call fFillAutomaticPaymentTransactionsStatus (objDataDict,iRowCountRef)
		Call fFillAutomaticPaymentTransactionsParameters(objDataDict,iRowCountRef)
		Call fFillAutomaticPaymentTransactionsFreeSelection(objDataDict,iRowCountRef)
		Call fFillAutomaticPaymentTransactionsAdditionalLog(objDataDict,iRowCountRef)
		Call fSaveAutomaticPaymentTransactionDetails(objDataDict,iRowCountRef)
		Call fScheduleProposalAutomaticPaymentTransaction(objDataDict,iRowCountRef)
		Call fFioriLogOut()
		Call fFioriLogin("APManager")
		strSearchPageName = objDataDict.Item("SearchPageName_Mngr" & iRowCountRef)
		Call fFioriFetchAppFromHomePage (strSearchPageName)
		Call fFillPaymentProposalWorkflowSelectionValues(objDataDict,iRowCountRef)
		Call fFioriLogOut()
		Call fFioriLogin("APManager")
		strSearchPageName = objDataDict.Item("MyInbox" & iRowCountRef)
		Call fFioriFetchAppFromHomePage (strSearchPageName)
		Call fFioriApprovePaymentProposal(objDataDict,iRowCountRef)
		Call fFioriLogOut()
		Call fFioriLogin("APAnalyst")
		strSearchPageName = objDataDict.Item("SearchPageName" & iRowCountRef)
		Call fFioriFetchAppFromHomePage (strSearchPageName)
		Call fFillAutomaticPaymentTransactionsStatus(objDataDict,iRowCountRef)
		Call fScheduleProposalAutomaticPaymentRunTransactionTwo(objDataDict,iRowCountRef)
		Call fFioriLogOut()
    On error goto 0
End Function

'****************************************************************************************************************
''	Function Name					:				fRemoveSpecialCharectersFromString
''	Objective						:				Remove Special Charecters From String
''	Input Parameters				:				strPattern,strString,strReplaceChar
''	Output Parameters			    :				Nil
''	Date Created					:				15/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************
Public Function fRemoveSpecialCharectersFromString(strPattern,strString,strReplaceChar)
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	Dim objRe
		Set objRe = New RegExp
		objRe.Global = True
		objRe.Pattern = strPattern
		fRemoveSpecialCharectersFromString = objRe.Replace(strString,strReplaceChar)
End Function
'****************************************************************************************************************
''	Function Name					:				fFioriVIMWorkplaceForProcessPOProcessing
''	Objective						:				Remove Special Charecters From String
''	Input Parameters				:				strPattern,strString,strReplaceChar
''	Output Parameters			    :				Nil
''	Date Created					:				15/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************
Public Function fFioriVIMWorkplaceForProcessPOProcessingTwo(objDataDict,iRowCountRef)
	On error resume next
	'Verify if Step Failed, If yes, it will not run the function
    If Environment("StepFailed") = "YES" Then
		Exit Function
	End If	
	'Variable declaration
	Dim strSwitchWorkView
	Dim strComment
	Dim intBalance
	Dim strTaxCode
	Dim blnExecute
	strSwitchWorkView = objDataDict.Item("SwitchWorkView" & iRowCountRef)
	strComment = objDataDict.Item("BypassComment" & iRowCountRef)
	intBalance = objDataDict.Item("Balance" & iRowCountRef)
	strTaxCode = objDataDict.Item("TaxCode" & iRowCountRef)
	'Frame object declaration
	Set objBrwAndPg = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
		If fVerifyObjectExist(objBrwAndPg.SAPButton("btnSwitchWorkView")) Then
			'Click on Switch work view
			Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnSwitchWorkView"),MID_WAIT)
			Call fClick(objBrwAndPg.SAPButton("btnSwitchWorkView"),"Switch Work View")
			'select Switch work view
			Call fSynUntilObjExists(objBrwAndPg.SAPRadioGroup("rbSwitchWorkView"),MID_WAIT)
			Call fSelect(objBrwAndPg.SAPRadioGroup("rbSwitchWorkView"),strSwitchWorkView,"Switch Work View")
			Call fClick(objBrwAndPg.SAPButton("btnContinue"),"Continue Switch Work View")
			'Exeute invoice
			blnExecute =  fFioriExecteInvoice(objBrwAndPg)
			Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnSimulateRules"),MID_WAIT)
				If fVerifyObjectExist(objBrwAndPg.SAPButton("btnSimulateRules")) AND blnExecute Then
					Call fClick(objBrwAndPg.SAPButton("btnSimulateRules"),"Simulate Rules")
					Call fFioriExceptionValidationForSimulateBusinessRules(objBrwAndPg,strComment)
					Call fSynUntilObjExists(objBrwAndPg.WebElement("weLineItems"),MID_WAIT)
					Call fClick(objBrwAndPg.WebElement("weLineItems"),"Line Items")
					Call fSynUntilObjExists(objBrwAndPg.WebTable("tblLineItems"),MID_WAIT)
					intPODoc = fFioriPurchasingDocID()
					intLineItemsRows = objBrwAndPg.WebTable("tblLineItems").RowCount()
						For i = 1 To intLineItemsRows
							intPurchasingDoc = objBrwAndPg.WebTable("tblLineItems").ChildItem(i,5,"SAPEdit",0).getroproperty("value")
								If intPurchasingDoc = intPODoc Then
									objBrwAndPg.WebTable("tblLineItems").ChildItem(1,10,"WebList",0).highlight
									Wait(2)
									objBrwAndPg.WebTable("tblLineItems").ChildItem(1,14,"WebList",0).click
									wait(2)
									objBrwAndPg.WebTable("tblLineItems").ChildItem(1,14,"SAPList",0).Select strTaxCode
									Wait(2)
									Exit For
								End If
						Next
					Call fClick(objBrwAndPg.SAPButton("btnSave"),"Save ")
					Call fSynUntilObjExists(objBrwAndPg.SAPEdit("txtDocumentBalance"),MID_WAIT)
						If intBalance = fGetRoProperty(objBrwAndPg.SAPEdit("txtDocumentBalance"),"value","Document Balance")Then
							Call fRptWriteReport("PASSWITHSCREENSHOT", "Document Balance shown as Zero","Successfully, the Document Balance shown as Zero")
						Else
							Call fRptWriteReport("Fail", "Document Balance not turned to Zero","Still, the Document Balance is not turned to Zero")
						End  If
					Call fClick(objBrwAndPg.SAPButton("btnApplyRules"),"Apply Rules")
					Call fSynUntilObjExists(objBrwAndPg.SAPButton("btnReset"),MID_WAIT)
					Call fClick(objBrwAndPg.SAPButton("btnReset"),"Reset")
					Call fFioriApproveInvoice(objBrwAndPg)
				Else
					Call fRptWriteReport("Fail", "Unable to launch the Execute Workitem frame","Unsuccessfully, the Execute Workitem frame not displayed")
					Call fRptWriteResultsSummary() 
					Exit Function
				End  If	
		Else
			Call fRptWriteReport("Fail", "Unable to launch the switch work view","Unsuccessfully, the Execute switch work view not displayed")
			Call fRptWriteResultsSummary() 
			Exit Function
		End If
	On error goto 0
End Function
'****************************************************************************************************************
''	Function Name					:				fScheduleProposalAutomaticPaymentRunTransactionTwo
''	Objective						:				
''	Input Parameters				:				objDataDict,iRowCountRef
''	Output Parameters			    :				Nil
''	Date Created					:				15/May/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				NIL  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************
Public Function fScheduleProposalAutomaticPaymentRunTransactionTwo(objDataDict,iRowCountRef)
		On error resume next
		'Verify if Step Failed, If yes, it will not run the function
	    If Environment("StepFailed") = "YES" Then
			Exit Function
		End If	
		'Declarations
		Dim objGLFrame
		Dim blnStartImmediately
		Dim strMsg
		Dim objFrame
		'Read data from excel
		blnStartImmediately = objDataDict.Item("StartImmediately" & iRowCountRef)
		strMsg = objDataDict.Item("PaymentRunProMsg" & iRowCountRef)
		Set objGLFrame = Browser("brFioriAutoDesk").Page("pgGLAccountLineItem").SAPFrame("frmAutomaticPaymentTransactions")
		Set objFrame = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
		Call fSynUntilObjExists(objGLFrame.WebElement("weParameter"),MID_WAIT)
		'Click Parameter
		Call fClick(objGLFrame.WebElement("weParameter"),"Parameter")
		Call fClick(objGLFrame.SAPButton("btnPaymentRun"),"Payment Run")
		'Wait till Popup exist and Start Immediately check box enabled 
		Call fSynUntilObjExists(objGLFrame.SAPCheckBox("chkStartImmediately"),MAX_WAIT)
		'Check Start Immediately check box 
		Call fSelect(objGLFrame.SAPCheckBox("chkStartImmediately"),blnStartImmediately,"Start Immediately")
		'Click Proposal 
		Call fClick(objGLFrame.SAPButton("btnSchedule"),"Proposal Tab")
		'Verify Message details in Status section
		Call fVerifyStatusInAutomaticPaymentTransactionStatusTable(strMsg)	
		objGLFrame.SAPButton("btnPayment").Highlight
		objGLFrame.SAPButton("btnPayment").Click
'		Call fClick(objGLFrame.SAPButton("btnPayment"),"Payment")
		Call fSynUntilObjExists(objFrame.WebElement("weClearingDocument"),MID_WAIT)
			If fVerifyObjectExist(objFrame.WebElement("weClearingDocument")) Then
				intDocNumber = split(fGetRoProperty(objFrame.WebElement("weClearingDocument"),"innertext","Clear Document")," ")(2)
				Call fRptWriteReport("PASS", "Verification of Clearing document","Clearing document "&intDocNumber&" has been created")
			Else
				Call fRptWriteReport("Fail","Verification of Clearing document","Clearing document has not been created")
			End If
		'Clear Object
		Set objGLFrame = Nothing
	On error goto 0
End Function



'****************************************************************************************************************
''	Function Name					:				fShellScriptTabOut
''	Objective						:				Tab out from field
''	Input Parameters				:				Nil
''	Output Parameters			    :				Nil
''	Date Created					:				06/June/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				Nil  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************
 Function fShellScriptTabOut()
 	 ' W Shell script for page down
        Set objWshShell = CreateObject("WScript.shell")
        objWshShell.SendKeys "{TAB}"
        wait (2)
 End Function

'****************************************************************************************************************
''	Function Name					:				fShellScriptEnter
''	Objective						:				Enter From Key board
''	Input Parameters				:				Nil
''	Output Parameters			    :				Nil
''	Date Created					:				06/June/2020
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				Nil  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************
 Function fShellScriptEnter()
 	 ' W Shell script for page down
        Set objWshShell = CreateObject("WScript.shell")
        objWshShell.SendKeys "{ENTER}"
        wait (2)
 End Function


'****************************************************************************************************************
''	Function Name					:				fUpdateBaseDataTaxCode
''	Objective						:				Update the Basic Data Amount
''	Input Parameters				:				
''	Output Parameters			    :				
''	Date Created					:				
''	UFT/QTP Version 				:				15.0
''	Pre-requisites					:				Nil  
''	Created By						:				Cigniti
''	Modification Date		        :		   		
'*****************************************************************************************************************
'fUpdateBasicDataTaxCode(objDataDict,iRowCountRef)
Public Function fUpdateBasicDataTaxCode(objDataDict,iRowCountRef)
	fUpdateBasicDataTaxCode = False	
	'Get the values from testdata sheet
	intTaxAmountdata =  objDataDict.Item("Tax Amount(Non US/Canada)" & iRowCountRef)

	set objBrwAndPg =  Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
	
	Call fClick(objBrwAndPg.WebElement("weBasicData"),"Basic Data")
	
	If intTaxAmountdata = "" Then
        If fGetRoProperty(objBrwAndPg.SAPCheckBox("chkAutoCalculateTax"),"checked","Auto Calculate Tax") = 0 Then
            Call fSelect(objBrwAndPg.SAPCheckBox("chkAutoCalculateTax"),"ON","Auto Calculate Tax")
            'Hit Enter
             Call fClick(objBrwAndPg.SAPButton("btnEnter"),"Enter")
        End If
    End If
	
	intBalance = fGetRoProperty(objBrwAndPg.SAPEdit("txtBalance") ,"value","Basic Data Balance")
	
	inttxtGrossAmount = fGetRoProperty(objBrwAndPg.SAPEdit("txtGrossAmount") ,"value","Basic Data Gross Amount")
	
	If instr(intBalance,"-") Then
		intBalance = Replace(intBalance,"-","")
		intGrossTotal = cdbl(inttxtGrossAmount) +  cdbl(intBalance)
	Else
	
		intGrossTotal = cdbl(inttxtGrossAmount) -  cdbl(intBalance)
	End If

	Call fEnterText(objBrwAndPg.SAPEdit("txtGrossAmount"),intGrossTotal,"Update Gross Amount")
	
	Call fClick(objBrwAndPg.SAPButton("btnSave"),"Save")
	Wait(6)
	
'	If "0.00" <> fGetRoProperty(objBrwAndPg.SAPEdit("txtBalance") ,"value","Basic Data Balance")Then
'		intBalance = fGetRoProperty(objBrwAndPg.SAPEdit("txtBalance") ,"value","Basic Data Balance")
'	
'		inttxtGrossAmount = fGetRoProperty(objBrwAndPg.SAPEdit("txtGrossAmount") ,"value","Basic Data Gross Amount")
'		
'			If instr(intBalance,"-") Then
'				intBalance = Replace(intBalance,"-","")
'				intGrossTotal = cdbl(inttxtGrossAmount) -  cdbl(intBalance)
'			Else
'			
'				intGrossTotal = cdbl(inttxtGrossAmount) +  cdbl(intBalance)
'			End If
'			
'			Call fEnterText(objBrwAndPg.SAPEdit("txtGrossAmount"),intGrossTotal,"Update Gross Amount")
'	
'			Call fClick(objBrwAndPg.SAPButton("btnSave"),"Save")
'			Wait(4)
'		
'	End If
	'Get the Balance
	If Ucase("Balance is ok") = Trim(Ucase(fGetRoProperty(Browser("BrProcessing").Page("PgProcessing").SAPFrame("frmProcessInvOrCM").WebElement("weCreditMemoBalanceIcon")),"title","Balance Title")) Then
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Basic Data Balance shown as OK","Balance is OK icon displayed")
	Else
		If "0.00" = fGetRoProperty(objBrwAndPg.SAPEdit("txtBalance") ,"value","Basic Data Balance")Then
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Basic Data Balance shown as Zero","Successfully, the Balance is shown as Zero")
			fUpdateBasicDataTaxCode = True
		Else
			Call fRptWriteReport("Fail", "Basic Data Balance not turned to Zero","Still the Document Balance is not turned to Zero")
		End  If
End  if	
End Function
'******************************************************************************************************************************************************************************
'	Function Name						:		fFetchCancelledPODetailsInFioriPurchaseOrderPage
'	Objective							:		Used to Fetch the Cancelled PO Details in Ariba
'	Input Parameters					:		
'	Output Parameters					:		strActReqID
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Function fFetchCancelledPODetailsInFioriPurchaseOrderPage()
	
	Set obFioriPOPage = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk")
	strPurchaseOrder = fGetSingleValue("AutoPONumber","TestData",Environment("TestName"))
	Call fSynUntilObjExists(obFioriPOPage.WebElement("weStandardPONotification"),MIN_WAIT)
	
'	Verify Fiori Standard PO Page is Navigated
	If obFioriPOPage.WebElement("weStandardPONotification").Exist(1) Then
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Navigated to Standard PO Page in Fiori","Navigated to Standard PO Page in Fiori Successfully")
'		Click on Other Purchase Order Button
		fClick obFioriPOPage.SAPButton("btnOtherPurchaseOrder"),"Other Purchase Order"		
		Call fSynUntilObjExists(obFioriPOPage.SAPEdit("txtPurchaseOrder"),MIN_WAIT)
		
'		Verify Select Document POP UP is Displayed
		If obFioriPOPage.SAPEdit("txtPurchaseOrder").Exist(1) Then
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Select Document POP UP is Displayed","Displayed Select Document POP UP Successfully")
			obFioriPOPage.SAPEdit("txtPurchaseOrder").Set strPurchaseOrder
			fClick obFioriPOPage.SAPButton("btnOtherPurchaseOrder"),"Other Purchase Order"		
		 	fClick obFioriPOPage.SAPButton("btnOtherDocument"),"Other Document"		
		  
'		  Verify PO Details are displayed
			Call fSynUntilObjExists(obFioriPOPage.WebElement("weStandardPONotification"),MID_WAIT)
			If obFioriPOPage.WebElement("weStandardPONotification").Exist(1) Then
				Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify Displayed Standard PO Details Successfully","Displayed Standard PO Details Successfully")
			Else
				Call fRptWriteReport("Fail", "Verify Displayed Standard PO Details Successfully","Failed to Display Standard PO Details")
				Call fRptWriteResultsSummary() 
        		'ExitAction
			End If
			
		Else
			Call fRptWriteReport("Fail", "Verify Select Document POP UP is Displayed","Failed to Display Select Document POP UP")
			Call fRptWriteResultsSummary() 
    		'ExitAction
		End If
		
	Else
		Call fRptWriteReport("Fail", "Verify Navigated to Standard PO Page in Fiori","Failed to Navigate to Standard PO Page in Fiori")
		Call fRptWriteResultsSummary()        
        'ExitAction
	End If	
	
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fVerifyCancelledPOStatusInFioriPurchaseOrderPage
'	Objective							:		Used to Verify the Cancelled PO Status in Ariba
'	Input Parameters					:		
'	Output Parameters					:		strActReqID
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Function fVerifyCancelledPOStatusInFioriPurchaseOrderPage()
	
	Set obFioriPODetailsTable = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").WebTable("wbtPOStatusDetails")
	intRowCount = obFioriPODetailsTable.RowCount
	intDelItemCount = 0
	IntItemCnt = 0
	For intRow = 1 To intRowCount
		If obFioriPODetailsTable.GetCellData(intRow,3) <> "" Then
			Set objSapButton = obFioriPODetailsTable.ChildItem(1,2,"SAPButton",0)
			strItemStatus = objSapButton.getroproperty("name")
			IntItemCnt = intRow
			If Ucase(Trim(strItemStatus)) = "DELETED" Then
				intDelItemCount = intDelItemCount+1
			End If			
		End If		
	Next
	
	If IntItemCnt > 0 and IntItemCnt = intDelItemCount Then
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify the Staus of the PO Should be displayed as Deleted","The Staus of the PO is displayed as Deleted")
	Else		
		Call fRptWriteReport("Fail", "Verify the Staus of the PO Should be displayed as Deleted","Failed to Display the Staus of the PO is as Deleted")
		Call fRptWriteResultsSummary() 
        'ExitAction
	End If	
	
End Function

'******************************************************************************************************************************************************************************
'	Function Name						:		fFioriNavigateDisplayPO
'	Objective							:		Used to Navigate to DisplayPO page in Fiori
'	Input Parameters					:		
'	Output Parameters					:		strActReqID
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Function fFioriNavigateDisplayPO()
	
	strSearchPageName = objDataDict.Item("WorkSpace" & iRowCountRef)
	Set objFioriPOPage = Browser("brFiori").Page("pgFiori")
	Call fSynUntilObjExists(objFioriPOPage.WebButton("btDisplayPurchaseOrder"),MID_WAIT)
	
'	Verify DisplayPObutton is displayed
	If objFioriPOPage.WebButton("btDisplayPurchaseOrder").Exist(1) Then
		Call fRptWriteReport("PASSWITHSCREENSHOT", "Verify "&strSearchPageName&" is Searched in Fiori Home Page Successfully",strSearchPageName&" is Searched in Fiori Home Page Successfully")
		
'		Click on DisplayPO Button
		fClick objFioriPOPage.WebButton("btDisplayPurchaseOrder"),strSearchPageName
	Else		
		Call fRptWriteReport("Fail", "Verify "&strSearchPageName&" is Searched in Fiori Home Page Successfully",strSearchPageName&" is Failed to Search in Fiori Home Page")
		Call fRptWriteResultsSummary() 
        'ExitAction
	End If	
	
End Function


'******************************************************************************************************************************************************************************
'	Function Name						:		fFioriUpdateGSTRegistrationNumber
'	Objective							:		Used to Update the GST Registration Number 
'	Input Parameters					:		
'	Output Parameters					:		
'	Date Created						:		
'	UFT Version							:		UFT 15.0	
'	Pre-requisites						:		NIL  
'	Created By							:		Cigniti 						
'	Modification Date					:		   
'******************************************************************************************************************************************************************************

Public Function fFioriUpdateGSTRegistrationNumber(objDataDict,iRowCountRef)
	
If Environment("StepFailed") = "YES" Then
	Exit Function
End If	

intGSTNumber = objDataDict.Item("GST Registration Number" & iRowCountRef)

Dim intSimulateRules
Dim intIter
fFioriUpdateGSTRegistrationNumber = False
set sObjectName =  Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
Call fSynUntilObjExists(sObjectName.WebTable("tblException"),MID_WAIT)
    If fVerifyObjectExist(sObjectName.WebTable("tblException")) Then
            Call fSynUntilObjExists(sObjectName.WebTable("tblException"),MID_WAIT)
            intSimulateRules = sObjectName.WebTable("tblException").RowCount()
                For intIter = 2 To intSimulateRules
                      If "Exception occured" = sObjectName.WebTable("tblException").ChildItem(intIter,3,"WebElement",0).GetRoProperty("innertext") Then    
                        If "Missing/Invalid Vendor GST Registration number (PO" = Trim(sObjectName.WebTable("tblException").ChildItem(intIter,3,"WebElement",0).GetRoProperty("innertext"))  Then 
                            Call fClick(sObjectName.SAPButton("btnExit"),"Exit")                            
                        End If
     				End  IF	                                   
                Next
            Call fClick(sObjectName.SAPButton("btnExit"),"Exit")
            fFioriUpdateGSTRegistrationNumber = TRUE
        Else        
            Call fRptWriteReport("Fail", "Verify Simulpate business rules","Simulate rules page not displayed")
            Call fRptWriteResultsSummary() 
            Exit Function
        End If
	
		
		'Select the Basic Data Tab
		Call fSelect(sObjectName.WebTabStrip("wtsMyTask"),"Basic Data","Basic Data")
		'Enter GST Number
		Call fEnterText(sObjectName.SAPEdit("txtGSTRegNo"),intGSTNumber,"GST Registration Number")
		'Verify the Auto Calacuate Tax checkbox
		If fGetRoProperty(sObjectName.SAPCheckBox("chkAutoCalculateTax"),"checked","Auto Calculate Tax") = 0 Then
			Call fSelect(sObjectName.SAPCheckBox("chkAutoCalculateTax"),"ON","Auto Calculate Tax")
			'Hit Enter
			 Call fClick(sObjectName.SAPButton("btnEnter"),"Enter")
		End If
		
		'Get the Tax Amount
		intTaxAmount = objBrwAndPg.SAPEdit("txtTaxAmount").GetROProperty("value")
		
		'Get the Gross Amount
		intGrossAmount = objBrwAndPg.SAPEdit("txtGrossAmount").GetROProperty("value")
		
		If intTaxAmount > 0 Then
			intTotalGross = int(intGrossAmount) + int(intTaxAmount)		
		Else
			intTotalGross = intGrossAmount
		End If
		
		'Set the value in Gross Amount
		Call fEnterText(objBrwAndPg.SAPEdit("txtGrossAmount"),intTotalGross,"Update Gross Amount")
		
		 Call fClick(sObjectName.SAPButton("btnSave"),"Save")
		
		'Get the Balance
		If intBalance = fGetRoProperty(objBrwAndPg.SAPEdit("txtBalance") ,"value","Basic Data Balance")Then
			Call fRptWriteReport("PASSWITHSCREENSHOT", "Basic Data Balance shown as Zero","Successfully, the Balance is shown as Zero")
			fUpdateBasicDataTaxCode = True
		Else
			Call fRptWriteReport("Fail", "Basic Data Balance not turned to Zero","Still, the Document Balance is not turned to Zero")
		End  If
		
		intSimulateRules = sObjectName.WebTable("tblException").RowCount()
        For intIter = 2 To intSimulateRules
          If "Exception occured" = sObjectName.WebTable("tblException").ChildItem(intIter,3,"WebElement",0).GetRoProperty("innertext") Then    
            If "Check Withholding Tax Data (PO)" = Trim(sObjectName.WebTable("tblException").ChildItem(intIter,3,"WebElement",0).GetRoProperty("innertext"))  Then 
                Call fClick(sObjectName.SAPButton("btnExit"),"Exit")                            
            End If
		  End  IF	                                   
        Next        
        
        'Select the Basic Data Tab
		Call fSelect(sObjectName.WebTabStrip("wtsMyTask"),"Tax","Tax")
        'Select With Holding checkbox
        Call fSelect(sObjectName.SAPCheckBox("chkWithholdTax"),"ON","With Holding Tax")
        
		'Click on Save button
		 Call fClick(sObjectName.SAPButton("btnSave"),"Save")
		 
		intSimulateRules = sObjectName.WebTable("tblException").RowCount()
        For intIter = 2 To intSimulateRules         
            If "Check Withholding Tax Data (PO)" = Trim(sObjectName.WebTable("tblException").ChildItem(intIter,3,"WebElement",0).GetRoProperty("innertext"))  Then 
             If "Processed" = sObjectName.WebTable("tblException").ChildItem(intIter,3,"WebElement",0).GetRoProperty("innertext") Then    
                Call fClick(sObjectName.SAPButton("btnExit"),"Exit")                            
            End If
		  End  IF	                                   
        Next 		 
End Function

' Get Document status
public Function fFioriGetInvoiceORCreditMemoDocumentStatus()
		On error Resume Next
		
		Dim objBrowAndPageAndFrm
		Dim intColumnNo
		Dim intRowCount
		Dim strDocumentStatus
		
		Set objBrowAndPageAndFrm = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")
		'Get Column number based on Column name
        intColumnNo = fGetTableHeaderColumnNumber(objBrowAndPageAndFrm.WebTable("tblDocumentsListHeader"),1,1,"Document Status")
        'Get Row count based on Column number
        intRowCount = fGetRoProperty(objBrowAndPageAndFrm.WebTable("tblDocumentsDetailsList"),"rows","Document List Data")
        'Get Exception Reason
        strDocumentStatus = fGetCelldata(objBrowAndPageAndFrm.WebTable("tblDocumentsDetailsList"),intRowCount,intColumnNo,"Document Status")        
        
        fFioriGetInvoiceORCreditMemoDocumentStatus = strDocumentStatus
   	

End Function 

Public Function fFioriGetTaxHoldingData(objDataDict,iRowCountRef)
    'On error resume next
    Dim intCompanyCode
    Dim objBrowAndPage
    Dim objBrowAndPageAndFrm
    Dim intColumnNo
    Dim arrDocumentNumbers()
        'Get the Invoice Number from testdata sheet
		intInvoice = fGetSingleValue("AutoInvoiceNumber","TestData",Environment("TestName")) 
        intCompanyCode = objDataDict.Item("CompanyCode" & iRowCountRef)
        intVendor = objDataDict.Item("Vendor" & iRowCountRef)
       
        Call fSynUntilObjExists(Browser("brFiori").Page("pgFiori"),MID_WAIT)
        Set objBrowAndPage = Browser("brFiori").Page("pgFiori")
        Set objBrowPage = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk")
        Set objBrowAndPageAndFrm = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")

 		Call fSynUntilObjExists(objBrowAndPageAndFrm.SAPButton("btnDocumentList"),MID_WAIT)
            If fVerifyObjectExist(objBrowAndPageAndFrm.SAPButton("btnDocumentList")) Then
                Call fClick(objBrowAndPageAndFrm.SAPButton("btnDocumentList"),"Document List")
                Call fSynUntilObjExists(objBrowAndPageAndFrm.SAPEdit("txtCompanyCode"),MID_WAIT)
                Call fEnterText(objBrowAndPageAndFrm.SAPEdit("txtCompanyCode"),intCompanyCode,"Company Code")
                Call fEnterText(objBrowAndPageAndFrm.SAPEdit("txtReferenceNumber"),intInvoice,"Reference Number")
                Call fClick(objBrowAndPageAndFrm.SAPButton("btnExecute"),"Execute")
                Call fSynUntilObjExists(objBrowAndPageAndFrm.WebTable("tblDocumentListHeader"),MID_WAIT)          
                Setting.WebPackage("ReplayType") = 2
				objBrowAndPageAndFrm.WebTable("tblDocumentListData").ChildItem(1,3,"WebList",0).FireEvent "ondblclick"
				Call fSynUntilObjExists(objBrowAndPageAndFrm.WebTable("tblDataEntryDataHeader"),MID_WAIT)
                intColuNo = fGetTableHeaderColumnNumber(objBrowAndPageAndFrm.WebTable("tblDataEntryDataHeader"),1,1,"Account")
                intRowCount = fGetRowNumberInTableBasedonColumnData(objBrowAndPageAndFrm.WebTable("tblDataEntryData"),intColuNo,right(intVendor,6))
                objBrowAndPageAndFrm.WebTable("tblDataEntryData").ChildItem(intRowCount,intColuNo,"WebList",0).FireEvent "ondblclick"
                Setting.WebPackage("ReplayType") = 1
                Call fSynUntilObjExists(objBrowAndPageAndFrm.SAPButton("btnMore"),MID_WAIT) 
                Call fClick(objBrowAndPageAndFrm.SAPButton("btnMore"),"More")
                Wait(2)
                Setting.WebPackage("ReplayType") = 2
                Browser("brFioriAutoDesk").InsightObject("lstWithholdingTaxData").Click
                'Call fSelect(objBrowAndPageAndFrm.SAPDropDownMenu("lstAdditionalData"),"Withholding Tax Data (Ctrl+F5)","Withholding Tax Data")
				Setting.WebPackage("ReplayType") = 1
				Call fSynUntilObjExists(objBrowAndPageAndFrm.WebTable("tblTaxInformationHeader"),MID_WAIT) 
					If fVerifyObjectExist(objBrowAndPageAndFrm.WebTable("tblTaxInformationHeader")) Then
					
						intNameOfWTaxTypeColNum = fGetTableHeaderColumnNumber(objBrowAndPageAndFrm.WebTable("tblTaxInformationHeader"),1,1,"Name of WTax Type")
						intWTaxCodeColNum = fGetTableHeaderColumnNumber(objBrowAndPageAndFrm.WebTable("tblTaxInformationHeader"),1,1,"WTax Code")
						intWTaxBaseColNum = fGetTableHeaderColumnNumber(objBrowAndPageAndFrm.WebTable("tblTaxInformationHeader"),1,1,"W/Tax Base")
						intWTaxAmtColNum = fGetTableHeaderColumnNumber(objBrowAndPageAndFrm.WebTable("tblTaxInformationHeader"),1,1,"W/Tax Amt")
						intWTaxBaseLCColNum = fGetTableHeaderColumnNumber(objBrowAndPageAndFrm.WebTable("tblTaxInformationHeader"),1,1,"W/Tax Base LC")
						intWTaxAmntLCColNum = fGetTableHeaderColumnNumber(objBrowAndPageAndFrm.WebTable("tblTaxInformationHeader"),1,1,"W/Tax Amnt LC")
						
						strNameOfWTaxType = fGetCelldata(objBrowAndPageAndFrm.WebTable("tblTaxInformationData"),1,intNameOfWTaxTypeColNum,"Name of WTax Type")
						strWTaxCode = fGetCelldata(objBrowAndPageAndFrm.WebTable("tblTaxInformationData"),1,intWTaxCodeColNum,"WTax Code")
						strWTaxBase = fGetCelldata(objBrowAndPageAndFrm.WebTable("tblTaxInformationData"),1,intWTaxBaseColNum,"W/Tax Base")
						strWTaxAmt = fGetCelldata(objBrowAndPageAndFrm.WebTable("tblTaxInformationData"),1,intWTaxAmtColNum,"W/Tax Amt")
						strWTaxBaseLC = fGetCelldata(objBrowAndPageAndFrm.WebTable("tblTaxInformationData"),1,intWTaxBaseLCColNum,"W/Tax Base LC")
						strWTaxAmntLC = fGetCelldata(objBrowAndPageAndFrm.WebTable("tblTaxInformationData"),1,intWTaxAmntLCColNum,"W/Tax Amnt LC")
						
						Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strNameOfWTaxType,"TestData","Name of WTax Type")
						Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strWTaxCode,"TestData","Wtax Code")
						Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strWTaxBase,"TestData","W/Tax base")
						Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strWTaxAmt,"TestData","W/Tax Amt")
						Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strWTaxBaseLC,"TestData","W/Tax Base LC")
						Call fWriteOutputValueInExcel(Environment("TestName"),gTestExecutionIteration,strWTaxAmntLC,"TestData","W/Tax Amnt LC")
						
					End If
						If objBrowAndPageAndFrm.SAPButton("btnContinue").Exist(1) Then
							Call fClick(objBrowAndPageAndFrm.SAPButton("btnContinue"),"Continue")
						End If					
			Else
				Call fRptWriteReport("Fail", "Verify Display Document Icon in Searched page","Display document is not displayed in Searched results page")
                Call fRptWriteResultsSummary() 
                Exit Function
            End If
        Set objBrowAndPage = Nothing
        Set objBrowAndPageAndFrm = Nothing
    
End Function











Function fFioriProcessInvoiceManually(strInvoiceOrCreditMemo)
	
	
	On error Resume Next
		
	Dim objPgInvoiceProcess
	
	Set objPgInvoiceProcess = Browser("BrProcessing").Page("PgProcessing").SAPFrame("frmProcessInvOrCM")
	
	Set objBrwAndPg = Browser("brFioriAutoDesk").Page("pgFioriAutoDesk").SAPFrame("frmFioriAutoDesk")	
	Call fFioriExecteInvoice(objBrwAndPg,strInvoiceOrCreditMemo)			
	'Click on Post Invoice			
	If fVerifyObjectExist(objPgInvoiceProcess.SAPButton("btnPostInvoice")) Then
		Call fClick(objPgInvoiceProcess.SAPButton("btnPostInvoice"),"Post Invoice")
		Call fSynUntilObjExists(objPgInvoiceProcess.SAPButton("btnPostInvoiceCreditMemoOK"),2)
	Else
		Call fRptWriteReport("Fail", "Verify Post Invoice button","Post Invoice not displayed")
	End If
	'Click on OK
	If fVerifyObjectExist(objPgInvoiceProcess.SAPButton("btnPostInvoiceCreditMemoOK")) Then
		Call fClick(objPgInvoiceProcess.SAPButton("btnPostInvoiceCreditMemoOK"),"OK")
		Call fSynUntilObjExists(objPgInvoiceProcess.SAPButton("btnPostInvoiceCreditMemoOK"),MIN_WAIT)
	End  IF
	'Click on OK
	If fVerifyObjectExist(objPgInvoiceProcess.SAPButton("btnPostInvoiceCreditMemoOK")) Then
		Call fClick(objPgInvoiceProcess.SAPButton("btnPostInvoiceCreditMemoOK"),"OK")
		Call fSynUntilObjExists(objPgInvoiceProcess.SAPButton("btnPostInvoiceOrCreditMemo"),MIN_WAIT)
		Call fSynUntilObjExists(objPgInvoiceProcess.SAPButton("btnPostInvoiceOrCreditMemo"),MIN_WAIT)
	End  IF
	'Click on POST
	If fVerifyObjectExist(objPgInvoiceProcess.SAPButton("btnPostInvoiceOrCreditMemo")) Then
		Call fClick(objPgInvoiceProcess.SAPButton("btnPostInvoiceOrCreditMemo"),"POST")
		Call fSynUntilObjExists(objBrwAndPg.SAPEdit("txtReference"),MID_WAIT)
	Else
		Call fRptWriteReport("Fail", "Verify Post button","Post not displayed")
	End  IF
		
	Set objPgInvoiceProcess = Nothing
	On error Goto 0
	
End Function
