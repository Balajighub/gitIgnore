On error resume next

'Global Variables
Public gstrResultsFolder
Public gstrProjectResultPath
Public gstrRootFolder
Public gstrFrameWorkFolder
Public gstrTestcasesPath
Public gstrTestData
Public gstrFile
Public gstrFolderName
Public gstrProjectResourcesPath
Public objExcel
Public objWorkbook
Public objSheet
Public objActiveWorkbook
Public objActiveSheet
Public gEnvStatus
Public gstrCurrentDrive
Public gstrProjectTestdataPath
Public gstrProjectObjectRepository
Public gstrLogFloder
Public gstrResultName
Public gstrProjectConfigFilePath
Public gstrGlobalLibraryFilePath
Public gstrCommonLibraryFilePath
Public gstrBusinessLibraryFilePath
Public gstrQCutilLibraryFilePath
Public gstrReportsLibraryFilePath
Public gstrCoreFrameworkFolder
Public gstrAcceleratorsFolder	
Public gstrProjectsFolder
public gstrResultsRootFolder
Public gstrRootFolderName
Public gstrProjectRecoveryScenariosPath
Public gstrProjectTestScenariosPath
Public gstrFrameworkUtilityLibrariesPath
Public gstrFrameworkGlobalSettingsPath
Public gstrAcceleratorsSAPFioriLibraryPath
Public gstrAcceleratorsSAPFioriORPath
Public gstrAcceleratorsSAPAribaLibraryPath
Public gstrAcceleratorsSAPAribaORPath
Public gstrAcceleratorsSAPConcurLibraryPath
Public gstrAcceleratorsSAPConcurORPath
Public gstrAcceleratorsSAPGUILibraryPath
Public gstrAcceleratorsSAPGUIORPath

Public gstrProjectTestPlanPath
Public gstrModule
public gstrTestPlanName
public gstrOutputTestDataFolderPath
Public gstrOutputTestDataFile

Public gstrStatusPath
Public gstrTestRailtFolder
Public trTID
Public trRID
Public gstrFrameworkName

Public gstrTestEnvironment
Public gstrChromeBrowser
Public gstrIEBrowser
Public gstrAribaApplicationURL
Public gstrAribaApplicationUsername
Public gstrAribaApplicationPassword
Public gstrAribaSupplierApplicationURL
Public gstrAribaSupplierUsername
Public gstrAribaSupplierPassword

Public gTestExecutionIteration

'Project Paths
gstrRootFolder= Split(Environment("TestDir"),"\Projects")(0)
gstrModule = Split(Split(Environment("TestDir"),"\Projects\")(1),"\")(0)

'Framework Subfolder Paths
gstrCoreFrameworkFolder = gstrRootFolder&"\Framework"
gstrAcceleratorsFolder = gstrRootFolder&"\Accelerators"
gstrProjectsFolder = gstrRootFolder&"\Projects"

'Prject Level Path
gstrFrameWorkFolder= gstrProjectsFolder&"\"&gstrModule
gstrTestRailtFolder= gstrProjectsFolder&"\"&gstrModule&"\API_TestRail\"

'Project Level Folders Path
'gstrProjectConfigFilePath = gstrFrameWorkFolder&"\Config\Config.xml"
gstrProjectResourcesPath =gstrFrameWorkFolder&"\Resources"
'gstrProjectTestdataPath = gstrFrameWorkFolder&"\TestData\"
'gstrProjectResultPath = gstrFrameWorkFolder&"\TestResults"
gstrProjectRecoveryScenariosPath = gstrFrameWorkFolder&"\RecoveryScenarios"
gstrProjectTestScenariosPath = gstrFrameWorkFolder&"\TestScripts"

gstrTestPlanName = "TestExecutionPlan"
gstrProjectTestPlanPath = gstrFrameWorkFolder&"\"&gstrTestPlanName
gstrProjectConfigFilePath = gstrProjectTestPlanPath&"\TestExecutionConfig.xml"
gstrProjectTestdataPath = gstrProjectTestPlanPath&"\TestData\"
gstrProjectResultPath = gstrProjectTestPlanPath&"\TestResults"
gstrProjectFilesPath = gstrProjectTestPlanPath&"\Files"
gstrProjectPDFFilesPath = gstrProjectFilesPath&"\PDFFiles"

'Framework Level Folder Paths
gstrFrameworkUtilityLibrariesPath = gstrCoreFrameworkFolder&"\UtilityLibraries"
gstrFrameworkGlobalSettingsPath = gstrCoreFrameworkFolder&"\GlobalSettings"

'Accelarator Level Folder Paths
gstrAcceleratorsSAPFioriLibraryPath=gstrAcceleratorsFolder&"\SAPFiori\BusinessFunctions"
gstrAcceleratorsSAPFioriORPath = gstrAcceleratorsFolder&"\SAPFiori\ObjectRepository"
gstrAcceleratorsSAPAribaLibraryPath = gstrAcceleratorsFolder&"\SAPAriba\BusinessFunctions"
gstrAcceleratorsSAPAribaORPath = gstrAcceleratorsFolder&"\SAPAriba\ObjectRepository"
gstrAcceleratorsSAPConcurLibraryPath = gstrAcceleratorsFolder&"\SAPConcur\BusinessFunctions"
gstrAcceleratorsSAPConcurORPath = gstrAcceleratorsFolder&"\SAPConcur\ObjectRepository"
gstrAcceleratorsSAPGUILibraryPath = gstrAcceleratorsFolder&"\SAPGUI\BusinessFunctions"
gstrAcceleratorsSAPGUIORPath = gstrAcceleratorsFolder&"\SAPGUI\ObjectRepository"
'gstrGlobalConfigPath = "C:\Autodesk_PTPGlobalConfig\TestExecutionConfig.xml"
gstrGlobalConfigPath = gstrProjectTestPlanPath&"\TestExecutionConfig.xml"

gstrFolderName = Environment("FolderName")
gstrResultName=gstrFolderName

'Global Syncronizations Statements
Const MIN_WAIT = 5
Const MID_WAIT = 10
Const MAX_WAIT = 30

'ErrorHandling Statements
On Error Resume Next
On Error Goto 0
Environment("ERRORFLAG") = True
Environment("StepFailed") = "NO"
Environment("ROWCOUNT") = 1
gTestExecutionIteration = 1

Public gstrDocumentFolder
    
gstrDocumentFolder = gstrRootFolder&"\Documentation"

