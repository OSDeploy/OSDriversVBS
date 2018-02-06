'============================================================================================== DISCLAIMER

' This script, macro, and other code examples for illustration only, without warranty either expressed or implied, including but not
' limited to the implied warranties of merchantability and/or fitness for a particular purpose. This script is provided 'as is' and the Author does not
' guarantee that the following script, macro, or code can be used in all situations.

'============================================================================================== DECLARATIONS

	Option Explicit
	'On Error Resume Next

'============================================================================================== USER CONFIGURABLE CONSTANTS

	Const Author			=	"David Segura"
	Const Web				=	"http://osdeploy.com"
	Const Company			=	""
	Const Script			=	"OSDrivers.vbs"
	Const Description		=	""
	Const Release			=	""
	Const Reference			=	""

	Const Title 			=	"OSDrivers"
	Const Version 			=	20180131
	Const VersionFull 		=	20180131.01
	Dim TitleVersion		:	TitleVersion = Title & " (" & Version & ")"

'============================================================================================== SYSTEM CONSTANTS

	Const ForReading			=	1
	Const ForWriting			=	2
	Const ForAppending			=	8
	Const OverwriteExisting 	=	True

	Const HKEY_CLASSES_ROOT		= 	&H80000000
	Const HKEY_CURRENT_USER		= 	&H80000001
	Const HKEY_LOCAL_MACHINE	= 	&H80000002
	Const HKEY_USERS			= 	&H80000003
	Const HKEY_CURRENT_CONFIG	= 	&H80000005

'============================================================================================== INI or TXT File
	'[OSDrivers]
	'Source = SELF				(Copies the Directory the INI file is in)
	'Source = Something.cab		(File in the same Directory as the INI)
	'
	'OperatingSystem = All
	'Windows 7 x86 = No
	'Windows 7 x64 = No
	'Windows 10 x86 = No
	'Windows 10 x64 = Yes
	'
	'
	'
	'



	
	
'============================================================================================== OBJECTS

	Dim objComputer				: 	objComputer				=	"."
	'Dim objComputer			: 	objComputer				=	GetObject("WinNT://.,computer")
	Dim objShell				: 	Set objShell			=	CreateObject("Wscript.Shell")
	Dim objShellApp				: 	Set objShellApp			=	CreateObject("Shell.Application")
	Dim objFSO					: 	Set objFSO 				=	CreateObject("Scripting.FileSystemObject")
	Dim objDictionary			: 	Set objDictionary		=	CreateObject("Scripting.Dictionary")
	Dim objWMIService			: 	Set objWMIService 		=	GetObject("winmgmts:{impersonationLevel = impersonate}!\\" & objComputer & "\root\cimv2")
	Dim objRegistry				: 	Set objRegistry 		=	GetObject("winmgmts:{impersonationLevel = impersonate}!\\" & objComputer & "\root\default:StdRegProv")

'============================================================================================== VARIABLES: SYSTEM
	
	Dim MyUserName				: 	MyUserName				= Lcase(objShell.ExpandEnvironmentStrings("%UserName%"))
	Dim MyComputerName			: 	MyComputerName			= Ucase(objShell.ExpandEnvironmentStrings("%ComputerName%"))
	Dim MyTemp					: 	MyTemp					= Lcase(objShell.ExpandEnvironmentStrings("%Temp%"))
	Dim MyWindir				: 	MyWindir				= Lcase(objShell.ExpandEnvironmentStrings("%Windir%"))
	Dim MySystemDrive			: 	MySystemDrive			= Lcase(objShell.ExpandEnvironmentStrings("%SystemDrive%"))
	Dim MyArchitecture			: 	MyArchitecture			= Lcase(objShell.ExpandEnvironmentStrings("%Processor_Architecture%"))
	If MyArchitecture = "amd64" Then MyArchitecture = "x64"
	Dim MyExitCode				: 	MyExitCode				= 0
	
	'Alternate Method using Function GetVar
	'Dim MyUserName				: 	MyUserName				= Lcase(GetVar("%UserName%"))
	'Dim MyComputerName			: 	MyComputerName			= Ucase(GetVar("%ComputerName%"))
	'Dim MyTemp					: 	MyTemp					= Lcase(GetVar("%Temp%"))
	'Dim MyWindir				: 	MyWindir				= Lcase(GetVar("%Windir%"))
	'Dim MySystemDrive			: 	MySystemDrive			= Lcase(GetVar("%SystemDrive%"))
	'Dim MyArchitecture			: 	MyArchitecture			= Lcase(GetVar("%Processor_Architecture%"))
	
'============================================================================================== VARIABLES: CURRENT DIRECTORY

	Dim MyScriptFullPath		: 	MyScriptFullPath			= Wscript.ScriptFullName							'Full Path and File Name with Extension
	Dim MyScriptFileName		: 	MyScriptFileName			= objFSO.GetFileName(MyScriptFullPath)				'File Name with Extension
	Dim MyScriptBaseName		: 	MyScriptBaseName			= objFSO.GetBaseName(MyScriptFullPath)				'File Name 
	Dim MyScriptParentFolder	: 	MyScriptParentFolder		= objFSO.GetParentFolderName(MyScriptFullPath)		'Current Directory (Parent Folder)
	Dim MyScriptGParentFolder	: 	MyScriptGParentFolder		= objFSO.GetParentFolderName(MyScriptParentFolder)	'Parent of the Current Directory (Parent of the Parent Folder)
	Dim arrNames				:	arrNames					= Split(MyScriptParentFolder, "\")
	Dim intIndex				:	intIndex					= Ubound(arrNames)
	Dim MyParentFolderName		:	MyParentFolderName			= arrNames(intIndex)

'============================================================================================== VARIABLES: LOGGING

	'Only one line below must be uncommented
	Dim MyLogFile				: 	MyLogFile					= MyTemp & "\" & Title & ".log"						'Places the LOG in the Temp Directory	
	'Dim MyLogFile				: 	MyLogFile					= MyScriptParentFolder & "\" & Title & ".log"		'Places the LOG in the Script Directory
	'Dim MyLogFile				:	MyLogFile					= MyScriptParentFolder & "\" & "OSDrivers Script " & MyParentFolderName & ".log"
	'Dim MyLogFile				:	MyLogFile					= MyTemp & "\" & "OSDrivers Script " & MyParentFolderName & ".log"
	
	'Only one line below must be uncommented
	Dim DoLogging				: 	DoLogging					= True		'Creates a LOG
	'Dim DoLogging				: 	DoLogging					= False		'Prevents a LOG from being written

	'Only one line below must be uncommented
	'Dim TextFormat				: 	TextFormat					= True		'Results in a TEXT formatted LOG
	Dim TextFormat				: 	TextFormat					= False		'Results in a CMTRACE formatted LOG (default)
	
	LogStart					'Generate the LOG file

	'Identify
	'Dim MyLogInstall			:	MyLogInstall				= MyScriptParentFolder & "\" & "OSDrivers " & MyParentFolderName & ".log"
	Dim MyLogInstall			:	MyLogInstall				= MyTemp & "\" & "OSDrivers " & MyParentFolderName & ".log"
	Dim LocalPathComplete
	
	Dim LogTypeInfo				:	LogTypeInfo					= 1
	Dim LogTypeWarning			:	LogTypeWarning				= 2
	Dim LogTypeError			:	LogTypeError				= 3
'==============================================================================================
	TraceLog "================================================================================= Processing System Checks", LogTypeInfo
	'Gets the current date as 8 digit like 20180125
	Dim MyFullDate				:	MyFullDate					= Year(Date) & Right(String(2, "0") & Month(date), 2) & Right(String(2, "0") & Day(date), 2)
	Dim objTextFile
	Dim Return
	Dim Failed
	
	IsAdmin						'Will return IsAdmin = True if it is running with Admin Rights
	IsSystem					'Checks to see if this is running under the System Account
	'CheckArguments				'Checks for Command Line Arguments

	TraceLog "<Property> MyUserName: " & MyUserName, LogTypeInfo
	TraceLog "<Property> MyComputerName: " & MyComputerName, LogTypeInfo
	TraceLog "<Property> MyTemp: " & MyTemp, LogTypeInfo
	TraceLog "<Property> MyWindir: " & MyWindir, LogTypeInfo
	If Lcase(MySystemDrive) = "x:" Then
		MySystemDrive = "C:"
	End If
	TraceLog "<Property> MySystemDrive: " & MySystemDrive, LogTypeInfo
	TraceLog "<Property> MyArchitecture: " & MyArchitecture, LogTypeInfo
	
'==============================================================================================
	TraceLog "================================================================================= Processing Operating System", LogTypeInfo
	Dim MyOperatingSystem
	GetMyOperatingSystem		'Checks the Operating System.  We can stop specific OS's in this Sub

'==============================================================================================
	TraceLog "================================================================================= Processing Computer Information", LogTypeInfo
	Dim MyComputerManufacturer, MyComputerModel, MyBIOSVersion
	GetMyComputerInfo
	
'==============================================================================================
	TraceLog "================================================================================= Processing MDT Information", LogTypeInfo
	
	Dim MDTDeployRoot
	Dim MDTOSDisk
	Dim MDTImageBuild
	Dim MDTImageFlags
	Dim MDTLogPath
	GetMyMDTInfo

	CheckArguments
'==============================================================================================
	Dim sCmd
	Dim OSDriversDir
	Dim OSDriversKeyFile
	Dim OSDriversKeyFileName
	Dim OSDriversExtractDir
	Dim OSDriversOK
	OSDriversOK = False
	
	Dim OldLogFile
	OldLogFile = MyLogFile
	Dim TXTpnpids, objPNPIDs
	
	'Need to create the Drivers directory so the LOG file can be created
	If MDTOSDisk <> "" Then
		sCmd = "cmd /c md " & Chr(34) & MDTOSDisk & "\Drivers" & Chr(34)
		TraceLog "Command: " & sCmd, LogTypeWarning
		objShell.Run sCmd, 1, True
		MyLogFile		= MDTOSDisk & "\Drivers\" & Title & ".log"
		TXTpnpids		= MDTOSDisk & "\Drivers\PNPIDs.csv"
	Else
		sCmd = "cmd /c md " & Chr(34) & MySystemDrive & "\Drivers" & Chr(34)
		TraceLog "Command: " & sCmd, LogTypeWarning
		objShell.Run sCmd, 1, True
		MyLogFile		= MySystemDrive & "\Drivers\" & Title & ".log"
		TXTpnpids		= MySystemDrive & "\Drivers\PNPIDs.csv"
	End If
	
	'Export all the PNPIDs from the system
	On Error Resume Next
	Set objPNPIDs = objFSO.OpenTextFile(TXTpnpids, ForWriting, True)
	Dim colPNPItems, objPNPItem
	objPNPIDs.WriteLine "DeviceID" & "," & "Manufacturer" & "," & "Description" & "," & "Caption" & "," & "ClassGuid" & "," & "Driver"
	Set colPNPItems = objWMIService.ExecQuery("Select * From Win32_PnPEntity")
	'Set objPNPIDs = objFSO.OpenTextFile(TXTpnpids, ForAppending, True)
		For Each objPNPItem In colPNPItems
			'MsgBox objPNPItem
			objPNPIDs.WriteLine objPNPItem.DeviceID & "," & objPNPItem.Manufacturer & "," & objPNPItem.Description & "," & objPNPItem.Caption & "," & objPNPItem.ClassGuid
		Next
	objPNPIDs.Close

	'Read the new file into a Variable
	Dim AllPNPIDs
	Set objTextFile = objFSO.OpenTextFile(TXTpnpids, ForReading)
	Do Until objTextFile.AtEndOfStream
		AllPNPIDs = AllPNPIDs & vbCrLF & objTextFile.ReadLine
	Loop
	objTextFile.Close
	
	TraceLog "Start Date and Time is "											& Now, 0
	GetINIFiles MyScriptParentFolder

'==============================================================================================
	TraceLog "================================================================================= Complete!", LogTypeInfo

'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Get INI Files
	' /////////////////////////////////////////////////////////
	Sub GetINIFiles(MyScriptParentFolder)
		Dim objFolder, objFile, objSubFolder
		
		'Get the Parent Directory of this file
		Set objFolder = objFSO.GetFolder(MyScriptParentFolder)
		
		'Go through each file in this Directory and all Subs
		For each objFile in objFolder.Files
			OSDriversKeyFileName = ""

			'If the file is a TXT or an INI file, we need to read it
			If UCase(objFSO.GetExtensionName(objFile.Name)) = "TXT" or UCase(objFSO.GetExtensionName(objFile.Name)) = "INI" or UCase(objFSO.GetExtensionName(objFile.Name)) = "OSD" Then
			
				'Assume the file is not OK for us to use
				OSDriversOK = False
				
				'Set the Directory we are looking in
				OSDriversDir = objFolder
				
				'Set the Answer File
				OSDriversKeyFileName = objFile.Name
					
				'Process the Answer File
				OSDriversKeyFile = OSDriversDir & "\" & OSDriversKeyFileName
				'TraceLog "Processing OSD File: " & OSDriversKeyFile, LogTypeInfo
				OSDriversConfig OSDriversKeyFile
			End If
		Next

		For each objSubFolder in objFolder.SubFolders
			GetINIFiles objSubFolder.Path
		Next
	End Sub
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Read INI File Values
	' /////////////////////////////////////////////////////////
	Sub OSDriversConfig(OSDriversKeyFile)
		Dim objTextFile
		Dim OSDriversFileReadAll
		
		OSDriversOK = True
		
		TraceLog "================================================================================= Processing:" & OSDriversKeyFile, LogTypeInfo
		'TraceLog OSDriversKeyFile, LogTypeInfo
		
		'Valid INI file must contain [OSDrivers]
		Set objTextFile = objFSO.OpenTextFile(OSDriversKeyFile, ForReading)
		OSDriversFileReadAll = objTextFile.ReadAll
		'MsgBox OSDriversFileReadAll
		objTextFile.Close
		
		If Not Instr(lcase(OSDriversFileReadAll),lcase("[osdrivers]")) > 0 Then
			TraceLog "Issue: Could not find [OSDrivers] in Key File", LogTypeWarning
			TraceLog "Issue: Make sure this file is ANSI encoded [" & OSDriversKeyFile & "]", LogTypeWarning
			Exit Sub
		End If
		'==============================================================================================
		Dim INISource
		'Reads the SOURCE entry in the TXT file
		INISource		= ReadIni(OSDriversKeyFile, "OSDrivers", "Source")
		
		'Check if there is an entry
		If INISource <> "" Then
			INISource = Trim(INISource)
			
			If Lcase(INISource) = "self" Then
				'The Source is the directory the INI file is in
				INISource = OSDriversDir
				TraceLog "[OSDrivers] Source: " & INISource, LogTypeInfo
			Else
				INISource = OSDriversDir & "\" & INISource
				TraceLog "[OSDrivers] Source: " & INISource, LogTypeInfo

				'Check to see if we are working with a CAB file
				If Lcase(Right(INISource,4)) = ".cab" Then
					'If Not objFSO.FileExists(OSDriversDir & "\" & INISource) Then
					If Not objFSO.FileExists(INISource) Then
						TraceLog "Not Found: " & INISource, LogTypeWarning
						INISource = ""
					End If
				End If
			End If
		End If
		
		'Find the CAB
		If INISource = "" Then
			If objFSO.FileExists(OSDriversDir & "\" & objFSO.GetBaseName(OSDriversDir & "\" & OSDriversKeyFile) & ".cab") Then
				INISource = OSDriversDir & "\" & objFSO.GetBaseName(OSDriversDir & "\" & OSDriversKeyFile) & ".cab"
				TraceLog "Found: " & INISource, LogTypeInfo
			End If
		End If
		
		'At this point if INISource is still blank, then we assume that it is a Directory since a CAB was not located
		If INISource = "" Then
			'Could not locate the Source, Skip Driver
			TraceLog "Issue: Could not locate the Source", LogTypeWarning
			Exit Sub
		End If
		'==============================================================================================
		Dim INIDestination
		INIDestination	= ReadIni(OSDriversKeyFile, "OSDrivers", "Destination")
		INIDestination	= Trim(INIDestination)

		If INIDestination = "" Then
			INIDestination = "C:\Drivers" & "\" & objFSO.GetBaseName(INISource)
		ElseIf Lcase(Left(INIDestination, 3)) <> "c:\" Then
			INIDestination = "C:\Drivers" & "\" & INIDestination
		End If
		
		'Need to change the Driver Letter to match MDT's OSDisk
		If MDTOSDisk <> "" Then
			If Ucase(MDTOSDisk) <> "C:" Then
				'Replace the Left 2 Characters
				INIDestination = Replace(INIDestination,Left(INIDestination,2),MDTOSDisk)
			End If
		End If
				
		TraceLog "[OSDrivers] Destination: " & INIDestination, LogTypeInfo
		'==============================================================================================
		' /////////////////////////////////////////////////////////
		' Operating System Validation
		'
		' If not declared, the default is No for all OS's
		' If the INI is to apply to all OS's, then 
		'	
		'	OperatingSystem		= All
		'	Windows 7 x86		= No
		'	Windows 7 x64		= No
		'	Windows 10 x86		= No
		'	Windows 10 x64		= Yes
		' /////////////////////////////////////////////////////////
		Dim INIOperatingSystem, INIOperatingSystems
		
		'Section = [OSDrivers]
		'OperatingSystem = All
		'If this entry is set to All, then the Driver will be installed regardless of Operating System
		INIOperatingSystems	= ReadIni(OSDriversKeyFile, "OSDrivers", "OperatingSystem")
		INIOperatingSystems	= Trim(INIOperatingSystems)
		If INIOperatingSystems <> "" Then TraceLog "[OSDrivers] OperatingSystem: " & INIOperatingSystems, LogTypeInfo
		
		'Section = [OSDrivers]
		'Validate the Host Operating System to see if we should install in this Operating System
		INIOperatingSystem	= ReadIni(OSDriversKeyFile, "OSDrivers", MyOperatingSystem & " " & MyArchitecture)
		INIOperatingSystem	= Trim(INIOperatingSystem)
		If INIOperatingSystem <> "" Then TraceLog "[OSDrivers] " & MyOperatingSystem & " " & MyArchitecture & ": " & INIOperatingSystem, LogTypeInfo
		
		If Lcase(INIOperatingSystem) = "no" Then
			'Operating System was set to No, then skip this Driver
			TraceLog "Skipping Driver Installation: This Driver is NOT supported on " & MyOperatingSystem & " " & MyArchitecture, LogTypeInfo
			Exit Sub
		ElseIf Lcase(INIOperatingSystem) = "yes" Then
			'Operating System was set to Yes, then we are OK
		ElseIf Lcase(INIOperatingSystems) = "all" or Lcase(INIOperatingSystems) = "yes" Then
			'All Operating Systems was set to Yes, then we are OK
		Else
			'Operating System was not OK, then skip this Driver
			TraceLog "Skipping Driver Installation: This Driver is NOT supported on " & MyOperatingSystem & " " & MyArchitecture, LogTypeInfo
			Exit Sub
		End If
		'==============================================================================================
		Dim INIComputerMake
		INIComputerMake = Trim(ReadIni(OSDriversKeyFile, "OSDrivers", "Make"))
		INIComputerMake	= Trim(INIComputerMake)
		TraceLog "[OSDrivers] Make: " & INIComputerMake, LogTypeInfo
		
		If INIComputerMake = "" Then
			'TraceLog "Success: No Condition Specified", LogTypeInfo
		Else
			If NOT Instr(lcase(INIComputerMake), lcase(MyComputerManufacturer)) > 0 Then
				TraceLog "Skipping Driver Installation: This system does not meet the Computer Make requirements.", LogTypeInfo
				Exit Sub
			Else
				'TraceLog "Success: This system meets the Computer Make requirements.", LogTypeInfo
			End If
		End If
		'==============================================================================================
		Dim INIComputerModel
		INIComputerModel = Trim(ReadIni(OSDriversKeyFile, "OSDrivers", "Model"))
		TraceLog "[OSDrivers] Model: " & INIComputerModel, LogTypeInfo
		
		If INIComputerModel = "" Then
			'TraceLog "Success: No Condition Specified", LogTypeInfo
		Else
			If NOT Instr(lcase(INIComputerModel), lcase(MyComputerModel)) > 0 Then
				TraceLog "Skipping Driver Installation: This system does not meet the Computer Model requirements.", LogTypeInfo
				Exit Sub
			Else
				'TraceLog "Success: This system meets the Computer Model requirements.", LogTypeInfo
			End If
		End If
		'==============================================================================================
		Dim INIComputerNotMake
		INIComputerNotMake = Trim(ReadIni(OSDriversKeyFile, "OSDrivers", "NotMake"))
		INIComputerNotMake	= Trim(INIComputerNotMake)
		TraceLog "[OSDrivers] NotMake: " & INIComputerNotMake, LogTypeInfo
		
		If INIComputerNotMake = "" Then
			'TraceLog "Success: No Condition Specified", LogTypeInfo
		Else
			If Instr(lcase(INIComputerNotMake), lcase(MyComputerManufacturer)) > 0 Then
				TraceLog "Skipping Driver Installation: This system does not meet the Computer Not Make requirements.", LogTypeInfo
				Exit Sub
			Else
				'TraceLog "Success: This system meets the Computer Not Make requirements.", LogTypeInfo
			End If
		End If
		'==============================================================================================
		Dim INIComputerNotModel
		INIComputerNotModel = Trim(ReadIni(OSDriversKeyFile, "OSDrivers", "NotModel"))
		TraceLog "[OSDrivers] NotModel: " & INIComputerNotModel, LogTypeInfo
		
		If INIComputerNotModel = "" Then
			'TraceLog "Success: No Condition Specified", LogTypeInfo
		Else
			If Instr(lcase(INIComputerNotModel), lcase(MyComputerModel)) > 0 Then
				TraceLog "Skipping Driver Installation: This system does not meet the Computer Not Model requirements.", LogTypeInfo
				Exit Sub
			Else
				'TraceLog "Success: This system meets the Computer Not Model requirements.", LogTypeInfo
			End If
		End If
		'==============================================================================================
		Dim INIMasterPNPID
		INIMasterPNPID = Trim(ReadIni(OSDriversKeyFile, "OSDrivers", "MasterPNPID"))
		TraceLog "[OSDrivers] MasterPNPID: " & INIMasterPNPID, LogTypeInfo
		
		If INIMasterPNPID = "" Then
			'TraceLog "Success: No Condition Specified", LogTypeInfo
		Else
			INIMasterPNPID = Replace(INIMasterPNPID,"\","\\")
			TraceLog "Master WMI Query: Select * FROM Win32_PnPEntity WHERE DeviceID LIKE '%" & INIMasterPNPID & "%'", LogTypeInfo
		
			On Error Resume Next
			Dim colPNPItems, objPNPItem
			Set colPNPItems = objWMIService.ExecQuery("Select * From Win32_PnPEntity WHERE DeviceID LIKE '%" & INIMasterPNPID & "%'")
			If colPNPItems.Count <> 0 Then
				TraceLog "Success: MasterPNPID is found in the following " & colPNPItems.Count & " devices", LogTypeInfo
				For Each objPNPItem In colPNPItems
					TraceLog "***** " & objPNPItem.DeviceID, LogTypeInfo
				Next
			Else
				TraceLog "Skipping Driver Installation: MasterPNPID Hardware is NOT in this Computer", LogTypeInfo
				Exit Sub
			End If
		End If
		'==============================================================================================
		'Get PNPIDS
		Dim osdPNPIDs
		Set objTextFile = objFSO.OpenTextFile(OSDriversKeyFile, ForReading)
		Do Until objTextFile.AtEndOfStream
			osdPNPIDs = objTextFile.ReadLine
			Trim(osdPNPIDs)
			
			'Keep reading until we find [PNPIDS]
			If Instr(lcase(osdPNPIDs),lcase("[pnpids]")) > 0 Then
				TraceLog osdPNPIDs, LogTypeInfo
				Do Until objTextFile.AtEndOfStream
					osdPNPIDs = objTextFile.ReadLine
					
					'Stop if we get to the next section in the INI
					If Instr(osdPNPIDs,"[") > 0 Then Exit Do
					
					'Since we have some PNPIDS in here, we need to check for a value
					If osdPNPIDs <> "" Then
						'Set an exception as default until we find the hardware
						OSDriversOK = 0
					
						'Check if we have the hardware
						PNPCheck osdPNPIDs
						
						If OSDriversOK = True Then Exit Do
					End If
					'Log the Line
					'TraceLog osdPNPIDs, LogTypeInfo
				Loop
			End If
		Loop
		objTextFile.Close
		
		'==============================================================================================
		'==============================================================================================
		
		If OSDriversOK = True Then
			'Dim oFile
			
			OSDriversExtractDir = INIDestination
			'OSDriversExtractDir = INIDestination & "\" & objFSO.GetBaseName(INISource)
			'TraceLog "OSDriversExtractDir: " & OSDriversExtractDir, LogTypeInfo
			
			If Right(lcase(INISource),4) = ".cab" Then
				'Set oFile = objFSO.GetFile(INISource)
				'TraceLog "Creating " & OSDriversExtractDir & " on " & oEnvironment.Item("OSDisk"), LogTypeInfo
				TraceLog "Creating " & OSDriversExtractDir, LogTypeInfo
				sCmd = "cmd /c md " & Chr(34) & OSDriversExtractDir & Chr(34)
				TraceLog sCmd, LogTypeWarning
				objShell.Run sCmd, 1, True
				
				If objFSO.FileExists(INISource) Then
					sCmd = "cmd /c expand " & Chr(34) & INISource & Chr(34) & " -F:* " & Chr(34) & OSDriversExtractDir & Chr(34)
					TraceLog sCmd, LogTypeWarning
					objShell.Run sCmd, 1, True
				Else
					TraceLog INISource, LogTypeWarning
				End If
			Else
				sCmd = "robocopy " & Chr(34) & OSDriversDir & Chr(34) & " " & Chr(34) & OSDriversExtractDir & Chr(34) & " *.* /e /ndl /nfl /r:0 /w:0 /xj /xf " & Chr(34) & OSDriversKeyFile & Chr(34)
				TraceLog sCmd, LogTypeWarning
				objShell.Run sCmd, 1, True
			End If
		Else
			TraceLog "Skipping Driver Installation: This Driver does not apply to this Computer Hardware.", LogTypeInfo
		End If
	End Sub
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check PNP Hardware
	' /////////////////////////////////////////////////////////
	Function PNPCheck(osdPNPIDs)
		Dim PNPID	:	PNPID = Replace(osdPNPIDs,"\","\\")
		Dim PNPArr	:	PNPArr = Split(PNPID, " ", 2, 1)
		If UBound(PNPArr) >= 1 Then
			TraceLog "Checking for Device " & Chr(34) & Trim(Replace(PNPArr(1),"=","")) & Chr(34) & " in WMI Query: Select * FROM Win32_PnPEntity WHERE DeviceID LIKE '%" & PNPArr(0) & "%'", LogTypeInfo
		Else
			TraceLog "WMI Query: Select * FROM Win32_PnPEntity WHERE DeviceID LIKE '%" & PNPArr(0) & "%'", LogTypeInfo
		End If
		

		If Instr(lcase(AllPNPIDs),Replace(lcase(PNPArr(0)),"\\","\")) > 0 Then
			TraceLog "Found Matching Hardware:", LogTypeWarning
			
			Dim colItems, objItem
			Set colItems = objWMIService.ExecQuery("Select * From Win32_PnPEntity WHERE DeviceID LIKE '%" & PNPArr(0) & "%'")
			For Each objItem in colItems
				TraceLog "***** DeviceID = " & objItem.DeviceID, LogTypeInfo
				TraceLog "***** Manufacturer = " & objItem.Manufacturer, LogTypeInfo
				TraceLog "***** Name = " & objItem.Name, LogTypeInfo
				TraceLog "***** Caption = " & objItem.Caption, LogTypeInfo
				TraceLog "***** Description = " & objItem.Description, LogTypeInfo
				'TraceLog "***** PNPClass = " & objItem.PNPClass & " " & objItem.ClassGuid, LogTypeInfo
			Next
			OSDriversOK = True
		End If

		Exit Function
		
		'On Error Resume Next
		'Dim colItems, objItem
		Set colItems = objWMIService.ExecQuery("Select * From Win32_PnPEntity WHERE DeviceID LIKE '%" & PNPArr(0) & "%'")
		If colItems.Count <> 0 Then
			TraceLog "Found " & colItems.Count & " devices that match", LogTypeWarning
			For Each objItem in colItems
				TraceLog "***** " & objItem.Description & " (" & objItem.DeviceID & ")", LogTypeInfo
			Next
			'Found the hardware so we can copy the driver
			OSDriversOK = True
			Exit Function
		End If
	End Function
'==============================================================================================
	' /////////////////////////////////////////////////////////
	'	Function to check if Path is Writable
	' /////////////////////////////////////////////////////////
	
	Function IsPathWriteable(Path)
		Dim Temp_Path 'As String
		
		Temp_Path = Path & "\" & objFSO.GetTempName() & ".drs"
		
		On Error Resume Next
			objFSO.CreateTextFile Temp_Path
			IsPathWriteable = Err.Number = 0
			objFSO.DeleteFile Temp_Path
		On Error Goto 0
	End Function
'==============================================================================================
	' /////////////////////////////////////////////////////////
	'	Function to save Environment Variable to a Variable
	'
	'	Usage:	MyWindir = Lcase(GetVar("%Windir%"))
	'	Result:	MyWindir = c:\windows
	' /////////////////////////////////////////////////////////
	
	Function GetVar(sVar)
		'Using Windows Shell, return the value of an environment variable
		GetVar = objShell.ExpandEnvironmentStrings(sVar)
	End Function
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check if we have Admin Rights
	' /////////////////////////////////////////////////////////
	'	Usage:	If IsAdmin = False Then Wscript.Quit
	'	Result:	Script will exit
	'
	'	Usage:	If IsAdmin = False Then DoElevate
	'	Result:	Script will run the DoElevate Subroutine
	
	Function IsAdmin
		'LogLine
		'TraceLog "Function IsAdmin", 1
		
		Dim RegKey
		IsAdmin = False
		On Error Resume Next
		
		'Try to read a Registry Key that is only readable with Admin Rights
		RegKey = CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\")
		If Err.Number = 0 Then IsAdmin = True
		
		'Log Result
		If IsAdmin = True Then TraceLog "<IsAdmin = True> User has Admin Rights", 1
		If IsAdmin = False Then TraceLog "<IsAdmin = False> User does not have Admin Rights", 1
	End Function
'==============================================================================================
	Function CheckAdminRights
		Dim RegKey
		On Error Resume Next
		
		RegKey = CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\")
		
		If err.number <> 0 Then
			HasAdminRights = "NO"
			If MyOperatingSystem = "Windows XP" and Lcase(CreateObject("WScript.Network").UserName) = "administrator" Then HasAdminRights = "YES"
			If Lcase(CreateObject("WScript.Network").UserName) = "system" Then HasAdminRights = "YES"
		Else
			HasAdminRights = "YES"
		End If
	End Function
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check if we are running under SYSTEM Account
	' /////////////////////////////////////////////////////////
	Function IsSystem
		'LogLine
		'TraceLog "Function IsSystem", 1
		
		IsSystem = False
		
		'Determine if we are running this under the System Account and LOG result
		If Lcase(CreateObject("WScript.Network").UserName) = "system" Then
			IsSystem = True
			TraceLog "<IsSystem = True> Script is being run under the SYSTEM context, possibly from SCCM or as a Scheduled Task", LogTypeWarning
		Else
			TraceLog "<IsSystem = False> Script is NOT being run under the SYSTEM context", 1
		End If
	End Function
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check for Command Line Arguments
	' /////////////////////////////////////////////////////////
	Sub CheckArguments
		'LogLine
		TraceLog "Sub CheckArguments", 1
		
		Dim objArgs, strArg, i, colNamedArguments
	
		'Join all Arguments from Array to Variable ArgsFull
		'ReDim ArgsFull(WScript.Arguments.Count-1)
		'For i = 0 To WScript.Arguments.Count-1
		'  ArgsFull(i) = WScript.Arguments(i)
		'Next
		'ArgsFull = Join(ArgsFull)
		'TraceLog "Property ArgsFull: " & ArgsFull, 1

		' Store the Arguments in an Array
		Set objArgs = Wscript.Arguments
		
		Set colNamedArguments = WScript.Arguments.Named

		If colNamedArguments.Exists("OS") Then
			TraceLog "Argument Operating System: " & colNamedArguments.Item("OS"), 1
			MyOperatingSystem = colNamedArguments.Item("OS")
			TraceLog "<Variable> MyOperatingSystem = " & MyOperatingSystem, LogTypeWarning
		End If
		
		If colNamedArguments.Exists("Arch") Then
			TraceLog "Argument Architecture: " & colNamedArguments.Item("Arch"), 1
			MyArchitecture = colNamedArguments.Item("Arch")
			TraceLog "<Variable> MyArchitecture = " & MyArchitecture, LogTypeWarning
		End If
	End Sub
'==============================================================================================
'==============================================================================================
'==============================================================================================
'==============================================================================================
'============================================================================================== REFERENCE: DIALOG BOXES

	REM Constant			Value			Description
	REM vbOKOnly				0			Display OK button only.
	REM vbOKCancel				1			Display OK and Cancel buttons.
	REM vbAbortRetryIgnore		2			Display Abort, Retry, and Ignore buttons.
	REM vbYesNoCancel			3			Display Yes, No, and Cancel buttons.
	REM vbYesNo					4			Display Yes and No buttons.
	REM vbRetryCancel			5			Display Retry and Cancel buttons.
	REM vbCritical				16			Display Critical Message icon.
	REM vbQuestion				32			Display Warning Query icon.
	REM vbExclamation			48			Display Warning Message icon.
	REM vbInformation			64			Display Information Message icon.
	REM vbDefaultButton1		0			First button is default.
	REM vbDefaultButton2		256			Second button is default.
	REM vbDefaultButton3		512			Third button is default.
	REM vbDefaultButton4		768			Fourth button is default.
	REM vbApplicationModal		0			Application modal; the user must respond to the message box before continuing work in the current application.
	REM vbSystemModal			4096		System modal; all applications are suspended until the user responds to the message box.
	REM vbMsgBoxHelpButton		16384		Adds Help button to the message box
	REM VbMsgBoxSetForeground	65536		Specifies the message box window as the foreground window
	REM vbMsgBoxRight			524288		Text is right aligned
	REM vbMsgBoxRtlReading		1048576		Specifies text should appear as right-to-left reading on Hebrew and Arabic systems
	
'============================================================================================== FUNCTIONS: TRACE LOGGING
	' /////////////////////////////////////////////////////////
	' Logging Function with Trace Log
	' /////////////////////////////////////////////////////////
	Function TraceLog(LogText, LogError)
		Dim LogTemp
		Dim FileOut, MyLogFileX, TitelX, Tst
	
		If DoLogging = False Then Exit Function
		
		If TextFormat = True Then
			If LogError = 0 Then
				Set FileOut = objFSO.OpenTextFile( MyLogFile, ForWriting, True)
			Else
				Set FileOut = objFSO.OpenTextFile( MyLogFile, ForAppending, True)
			End If
			FileOut.WriteLine Now()& " - " & LogText
			FileOut.Close
			Set FileOut = Nothing
			Exit Function
		End If
	
		On Error Resume Next
		Tst = KeineLog
		On Error Goto 0
		If UCase( Tst ) = "JA" Then Exit Function

		On Error Resume Next
		TitelX = Titel
		' if not set 'Titel' outside procedure 'TitelX' is empty
		TitelX = title
		' if not set 'title' outside procedure 'TitelX' is empty

		If Len( TitelX ) < 2 Then TitelX = document.title
		' set title in .HTA
		If Len( TitelX ) < 2 Then TitelX = WScript.ScriptName
		' set title in .VBS
		On Error Goto 0

		On Error Resume Next
		MyLogFileX = MyLogFile
		' if not set 'MyLogFile' outside procedure, 'MyLogFileX' is empty
		If Len( MyLogFileX ) < 2    Then MyLogFileX = WScript.ScriptFullName & ".log"' .vbs
		If Len( MyLogFileX ) < 2    Then MyLogFileX = TitelX & ".log"        ' .hta
		On Error Goto 0

		' Enumerate Milliseconds
		Tst = Timer()               ' timer() in USA: 1234.22; dot separation
		Tst = Replace( Tst, "," , ".")        ' timer() in german: 23454,12; comma separation
		If InStr( Tst, "." ) = 0 Then Tst = Tst & ".000"
		Tst = Mid( Tst, InStr( Tst, "." ), 4 )
		If Len( Tst ) < 3 Then Tst = Tst & "0"

		' Enumerate Time Zone
		Dim AktDMTF : Set AktDMTF = CreateObject("WbemScripting.SWbemDateTime")
		AktDMTF.SetVarDate Now(), True : Tst = Tst & Mid( AktDMTF, 22 ) ' : MsgBox Tst, , "099 :: "
		' MsgBox "AktDMTF: '" & AktDMTF & "'", , "100 :: "
		Set AktDMTF = Nothing
		LogTemp = LogText
		LogTemp = "<![LOG[" & LogTemp & "]LOG]!>"
		LogTemp = LogTemp & "<"
		LogTemp = LogTemp & "time=""" & Hour( Time() ) & ":" & Minute( Time() ) & ":" & Second( Time() ) & Tst & """ "
		LogTemp = LogTemp & "date=""" & Month( Date() ) & "-" & Day( Date() ) & "-" & Year( Date() ) & """ "
		LogTemp = LogTemp & "component=""" & TitelX & """ "
		LogTemp = LogTemp & "context="""" "
		LogTemp = LogTemp & "type=""" & LogError & """ "
		LogTemp = LogTemp & "thread=""0"" "
		LogTemp = LogTemp & "file=""David.Segura"" "
		LogTemp = LogTemp & ">"

		Tst = 8							'ForAppending
		If LogError = 0 Then Tst = 2	'ForWriting

		Set FileOut = objFSO.OpenTextFile( MyLogFileX, Tst, True)
		If     LogTemp = vbCRLF Then FileOut.WriteLine ( LogTemp )
		If Not LogTemp = vbCRLF Then FileOut.WriteLine ( LogTemp )
		FileOut.Close
		Set FileOut	= Nothing
		'Set objFSO	= Nothing
	End Function
	' /////////////////////////////////////////////////////////
	' Trace Log Solid Line
	' /////////////////////////////////////////////////////////
	Sub LogLine
			TraceLog "=================================================================================", 1
	End Sub
	' /////////////////////////////////////////////////////////
	' Trace Log Blank Space
	' /////////////////////////////////////////////////////////
	Sub LogSpace
		TraceLog "", 1
	End Sub
	' /////////////////////////////////////////////////////////
	' Trace Log Contents
	' /////////////////////////////////////////////////////////
	Sub LogStart
		'Tracelog "Start a new Log File", 0											'Clears any existing content
		'TraceLog "This is a standard line", 1										'Create an Entry
		'TraceLog "This is a warning line", 2										'Create an Entry and highlight yellow (Warning)
		'TraceLog "This is an error line", 3										'Create an Entry and highlight red (Error or Critical)
		'LogSpace																	'Create a Line without content
		'LogLine																	'Create a Line with =====================================

		If WScript.Arguments.length = 0 Then TraceLog "Starting "					& WScript.ScriptFullName, 0
		If WScript.Arguments.length <> 0 Then TraceLog "Starting "					& WScript.ScriptFullName, 2
		TraceLog "Start Date and Time is "											& Now, 1
		TraceLog "Script Last Modified: " 											& CreateObject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).DateLastModified, 1
		LogLine

	End Sub
'==============================================================================================
Function ReadIni( myFilePath, mySection, myKey )
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )
	
    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Replace(Replace(Trim(objIniFile.ReadLine), vbTab, ""), """", "")
			
            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
				'Abort if the end of the INI file is reached
                If objIniFile.AtEndOfStream Then Exit Do
				strLine = Replace(Replace(Trim(objIniFile.ReadLine), vbTab, ""), """", "")
                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ))
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ))
							'MsgBox ReadIni
                            ' In case the item exists but value is blank
							'If ReadIni = "" Then
                            '    ReadIni = " "
                            'End If
                            'Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Replace(Replace(Trim(objIniFile.ReadLine), vbTab, ""), """", "")
                Loop
				Exit Do
            End If
        Loop
        objIniFile.Close
		'MsgBox "Section was not found"
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If
End Function


'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check Operating System Properties
	' /////////////////////////////////////////////////////////
	Sub GetMyOperatingSystem
		'LogLine
		TraceLog "Sub GetMyOperatingSystem", 1
		
		Dim objItem, colItems
		Dim Unsupported
		
		On Error Resume Next
		Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
		For Each objItem In colItems
			TraceLog "<Property> Caption: " & objItem.Caption,1
			TraceLog "<Property> OperatingSystemSKU: " & objItem.OperatingSystemSKU,1
			TraceLog "<Property> Organization: " & objItem.Organization,1
			TraceLog "<Property> OSArchitecture: " & objItem.OSArchitecture,1
			TraceLog "<Property> OSProductSuite: " & objItem.OSProductSuite,1
			TraceLog "<Property> OSType: " & objItem.OSType,1
			TraceLog "<Property> ProductType: " & objItem.ProductType,1
			TraceLog "<Property> RegisteredUser: " & objItem.RegisteredUser,1
			TraceLog "<Property> SerialNumber: " & objItem.SerialNumber,1
			TraceLog "<Property> Status: " & objItem.Status,1
			TraceLog "<Property> SuiteMask: " & objItem.SuiteMask,1
			TraceLog "<Property> Version: " & objItem.Version,1
			
			With objItem
			Select Case True
				'Client Operating Systems
				Case Left(.Version,3) = "5.1" and .ProductType = 1
					MyOperatingSystem = "Windows XP"
				Case Left(.Version,3) = "5.2" and .ProductType = 1
					MyOperatingSystem = "Windows XP"
				Case Left(.Version,3) = "6.0" and .ProductType = 1
					MyOperatingSystem = "Windows Vista"
				Case Left(.Version,3) = "6.1" and .ProductType = 1
					MyOperatingSystem = "Windows 7"
				Case Left(.Version,3) = "6.2" and .ProductType = 1
					MyOperatingSystem = "Windows 8"
				Case Left(.Version,3) = "6.3" and .ProductType = 1
					MyOperatingSystem = "Windows 8.1"
				Case Left(.Version,3) = "10." and .ProductType = 1
					MyOperatingSystem = "Windows 10"
				'Server Operating Systems
				Case Left(.Version,3) = "5.2" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2003"
				Case Left(.Version,3) = "6.0" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2008"
				Case Left(.Version,3) = "6.1" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2008 R2"
				Case Left(.Version,3) = "6.2" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2012"
				Case Left(.Version,3) = "6.3" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2012 R2"
				Case Left(.Version,3) = "10." and .ProductType > 1
					MyOperatingSystem = "Windows Server 2016"
				End Select
			End With
			
			'If MyOperatingSystem = "" Then MyOperatingSystem = objOperatingSystem.Caption
			If MyOperatingSystem = "" or Unsupported = True Then
				MyOperatingSystem = objItem.Caption
				TraceLog "<Property> MyOperatingSystem = " & MyOperatingSystem, 3
				TraceLog MyOperatingSystem & " is not supported by this Script", 3
				Wscript.Quit
			End If
			TraceLog "<Variable> MyOperatingSystem = " & MyOperatingSystem, LogTypeWarning
			
			If Lcase(objItem.OSArchitecture) = "64-bit" Then
				MyArchitecture = "x64"
			End If
			
			TraceLog "<Variable> MyArchitecture = " & MyArchitecture, LogTypeWarning			
		Next
	End Sub
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check Computer Properties
	' /////////////////////////////////////////////////////////
	Sub GetMyComputerInfo
		'LogLine
		TraceLog "Sub GetMyComputerInfo", 1
		
		Dim objItem, colItems
		Dim Unsupported
		
		On Error Resume Next
		Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
		For Each objItem In colItems
			TraceLog "<Property> DNSHostName: "				& objItem.DNSHostName,1
			TraceLog "<Property> Domain: "					& objItem.Domain,1
			TraceLog "<Property> DomainRole: "				& objItem.DomainRole,1
			TraceLog "<Property> Manufacturer: "			& objItem.Manufacturer,1
			TraceLog "<Property> Model: "					& objItem.Model,1
			TraceLog "<Property> PartOfDomain: "			& objItem.PartOfDomain,1
			TraceLog "<Property> PrimaryOwnerName: "		& objItem.PrimaryOwnerName,1
			TraceLog "<Property> TotalPhysicalMemory: "		& objItem.TotalPhysicalMemory,1

			MyComputerManufacturer = Trim(objItem.Manufacturer)
			
			With objItem
			Select Case True
				Case Instr(.Manufacturer, "Dell") > 0
					MyComputerManufacturer	= "Dell"
				Case Instr(.Manufacturer, "Hewlett") > 0
					MyComputerManufacturer	= "HP"
				Case Instr(.Manufacturer, "Microsoft") > 0
					MyComputerManufacturer	= "Microsoft"
			End Select
			End With
			
			MyComputerModel	= Trim(objItem.Model)
			
			TraceLog "<Variable> MyComputerManufacturer = " & MyComputerManufacturer, 2
			TraceLog "<Variable> MyComputerModel = " 		& MyComputerModel, 2
		Next
		
		Dim BIOS
		For Each BIOS in GetObject("winmgmts:\\.\root\cimv2").InstancesOf("Win32_BIOS")  
			MyBIOSVersion = Trim(Ucase(BIOS.SMBIOSBIOSVERSION))
			TraceLog "<Property> MyBIOSVersion: "		& MyBIOSVersion,1
		Next
	End Sub
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Get MDT Information
	' /////////////////////////////////////////////////////////
	Sub GetMyMDTInfo
		TraceLog "Sub GetMyMDTInfo", LogTypeInfo
		Dim oTSEnv, oVar
		On Error Resume Next
		Set oTSEnv = CreateObject("Microsoft.SMS.TSEnvironment")
		For Each oVar In oTSEnv.GetVariables
			'TraceLog oVar & "=" & oTSEnv(oVar), LogTypeInfo
			
			If Ucase(oVar)		= "DEPLOYROOT" Then
				MDTDeployRoot	= oTSEnv(oVar)
				TraceLog "<Variable> MDTDeployRoot = " & MDTDeployRoot, LogTypeWarning
			ElseIf Ucase(oVar)	= "OSDISK" Then
				MDTOSDisk 		= oTSEnv(oVar)
				TraceLog "<Variable> MDTOSDisk = " & MDTOSDisk, LogTypeWarning
			ElseIf Ucase(oVar)	= "IMAGEBUILD" Then
				MDTImageBuild	= oTSEnv(oVar)
				TraceLog "<Variable> MDTImageBuild = " & MDTImageBuild, LogTypeWarning
			ElseIf Ucase(oVar)	= "IMAGEFLAGS" Then
				MDTImageFlags	= oTSEnv(oVar)
				TraceLog "<Variable> MDTImageFlags = " & MDTImageFlags, LogTypeWarning
			ElseIf Ucase(oVar)	= "IMAGEPROCESSOR" Then
				MyArchitecture	= oTSEnv(oVar)
				TraceLog "<Variable> MyArchitecture = " & MyArchitecture, LogTypeWarning
			ElseIf Ucase(oVar)	= "LOGPATH" Then
				MDTLogPath	= oTSEnv(oVar)
				TraceLog "<Variable> MDTLogPath = " & MDTLogPath, LogTypeWarning
			End If
		Next
		Set oTSEnv = NOTHING
		
		If Left(MDTImageBuild,3) = "5.1" Then
			MyOperatingSystem = "Windows XP"
		ElseIf Left(MDTImageBuild,3) = "5.2" Then
			If Left(MDTImageFlags,6) = "Server" Then
				MyOperatingSystem = "Windows Server 2003"
			Else
				MyOperatingSystem = "Windows XP"	'x64 Edition
			End If
		ElseIf Left(MDTImageBuild,3) = "6.0" Then
			If Left(MDTImageFlags,6) = "Server" Then
				MyOperatingSystem = "Windows Server 2008"
			Else
				MyOperatingSystem = "Windows Vista"
			End If
		ElseIf Left(MDTImageBuild,3) = "6.1" Then
			If Left(MDTImageFlags,6) = "Server" Then
				MyOperatingSystem = "Windows Server 2008 R2"
			Else
				MyOperatingSystem = "Windows 7"
			End If
		ElseIf Left(MDTImageBuild,3) = "6.2" Then
			If Left(MDTImageFlags,6) = "Server" Then
				MyOperatingSystem = "Windows Server 2012"
			Else
				MyOperatingSystem = "Windows 8"
			End If
		ElseIf Left(MDTImageBuild,3) = "6.3" Then
			If Left(MDTImageFlags,6) = "Server" Then
				MyOperatingSystem = "Windows Server 2012 R2"
			Else
				MyOperatingSystem = "Windows 8.1"
			End If
		ElseIf Left(MDTImageBuild,3) = "10." Then
			If Left(MDTImageFlags,6) = "Server" Then
				MyOperatingSystem = "Windows Server 2016"
			Else
				MyOperatingSystem = "Windows 10"
			End If
		End If
	
		If MyOperatingSystem = "" Then MyOperatingSystem = "Unknown"
		TraceLog "<Variable> MyOperatingSystem = " & MyOperatingSystem, LogTypeWarning
	End Sub
'==============================================================================================