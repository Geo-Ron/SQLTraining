Const ScriptVersion = 20191105

'**************************************************************
' Login script gebouwd door Ron Peeters - r.peeters@vdl.nl
'
'	  ___            ___         _
'	 | _ \___ _ _   | _ \___ ___| |_ ___ _ _ ___
'	 |   / _ \ ' \  |  _/ -_) -_)  _/ -_) '_(_-<
'	 |_|_\___/_||_| |_| \___\___|\__\___|_|_/__/       _ _        _
'	  _ _  _ __  ___ ___| |_ ___ _ _ ___ / __ \__ ____| | |  _ _ | |
'	 | '_|| '_ \/ -_) -_)  _/ -_) '_(_-</ / _` \ V / _` | |_| ' \| |
'	 |_|(_) .__/\___\___|\__\___|_| /__/\ \__,_|\_/\__,_|_(_)_||_|_|
'		  |_|                            \____/
'
' 20100526	-	Eerste bouw
' 20100526	-	Login script gereed voor productie
' 20100927	-	Parameter aanpassingen gedaan.
'				Tevens Welcome MsgBox on top of everything
' 20101105	-	Wel of niet IE tonen ingebouw
' 				Mappings op basis van OU toegevoegd
' 20101210	-	Mapping naar LPT poort mogelijk gemaakt.
' 20101210	-	(BUG) Variabele voor persoonlijke schijf (U:) ingebouwd
'				Voorkomt loskoppelen nadat deze door AD gekoppelt is.
' 20101213	-	(BUG) Foutmelding aangepast bij persistent mappings.
' 20110131	-	(BUG) IE Voortgangsvenster verwijdert ivm compatibiliteit met
'				Windows 7
' 20110131	-	Eventlogging toegevoegd
' 20110201	-	Eventlogging uitgebreid naar verwijderde drives
' 20110207	-	(Cleaning) Dubbele functie verwijdert.
' 20110303	-	(BUG) Remap HomeDrive toegevoegd, ivm VPN problemen
' 20110311	-	Logging voor einde login script
' 20110317	-	Foutmelding trapping voor netwerkmappings aangepast
' 20110317-2-	Configfile login.ini.vbs ingevoerd
' 20110328	-	Removenetworkdrive updates User Profile
' 20110401	-	(BUG) Login Melding, tijdsberekening aangepast
' 20110525	-	(BUG) Groepslidmaatschap niet meer Case-sensitive
' 20120109  -   AD IntraForest lidmaatschap herkenning
' 20120109  -   Auto update functie bij inloggen
' 20120214	-	Add own domain to Local Intranet Zones in IE
' 20120628	-	Reset Persistent mapping by command line if not able to delete
' 20120830  -   Added RUN command
' 20121004	-	Eventlog shows share to map to
' 20121018	-	No WelcomeMSG if username is Administrator OR contains "extra"
' 20121108	-	Added KB937624 functionality
' 20121128	-	Added TMG Client Config
' 20130117	-	Remove BUG in TMG Client Config (echo all errors)
' 20130206  -   Enabled Timeout of 30s in every MsgBox that pops up
' 20130813	-	Change LDAP queries for performance on RODC
' 20130814	-	Bugfixes and implemented basic debug parameter. Less AD queries
' 20130827	-	BugFix Domain Users
' 20130917	-	Change most LDAP: queries to GC: queries (except MapHomeDrive) to enable binding to RODC
' 20130926	-	Bugfixes for GC queries when used cross forest and performance issues
' 20131001	-	Bugfix for DNSRoot to Local Intranet Zones
' 20131029	-	Will check if latency higher than 60ms, then wait until it droppes
' 20131112	-	Will Check Latency to GC you're logged onto. Else try %USERDNSDOMAIN%
' 20160406	-	Retry on drive map error
' 20160408	-	If mapped drive matches the one needed for drive letter and unc path, no action will be taken.
' 20160411	-	Script optimalisaties doorgevoerd.
' 20160412	-	BugFix via VPN. Alle koppelingen voor iedereen naar het einde van de ini file geschoven
'			-	HomeShare mapping lookup faalde via VPN. Dit ivm AD Site change. Code bevraagt nu
'				dichtsbijzijnde DC op basis van actuele data ipv %LOGONSERVER%
' 20160503	-	Local Intranet Zones - add VDLNET.nl and VDLGROEP.nl
' 20160609	-	Function to set homepage added
' 20160613	-	Preparing functions to merge VBS and normal login script... check AD OU of computer - next step choose login.ini.vbs based on computer ou.
' 20160825	-	Accepting Named parameters
'			-	inserted function hiding redirected drives in explorer
' 20160826	-	Merged VBS loging script and normal login script. Now we have the same Program Script again
'			-	VBScript uses different login.ini.vbs (using /inifile:<inifilename.ini.vbs>)
'			-	Including warning when using /debug to use cscript.exe instead of wscript.exe
' 20160829	-	Implemented restart script for RemoteApp Servers (using Login_RDSDisconnectDetection.vbs script that keeps running in the background)
' 20161010	-	Implemented duplicate run detection. Will quit reporting to eventlog with warning if so
' 20170328	-	Improved duplicate run detection
' 20180420	-	Bugfixes
' 20190221	-	added 3 functions to be able to map a shared printer on a client used to connect to a RDS Host
'						-	Function MapSharedClientPrinter(PRINTERNAMEPART) is used to search on LIKE basis and if found, map in RDSession.
' 20190328  -   Changed Function RenameMyComputer "Deze computer" to "Computer"
' 20190411	-	Changed function Get_AD_Ou. It breaks when a user account has a comma in it's Fullname. Behaviour will remain the same.
' 20190507	-	Bugfix in search client printers function. If you enumerate an array to count its contents, you need to define it first.
' 20191105	-	WHOAMI output. Session issue on Windows 10
' 20191106	-	KB937624 added for PowerUser of NetworkOperator elevation
'**************************************************************
'Error Handling en Variabelen declareren tbv KABI Accounts
ReDim Preserve outArray(3, 0)
Dim MappingArray()
Dim LDP, PingReply
Dim arrcount
Dim MappingCount
Dim Memberships
Dim OUImIn, BooTTime
Dim AdditionalSites

'**************************************************************
' Parameter Handling
'**************************************************************
If Wscript.Arguments.Count = 0 Then
     On error Resume Next
 Else
   '  For i = 0 to Wscript.Arguments.Count-1
   '      If Wscript.Arguments(i) = "/debug" Then
   '       debugflag=True 'Turn on the debug flag
			'wscript.echo "********************************"
			'wscript.echo "* Script Run in Debugging Mode *"
			'wscript.echo "********************************"
			'wscript.echo " "
   '     ElseIf Wscript.Arguments(i) = "/nolog" Then
   '       loggingflag = True 'Turn off the logging flag
		'ElseIf Wscript.Arguments(i) = "/?" OR Wscript.Arguments(i) = "?" Then
   '       wscript.echo "Only parameter allowed is /debug (use cscript to run it)"
		 ' wscript.echo "Example: cscript \\vdlgroep.local\netlogon\login.vbs /debug"
		 ' wscript.quit
		'Else
			'wscript.echo "Incorrect parameter. Try parameter /?"
			'wscript.quit
   'End If
   '  Next


Dim colNamedArguments,fso,ParamsForRun

Set colNamedArguments = Wscript.Arguments.Named
'Inifile Parameter Checking
	If colNamedArguments.Exists("inifile") Then
				'wscript.echo "**********************************"
				'wscript.echo "* Script uses alternate ini file *"
				'wscript.echo "**********************************"
				'wscript.echo " File found in parameter: "& colNamedArguments.Item("inifile")
				alternateINI = colNamedArguments.Item("inifile")
	End If
'Debug Parameter Checking
	If colNamedArguments.Exists("debug") Then
			If InStr(1, WScript.FullName, "WScript.exe", vbTextCompare) <> 0 Then
				wscript.echo "/debug needs to run using cscript! Will exit now."
				WScript.Quit(0)
			End If


			  debugflag=True 'Turn on the debug flag
				wscript.echo "********************************"
				wscript.echo "* Script Run in Debugging Mode *"
				wscript.echo "********************************"
				wscript.echo " "
	Else
		On error Resume Next
	End If
'Help Parameter Checking
	If colNamedArguments.Exists("?") or colNamedArguments.Exists("help") Then
			  debugflag=True 'Turn on the debug flag
			  wscript.echo "Pararameters allowed:"
				wscript.echo "/debug (use cscript to run it)"
				wscript.echo "/inifile:<fullfqdnpath>.ini.vbs (use cscript to run it) for alternate ini file"
			  wscript.echo "Example: cscript \\vdlgroep.local\netlogon\login.vbs /debug"
			  wscript.echo "Example: cscript \\vdlgroep.local\netlogon\login.vbs /inifile:\\vdlgroep.local\netlogon\alternative.ini.vbs"
			  wscript.quit
	End If
'**************************************************************
' End Of Parameter Handling
'**************************************************************
 End If

'**************************************************************

Call LogLogon 'Altijd als eerste!! Laat een waarschuwing in Application Eventlog zien als de gebruiker inlogt (Source WSH; EventID 2)
Call CheckDuplicateScriptInstance 'Controleert of het login script voor deze gebruiker al loopt. Zo ja, dan wordt het afgebroken en gemeld in het eventlog
If (RemoteAppServer <> True) then CALL CheckLatency 'Checks ping to %USERDNSDOMAIN% max 15 times for latency lower then 60ms
CALL KB937624 'Checks if script is run elevated and not on a server, if so then apply KB937624
'Call RemoveNetworkDrives 'Verwijdert alle gekoppelde netwerkshares m.u.v. vermeldde drive (bijv U:) �OBSOLETE
Call DisableIE8Customize 'Voorkomt het instellingen venster van IE8
Call RenameMyComputer() 'Zet de hostname bij "Deze Computer"
'Call ShowWinVer() 'Laat Windows versie op desktop zien
Call DisableNetworkPrinterBalloon() 'Schakelt netwerk printer balloon venster uit
AdditionalSites = Array("VDLNET.nl","VDLGROEP.nl") 'Additional sites in lokal intranet zones in IE
Call AddDomainSites 'Add own domain to Local Intranet Zones in IE



' *****************************************************
		' Load Configuration File
			if debugflag=True then
					wscript.echo "***************************************************"
					wscript.echo "* Ini File: Loading Ini File and start processing *"
					wscript.echo "***************************************************"

			End If
			err.Clear
			Set objNetwork = WScript.CreateObject("WScript.Network")
			Set WshShell = WScript.CreateObject("WScript.Shell")
			Set fso = CreateObject("Scripting.FileSystemObject")
			' Open Configuration File
			If alternateINI <> "" Then
				if debugflag=True then wscript.echo "Ini File: "&colNamedArguments.Item("inifile")
				WshShell.LogEvent 0, "Try to load Configfile "&colNamedArguments.Item("inifile")&""
				Set ConfigFile = fso.OpenTextFile(colNamedArguments.Item("inifile"),1,false)
				Inifile = alternateINI
			Else
				if debugflag=True then wscript.echo "Ini File: "&fso.GetParentFolderName(wscript.ScriptFullName)&"\login.ini.vbs"
				WshShell.LogEvent 0, "Try to load Configfile "&fso.GetParentFolderName(wscript.ScriptFullName)&"\login.ini.vbs"
				Set ConfigFile = fso.OpenTextFile(fso.GetParentFolderName(wscript.ScriptFullName)&"\login.ini.vbs",1,false)
				Inifile = fso.GetParentFolderName(wscript.ScriptFullName)&"\login.ini.vbs"
			End If

			'ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(fso.GetParentFolderName(wscript.ScriptFullName)&"\login.ini.vbs",1,false).ReadAll

			If Err.Number <> 0 Then
					MsgBoxPopup "Er vond een fout plaats tijdens het koppelen van het configuratiebestand." & VbCrLf & VbCrLf & "FoutCode: " & _
						Err.Number & VbCrLf & VbCrLf & "Beschrijving: " & VbCrLf & Err.Description & VbCrLf & VbCrLf & "File: "& Inifile & VbCrLf & "Neem contact op met de Servicedesk.", vbOKOnly + vbCritical, "Probleem koppelen Configuratie"
					WshShell.LogEvent 1, "Er vond een fout plaats tijdens het koppelen van het configuratiebestand." & VbCrLf & VbCrLf & "FoutCode: " & _
						Err.Number & VbCrLf & VbCrLf & "Beschrijving: " & VbCrLf & Err.Description & VbCrLf & VbCrLf & "File: "& Inifile & VbCrLf & "Neem contact op met de Servicedesk."
			Else
					WshShell.LogEvent 0, "Configfile "&IniFile &" geladen"
			End If
			Err.Clear

			ExecuteGlobal ConfigFile.ReadAll

			if debugflag=True then
					wscript.echo "*************************************"
					wscript.echo "* Ini File: Reached end of Ini File *"
					wscript.echo "*************************************"

			End If

' ********************************************************

If (RemoteAppServer <> True) then Call MapHomeDrive
If (RemoteAppServer <> True) then Call ShowWelcome(VDLCompany, Tekstregel1, Tekstregel2)
' Start Login_RDSDisconnectDetection.vbs tot rerun login script if session gets reconnected
' If (RemoteAppServer = "Doen we niet meer vanwege issues") then
' 	'Start Login_RDSDisconnectDetection.vbs
' 	Dim objShell
' 	Set objShell = Wscript.CreateObject("WScript.Shell")
' 	Set fso = CreateObject("Scripting.FileSystemObject")
' 	Set objArgs = Wscript.Arguments
' 	'Wscript.Echo "Command:"
' 	 totalArgs = ""
' 	 For Each strArg in objArgs
' 	   'WScript.Echo strArg
' 	   If Not Instr(lcase(strArg),"loginscript")>0 then totalArgs = totalArgs & " " &strArg
' 	 Next
' 	If debugflag=True Then
' 		Command = "CMD /C cscript "& fso.GetParentFolderName(wscript.ScriptFullName)&"\Login_RDSDisconnectDetection.vbs" & " " & totalArgs & " /loginscript:"& wscript.ScriptFullName & " & pause"
' 	Else
' 		'CMD wordt bewust gebruikt vanwege het veranderen van focus als dit script draait
' 		Command = "CMD /C cscript "& fso.GetParentFolderName(wscript.ScriptFullName)&"\Login_RDSDisconnectDetection.vbs" & " " & totalArgs & " /loginscript:"& wscript.ScriptFullName
' 	End If

' 	If Not colNamedArguments.Exists("rerun") Then
' 		if debugflag=True then
' 			Wscript.Echo "RDSReconnect Running: "&Command
' 			objShell.Run Command,1,false
' 		Else
' 			objShell.Run Command,0,false
' 		End If
' 	Else
' 		if debugflag=True then Wscript.Echo "RDSReconnect Rerun flag already set. No need to restart"
' 	End If

' End If

Call RemoveLeftOverNetworkDrives

'**************************************************************
if debugflag=True then
	wscript.echo " "
	wscript.echo "********************"
	wscript.echo "* Script Run Ended *"
	wscript.echo "********************"
End If

strUser = CreateObject("WScript.Network").UserName
WshShell.LogEvent 2, "Login Run of "&wscript.ScriptFullName&" has ended for user " & strUser
' Einde draaiende script
'**************************************************************

' ************************************************************************************************************************************************************************
' ************************************************************************************************************************************************************************
' ************************************************************** Programma Code bevindt zich onder dit stuk **************************************************************
' ************************************************************************************************************************************************************************
' ************************************************************************************************************************************************************************


' *****************************************************
' This add on will show Eventlog entry at logon
Function CalcBootTime()
if debugflag=True then wscript.echo "CalcBootTime: Query Boot Time"
if isEmpty(BootTime) then
	if debugflag=True then wscript.echo "CalcBootTime: Calculating Boot Time"
	On Error Resume Next
		strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colOperatingSystems = objWMIService.ExecQuery _
		("Select * from Win32_PerfFormattedData_PerfOS_System")
	 Set Count=0
	For Each objOS in colOperatingSystems
		Count = Count + 1
		'if debugflag=True then wscript.echo "LoopCount="&Count
		dtmUptime = objOS.SystemUpTime
		BootTime = dtmUptime /3600
		CalcBootTime = BootTime
	Next
Else
	CalcBootTime = BootTime
End If
End Function
' *****************************************************

' *****************************************************
' This add on will show Eventlog entry at logon
Function MapHomeDrive()
	if debugflag=False then On Error Resume Next

	Const ADS_SECURE_AUTHENTICATION = 1
	Const ADS_READONLY_SERVER = 4
	Const ADS_SERVER_BIND = 512

	Set WshShell = WScript.CreateObject("WScript.Shell")
	'WshShell.LogEvent 0, "Trying to find HOMESHARE..."
	if debugflag=True then wscript.echo "HOMEDRIVE: Trying to find HOMESHARE..."
		Set objConnection = CreateObject("ADODB.Connection")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Open "Active Directory Provider"
		Set objRootDSE = GetObject("GC://" & EnvString("userdnsdomain") & "/RootDSE")
		ClosestDC = objRootDSE.Get("dnsHostName")
	Set wshNetwork=CreateObject("WScript.Network")
	Set ADSysInfo=CreateObject("ADSystemInfo")
	'Set CurrentUser=GetObject("LDAP://" & EnvString("userdnsdomain") & "/" & ADSysInfo.UserName)
	Set CurrentUser=GetObject("LDAP://" & ClosestDC & "/" & ADSysInfo.UserName)
	'if debugflag=True then wscript.echo "Connected to: "& CurrentUser.Get("dnsHostName")
	'Set CurrentUser=GetObject("LDAP:").OpenDSObject("LDAP://" & ADSysInfo.UserName, Nothing, Nothing, ADS_SECURE_AUTHENTICATION Or ADS_READONLY_SERVER Or ADS_SERVER_BIND)
	if debugflag=True then wscript.echo "HOMEDRIVE: LDAP://" & ClosestDC & "/" & ADSysInfo.UserName
	'Currentuser.Getinfo
	'strHomeDirectory = Currentuser.Get("homeDirectory")
	'strHomeDrive = Currentuser.Get("homeDrive")
	strHomeDirectory = Currentuser.homedirectory
	strHomeDrive = Currentuser.homedrive
	'WshShell.LogEvent 1, "Finding homedrive, and found it!"
	'wscript.sleep 2000
	if debugflag=True then wscript.echo "HOMEDRIVE: Homedirectory: " & strHomeDirectory & ", Homedrive: " & strHomeDrive
	If strHomeDirectory <> "" AND strHomeDrive <> "" then
		Call KoppelShare(strHomeDrive,strHomeDirectory,"","","")
		WshShell.LogEvent 0, "Found homeshare at " & strHomeDirectory & " and mapped it to "& strHomeDrive
		if debugflag=True then wscript.echo "HOMEDRIVE: Found homeshare and mapped it if needed"
	else
		if debugflag=True then wscript.echo "HOMEDRIVE: No home share found in AD"
		WshShell.LogEvent 1, "No home share found in AD" & VbCrLf & "Homedirectory: "&strHomeDirectory& VbCrLf & "HomeDrive: "&strHomeDrive & VbCrLf & "ADSysInfo.Username: "&ADSysInfo.Username & VbCrLf & "Query: LDAP://" & ClosestDC & "/" & ADSysInfo.UserName
	End If
	'WshShell.LogEvent 0, "MapHomeDrive Finished"

End Function
' *****************************************************
' *****************************************************
' This add on will show Eventlog entry at logon
Function LogLogon()

  set WshShell = WScript.CreateObject( "WScript.Shell" )
    WshShell.LogEvent 2 ,"User " &EnvString("username") & " logged in"

End Function
' *****************************************************

' *****************************************************
' This add on will show Welcome Box
Function ShowWelcome(VDLCompany, Tekstregel1, Tekstregel2)
if debugflag=True then wscript.echo "Welcome: Starting Show Welcome"
' ---------------------------------------------------------------
' Algemene initialisatie Melding aan gebruik(st)ers
' ---------------------------------------------------------------
if debugflag=True then wscript.echo "Welcome: Running queries to show information"
	Set wshNetwork=CreateObject("WScript.Network")
	Set ADSysInfo=CreateObject("ADSystemInfo")
	Set CurrentUser=GetObject("GC://" & EnvString("userdnsdomain") & "/" & ADSysInfo.UserName)
	Set objSysInfo = CreateObject("ADSystemInfo")
		strUser = objSysInfo.UserName
		Set objUser = GetObject("GC://" & strUser)
		strADFullName = objUser.Get("displayName")
if debugflag=True then wscript.echo "Welcome: Queries done"
' ---------------------------------------------------------------
' Welkomtekst voorbereiden
' ---------------------------------------------------------------


	Dim Welkomtekst
	Dim dtmHour
	dtmHour = Hour(Now())
	' Bepaal welkomsttekst voor het logon scherm met de juiste tijdreferentie
	If dtmHour < 12 Then
	    strGreeting = "Goedemorgen "
	Else
	    strGreeting = "Goedemiddag "
	End If

	OpStartTijd = CalcBootTime()

	If COpStartTijd > 46 then
		Tekstregel2 = "Uw computer staat al enkele dagen aan! Zet de computer en het scherm 's avonds uit!"
	ElseIf OpStartTijd > 22 then
		Tekstregel2 = "Vergeet niet 's avonds je computer en scherm uit te zetten."
	Else
		Tekstregel2 = "Prettige werkdag gewenst."
	End If


	Welkomtekst = strGreeting & strADFullName & chr(13) & chr(10) & chr(13) & chr(10) & "Het is nu " & Now()  & chr(13) & chr(10) & chr(13) & chr(10) & "U bent ingelogd op het netwerk van " &VDLCompany&" via computer " & wshnetwork.computername & chr(13) & chr(10) & chr(13) & chr(10) & Tekstregel1 & chr(13) & chr(10) & chr(13) & chr(10) & Tekstregel2

set WshShell = WScript.CreateObject( "WScript.Shell" )
    WshShell.LogEvent 0 ,VDLCompany & ": "& WelkomTekst
' ---------------------------------------------------------------
' Als laatste laat het welkomscherm zien
' ---------------------------------------------------------------
	MsgBoxPopup WelkomTekst , vbOKOnly + vbInformation, "Welkom op het netwerk van "& VDLCompany


End Function
' *****************************************************


' *****************************************************
' This add on will show Windows version on Desktop
Function DisableIE8Customize()

  set WshShell = WScript.CreateObject( "WScript.Shell" )
  Path = "HKCU\Software\Microsoft\Internet Explorer\Main\DisableFirstRunCustomize"
  WshShell.RegWrite Path, 1 ,"REG_DWORD"

End Function
' *****************************************************

' *****************************************************
' This add on will show Windows version on Desktop
Function ShowWinVer()

  Set WSHNetwork = CreateObject("WScript.Network")
  HKEY_CURRENT_USER = &H80000001
  strComputer = WSHNetwork.Computername
  Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
  strKeyPath = "Control Panel\Desktop"
  objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
  ValueName = "PaintDesktopVersion"
  dwValue = 0
  objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue
End Function


' *****************************************************
' This add on will show Windows version on Desktop
Function DisableNetworkPrinterBalloon()

  set WshShell = WScript.CreateObject( "WScript.Shell" )
  Path = "HKCU\Printers\Settings\EnableBalloonNotificationsRemote"
  WshShell.RegWrite Path, 0 ,"REG_DWORD"

End Function


' *****************************************************
' This add on will rename the My Computer icon with the computer name
Function RenameMyComputer()

	set objShell = WScript.CreateObject( "WScript.Shell" )
	Set objNetwork = WScript.CreateObject("WScript.Network")
	strComputer = objNetwork.Computername
	MCPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
	ObjShell.RegWrite MCPath & "\", "Computer "&strComputer, "REG_SZ"
End Function

' *****************************************************


' *****************************************************
'This function checks to see if the passed group name contains the current
' user as a member. Returns True or False
Function IsMember(groupname)
If debugflag=False then On Error Resume Next
IsMember = False

If isEmpty(MemberShips) then
	'Set WshShell = WScript.CreateObject("WScript.Shell")
	'WshShell.LogEvent 1, "Groups not researched. Running query (should only run one time!)"
	if debugflag=True then wscript.echo "GroupMembership: Querying AD for group membership (should only run one time!)"
	Set Memberships = SearchGroups
	'WshShell.LogEvent 1, "Query Ran (should only run one time!)"
	if isEmpty(Memberships) then Memberships="No_Results"
End If

if debugflag=True then wscript.echo "GroupMembership: "&groupname & " Membership will be checked"
IsMember = CBool(MemberShips.Exists(groupName))
if debugflag=True then wscript.echo "GroupMembership: Membership result is "&IsMember

End Function

Function SearchGroups
  Dim objADSysInfo, objConnection, objRootDSE, objRecordSet, objGroups
  Dim strUserDN, strFilter
	if debugflag=False then on error resume next
  Set objADSysInfo = CreateObject("ADSystemInfo")
  'strUserDN = EnvString(UserName)
  strUserDN = objADSysInfo.UserName
  'strFilter = "(member=" & strUserDN & ")"
  ' Alternative search filter to test nested groups
   strFilter = "(member:1.2.840.113556.1.4.1941:=" & strUserDN & ")"

  Set objConnection = CreateObject("ADODB.Connection")
  objConnection.Provider = "ADsDSOObject"
  objConnection.Open "Active Directory Provider"

  Set objRootDSE = GetObject("GC://" & EnvString("userdnsdomain") & "/RootDSE")
  'Set objRootDSE = GetObject("GC://RootDSE")
  'Set objRootDSE = GetObject("LDAP://RootDSE")
  if debugflag=True then wscript.echo "GroupMembership: Connected to: "& objRootDSE.Get("dnsHostName")
  Set objRecordSet = objConnection.Execute( _
    "<GC://" & EnvString("userdnsdomain") & "/" & objRootDSE.Get("defaultNamingContext") & ">;" & _
    strFilter & ";distinguishedName,name;subtree")

  Set objGroups = CreateObject("Scripting.Dictionary")
  objGroups.CompareMode = VbTextCompare

  While Not objRecordSet.EOF
    strGroup = objRecordSet.Fields("name").Value
    If Not objGroups.Exists(strGroup) Then
      objGroups.Add UCase(strGroup), ""
	  if debugflag=True then wscript.echo "GroupMembership: Group added: "&strGroup
    End If
	If Not objGroups.Exists("Domain Users") Then
		objGroups.Add UCase("Domain Users"), ""
		if debugflag=True then wscript.echo "GroupMembership: Group added: "&"Domain Users"
	End If
    objRecordSet.MoveNext
  WEnd

  Set SearchGroups = objGroups
End Function
' *****************************************************

' *****************************************************
'This function returns a particular environment variable�s value.
' for example, if you use EnvString("username"), it would return
' the value of %username%.
Function EnvString(variab)
If debugflag=False then On Error Resume Next
    set objShell = WScript.CreateObject("WScript.Shell")
    variable = "%" & variab & "%"
    EnvString = ucase(objShell.ExpandEnvironmentStrings(variable))
    Set objShell = Nothing
End Function
' *****************************************************

' *****************************************************
'De koppeling van Share verzorgen
Function KoppelShare(strDriveLetter,strRemotePath,strUsername,strPassword,strDomain)
  If debugflag=False then On Error Resume Next


	TryCount = 0
	MountError = 0
	NeedToMount = 0
	MountErrorDescription = ""

	strSaveMappingInProfile = "FALSE"
	'--------------------------------------------------
	Set objNetwork = WScript.CreateObject("WScript.Network")
	Set WshShell = WScript.CreateObject("WScript.Shell")

	ReDim Preserve MappingArray(MappingCount)
	if debugflag=True then wscript.echo "DriveMapping: Adding "&ucase(strDriveLetter)&" to MappingArray at position "&MappingCount
	MappingArray(MappingCount) = ucase(strDriveLetter)
	Mappingcount = Mappingcount +1


	'Set CheckDrive = objNetwork.EnumNetworkDrives()

	'For intDrive = 0 To CheckDrive.Count - 1 Step 2
	'	If CheckDrive.Item(intDrive) =strDriveLetter _
	'	Then objNetwork.RemoveNetworkDrive strDriveLetter
	'Next

	NeedToMount = CheckForNetworkDriveMatch(strDriveLetter,strRemotePath)
	if debugflag=True then wscript.echo "DriveMapping: Need to Mount = "&NeedToMount

	IF (NeedToMount > 1) then
		'Matched same driveletter but not for UNC path, removed
		if debugflag=True then wscript.echo "DriveMapping: Need to Mount DriveMapping "&StrDriveLetter
			Do
				Err.Clear
				TryCount = TryCount + 1
				MountError = 0
				MountErrorDescription = ""
				if debugflag=True then wscript.echo "DriveMapping: Trying to map " &strDriveLetter & ", Try# "&TryCount

				If strUserName = "" then
						objNetwork.MapNetworkDrive strDriveLetter, strRemotePath, strSaveMappingInProfile
				ELSE
						objNetwork.MapNetworkDrive strDriveLetter, strRemotePath, strSaveMappingInProfile, strDomain & "\" & strUserName, strPassword
				END If

				MountError = Err.Number
				MountErrorDescription = Err.Description

				If MountError <> 0 Then
					If MountError = -2147024811 then
						if debugflag=False then MsgBoxPopup "Schijf "& strDriveLetter &" is reeds in gebruik en kan dus niet aangekoppeld worden!" & VbCrLf & "Neem contact op met de Servicedesk.", vbOKOnly + vbExclamation, "Probleem met "&strDriveLetter &" schijf koppeling!"
						WshShell.LogEvent 1, "Schijf "& strDriveLetter &" is reeds in gebruik en kan dus niet aangekoppeld worden!" & VbCrLf & "Het betreft share " & strRemotePath & " die aangekoppeld zou worden."& VbCrLf & "needtoMount waarde: "&NeedtoMount
					'ElseIf Err.Number = -2147023694 then
						'MsgBoxPopup "Schijf "& strDriveLetter &" is handmatig (persistent) door de gebruiker aangekoppeld" & VbCrLf & "en kan daarom niet opnieuw verbonden worden!" & VbCrLf & "Verwijder de netwerkverbinding," & VbCrLf & "of neem contact op met de Servicedesk.", vbOKOnly, "Probleem met Persistent "&strDriveLetter &" schijf koppeling!"
						'WshShell.LogEvent 1, "Schijf "& strDriveLetter &" is handmatig (persistent) door de gebruiker aangekoppeld" & VbCrLf & "en kan daarom niet opnieuw verbonden worden!" & VbCrLf & "Verwijder de netwerkverbinding," & VbCrLf & "of neem contact op met de Servicedesk."
					Else
						if debugflag=False then MsgBoxPopup "Er vond een fout plaats tijdens het koppelen van de "& strDriveLetter &" schijf." & VbCrLf & VbCrLf & "FoutCode: " & _
							MountError & VbCrLf & VbCrLf & "Beschrijving: " & VbCrLf & MountErrorDescription & VbCrLf & VbCrLf & "Share: "& strRemotePath & VbCrLf & "Neem contact op met de Servicedesk.", vbOKOnly + vbCritical, "Probleem koppelen netwerkshare"
						WshShell.LogEvent 1, "Er vond een fout plaats tijdens het koppelen van de "& strDriveLetter &" schijf." & VbCrLf & VbCrLf & "FoutCode: " & _
							MountError & VbCrLf & VbCrLf & "Beschrijving: " & VbCrLf & MountErrorDescription & VbCrLf & VbCrLf & "Share: "& strRemotePath & VbCrLf & "Neem contact op met de Servicedesk."
					End If
					if debugflag=False then WScript.Sleep(10000)
				Else
					'MsgBoxPopup "Schijf "& strDriveLetter &" succesvol gekoppeld.", vbOKOnly, strDriveLetter &"-schijf koppeling"
					WshShell.LogEvent 0, "Schijf "& strDriveLetter &" succesvol gekoppeld."& VbCrLf & VbCrLf & "Share: " & strRemotePath
					if debugflag=True then wscript.echo "DriveMapping: Schijf "& strDriveLetter &" succesvol gekoppeld aan share "&strRemotePath
				End If
				if ((debugflag=True) AND (MountError <> 0)) then wscript.echo "DriveMapping: Failed to map " &strDriveLetter & ", Try# "&TryCount &" Err.Number: "&MountError&" Share: "&strRemotePath
				if ((debugflag=True) AND (MountError <> 0)) then wscript.echo "DriveMapping: Failure description: "&MountErrorDescription

			Loop Until (((MountError = 0) OR (TryCount > 2)) OR (MountError = -2147024811))
	ELSE
		if debugflag=True then wscript.echo "DriveMapping: Matched same driveletter and UNC path, no need to mount drive"
		WshShell.LogEvent 0, "Schijf "& strDriveLetter &" niet opnieuw gekoppeld, deze is al in orde."& VbCrLf & VbCrLf & "Share: " & strRemotePath
	End If


	Err.Clear
End Function
' *****************************************************

Function RemoveNetworkDrives
	If debugflag=False then On Error Resume Next
	Set WshShell = WScript.CreateObject("WScript.Shell")

	DIM objNetwork,colDrives,i
	Set objNetwork = CreateObject("Wscript.Network")
	Set colDrives = objNetwork.EnumNetworkDrives
	For i = 0 to colDrives.Count-1 Step 2
	Err.Clear
		' Force Removal of network drive and remove from user profile
		' objNetwork.RemoveNetworkDrive strName, [bForce], [bUpdateProfile]
		'WshShell.LogEvent 2, "Removing "&colDrives.Item(i) &" number " & i

		if debugflag=True then wscript.echo "Removing "&colDrives.Item(i) &" number " & i
		objNetwork.RemoveNetworkDrive colDrives.Item(i),TRUE,TRUE
			If Err.Number <> 0 then
				WshShell.LogEvent 1, "Error removing "&colDrives.Item(i)
				Err.Clear
				Set objShell = WScript.CreateObject("WScript.shell")
				'objShell.run "cmd /C net use * /delete /yes",2, False
				objShell.run "cmd /C net use /persistent:no",2, False
				'objShell.run "cmd /C net use * /delete /yes",2, False
		Set objShell = Nothing
			Else
				WshShell.LogEvent 0, "Succesfull removed "&colDrives.Item(i)
				if debugflag=True then wscript.echo "Succesfull removed "&colDrives.Item(i)
			End If


	Next
End Function

' *****************************************************

Function CheckForNetworkDriveMatch (strDriveLetter,strRemotePath)
	If debugflag=False then On Error Resume Next
	Set WshShell = WScript.CreateObject("WScript.Shell")

	DIM objNetwork,colDrives,i, Matched
	Set objNetwork = CreateObject("Wscript.Network")
	Set colDrives = objNetwork.EnumNetworkDrives

	Matched = 0

	For i = 0 to colDrives.Count-1 Step 2
		Err.Clear
			' Force Removal of network drive and remove from user profile
			' objNetwork.RemoveNetworkDrive strName, [bForce], [bUpdateProfile]
			'WshShell.LogEvent 2, "Removing "&colDrives.Item(i) &" number " & i
		'if debugflag=True then wscript.echo "NetworkDrive: Checking "&colDrives.Item(i) &" mapping at "&colDrives.Item(i+1)

		IF (lcase(colDrives.Item(i)) = lcase(strDriveLetter)) then
			Matched = 1
			if debugflag=True then wscript.echo "NetworkDrive: Is a match for "&colDrives.Item(i) &" will check UNC path"
			if debugflag=True then wscript.echo "NetworkDrive: Question: Is it a match for mapped "&colDrives.Item(i+1) &" and UNC Patch "&strRemotePath&"?"
			IF ((lcase(colDrives.Item(i)) = lcase(strDriveLetter)) AND (lcase(colDrives.Item(i+1)) = lcase(strRemotePath))) then
				'IF Matches do not remove
				if debugflag=True then wscript.echo "NetworkDrive: Is a match for "&colDrives.Item(i) &" and UNC Patch "&strRemotePath&". No Need to remove"
				WshShell.LogEvent 0, colDrives.Item(i)&" is already mapped to "& colDrives.Item(i+1) & " and it has been requested to be mapped as "&strRemotePath & ". No Action will be taken"
				CheckForNetworkDriveMatch = 1 'Match no need to remove
			ELSE
				'Does not match, so will remove
				if debugflag=True then wscript.echo "NetworkDrive: Is a match for "&colDrives.Item(i) &" but not for UNC Patch "&strRemotePath&". Need to remove!"
				if debugflag=True then wscript.echo "NetworkDrive: Removing "&colDrives.Item(i) &" number " & i
				WshShell.LogEvent 0, colDrives.Item(i)&" is already mapped to "& colDrives.Item(i+1) & " and it has been requested to be mapped as "&strRemotePath & ". Previous mapping will be removed."

				objNetwork.RemoveNetworkDrive colDrives.Item(i),TRUE,TRUE

				If Err.Number <> 0 then
					WshShell.LogEvent 1, "Error removing "&colDrives.Item(i)&" it is mapped to "& colDrives.Item(i+1)
					Err.Clear
					Set objShell = WScript.CreateObject("WScript.shell")
					'objShell.run "cmd /C net use * /delete /yes",2, False
					objShell.run "cmd /C net use /persistent:no",2, False
					'objShell.run "cmd /C net use * /delete /yes",2, False
					Set objShell = Nothing
					CheckForNetworkDriveMatch = 3 'No Match, error removing
				Else
					WshShell.LogEvent 0, "Succesfull removed "&colDrives.Item(i)
					if debugflag=True then wscript.echo "Succesfull removed "&colDrives.Item(i)
					CheckForNetworkDriveMatch = 2 'No Match, removed
				End If
			End If
		End If
	Next

	IF Matched = 0 then CheckForNetworkDriveMatch = 4 'No Match, not mapped at all

End Function

' *****************************************************
Function RemoveLeftOverNetworkDrives
	if debugflag=False then On Error Resume Next
	Set WshShell = WScript.CreateObject("WScript.Shell")
	if debugflag=True then wscript.echo "RemoveMappings: Alle mappings die NIET via het script gedaan zijn, worden verwijdert!"
	DIM objNetwork,colDrives,i
	Set objNetwork = CreateObject("Wscript.Network")
	Set colDrives = objNetwork.EnumNetworkDrives
	For i = 0 to colDrives.Count-1 Step 2
	Err.Clear
		' Force Removal of network drive and remove from user profile
		' objNetwork.RemoveNetworkDrive strName, [bForce], [bUpdateProfile]
		'WshShell.LogEvent 2, "Removing "&colDrives.Item(i) &" number " & i

		MappingGevonden = False
		For z = 0 to UBound(MappingArray)
		'if debugflag=True then wscript.echo "MappingArray: "&z
			If (UCASE(MappingArray(z)) = UCASE(colDrives.Item(i))) THEN
				'if debugflag=True then wscript.echo "RemoveMappings: Mapping gevonden: "&colDrives.Item(i)&" Arrayvalue "&MappingArray(z)&" die gemapped hoort te worden, niets mee doen."
				MappingGevonden = True
			End If
			If ((Instr(UCASE(colDrives.Item(i+1)),"TSCLIENT")) AND colDrives.Item(i) = "")THEN
				'if debugflag=True then wscript.echo "RemoveMappings: Mapping gevonden naar "&colDrives.Item(i+1)&". Dit is een Drive Redirection share. Niets mee doen."
				MappingGevonden = True
			End If

		Next

		IF MappingGevonden = False and colDrives.Item(i) <> "" THEN
			if debugflag=True then wscript.echo "RemoveMappings: Mapping gevonden: "&colDrives.Item(i)&" welke verwijdert dient te worden"
			if debugflag=True then wscript.echo "RemoveMappings: Removing "&colDrives.Item(i) &" number " & i
			objNetwork.RemoveNetworkDrive colDrives.Item(i),TRUE,TRUE
				If Err.Number <> 0 then
					WshShell.LogEvent 1, "Error removing "&colDrives.Item(i)& ", path: "&colDrives.Item(i+1)
					Err.Clear
					Set objShell = WScript.CreateObject("WScript.shell")
					objShell.run "cmd /C net use "&colDrives.Item(i)&" /delete /yes",2, False
					objShell.run "cmd /C net use /persistent:no",2, False
					'objShell.run "cmd /C net use * /delete /yes",2, False
			Set objShell = Nothing
				Else
					WshShell.LogEvent 0, "Succesfull removed "&colDrives.Item(i)
					if debugflag=True then wscript.echo "Succesfull removed "&colDrives.Item(i)
				End If
		ELSE
			if debugflag=True then wscript.echo "RemoveMappings: Mapping gevonden naar "&colDrives.Item(i)&" ("&colDrives.Item(i+1)&"). Niets mee doen."
		End If

	Next
End Function

' *****************************************************




' *****************************************************
'  Controleer OU van gebruiker
' *****************************************************
Function Get_AD_Ou (depth)
	On Error Resume Next

if IsEmpty(OUImIn) then
	Const ADS_SCOPE_SUBTREE = 2

	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection

	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE

	objCommand.CommandText = _
		"SELECT distinguishedName FROM 'GC://" & EnvString("userdnsdomain") &"' " & _
			"WHERE objectCategory='user' " & _
				"AND sAMAccountName='" & EnvString("username") &"'"
	Set objRecordSet = objCommand.Execute

	objRecordSet.MoveFirst
	Do Until objRecordSet.EOF
		OUImIn = objRecordSet.Fields("distinguishedName").Value
		if debugflag=True then wscript.echo "AD Query for OU I'm In ="& vbCrLf&" " &OUImIn & " !"& vbCrLf&" Should only run once!"
		'Workaround for comma in AD User Full Name
		OUImIn = mid(OUImIn,instr(OUImIn,"OU"))
		'Need to re-add CN, else depth will be wrong
		OUImIn = "CN=HIERSTAATDEUSERFULLNAMENORMALITER," & OUImIn
		if debugflag=True then wscript.echo "String used for comparison ="& vbCrLf&" " &OUImIn & " !"& vbCrLf&" Should only run once!"

		objRecordSet.MoveNext
	Loop
End If

arrPath = Split(OUImIn, ",")
		intLength = Len(arrPath(depth))
		intNameLength = intLength - 3
		Get_AD_OU = Right(arrPath(depth), intNameLength)
		if debugflag=True then wscript.echo "OU Query:  Query for OU at depth "&depth&", is " & Get_AD_OU

End Function


' *****************************************************
' KABI gebruikers aan array toevoegen
sub kabitoevoegen(user,kabi,pass)
	'wscript.echo "Arrcount "&arrcount
	'wscript.echo UBound(outArray,1)
	ReDim Preserve outArray(3,arrcount)
	outArray(0,arrcount) = lcase(user)
	outArray(1,arrcount) = kabi
	outArray(2,arrcount) = pass
	arrcount = arrcount +1
End sub
' *****************************************************

' *****************************************************
Sub outputarray()
	'Alleen voor Development

	for j=0 to arrcount-1
		'wscript.echo "Arrayteller count "&arrcount-1
		wscript.echo "Row: "&j&" Account "& outArray(0,j) &" KabiNaam "& outArray(1,j) &" Kabipass "& outArray(2,j)
	next
End Sub
' *****************************************************

' *****************************************************
Sub CheckName(gebruikersaccount)
	Gevonden = False
	For z=0 to UBound(outArray,2)
		If outArray(0,z) = lcase(gebruikersaccount) THEN
			if debugflag=True then wscript.echo "Kabi User " & gebruikersaccount & "Gevonden"
			Call KoppelShare("O:","\\10.1.1.44\FileServerData",outArray(1,z),outArray(2,z),"10.1.1.44")
			Gevonden = True
		End If
	Next

	IF Gevonden = False THEN
		if debugflag=False then MsgBoxPopup "Uw gebruikersaccount is niet in de KABI account lijst gevonden!" & VbCrLf & "Neem contact op met de Servicedesk.", vbOKOnly + vbExclamation, "Probleem met "&strDriveLetter &" schijf koppeling!"
		if debugflag=True then wscript.echo "Uw gebruikersaccount is niet in de kabi lijst gevonden."
	End If
End Sub
' *****************************************************

' *****************************************************
Function Mapprinter(portname,printqueue)
	On error resume next
	Set WshNetwork = WScript.CreateObject("WScript.Network")
	Set WshShell = WScript.CreateObject("WScript.Shell")
		WshNetwork.RemovePrinterConnection portname, true, true 'Bestaande mappings verwijderen
		Err.Clear

	if debugflag=True then wscript.echo "Mapping Network Printer "&printqueue
	WshNetwork.AddPrinterConnection portname, printqueue
	    If Err.Number <> 0 Then
    	'wscript.echo "FOUT!"
		if debugflag=False then MsgBoxPopup "Er vond een fout plaats tijdens het koppelen van de printer." & VbCrLf & VbCrLf & "FoutCode: " & _
			Err.Number & VbCrLf & VbCrLf & "Beschrijving: " & VbCrLf & Err.Description & VbCrLf & VbCrLf & "Queue: "& printqueue & VbCrLf & VbCrLf & "Port: "& portname& VbCrLf & "Neem contact op met de Servicedesk.", vbOKOnly + vbCritical, "Probleem koppelen netwerkprinter"
		WshShell.LogEvent 1, "Er vond een fout plaats tijdens het koppelen van de printer." & VbCrLf & VbCrLf & "FoutCode: " & _
			Err.Number & VbCrLf & VbCrLf & "Beschrijving: " & VbCrLf & Err.Description & VbCrLf & VbCrLf & "Queue: "& printqueue & VbCrLf & VbCrLf & "Port: "& portname& VbCrLf & "Neem contact op met de Servicedesk."
		Else
		'MsgBoxPopup "Schijf "& strDriveLetter &" succesvol gekoppeld.", vbOKOnly, strDriveLetter &"-schijf koppeling"
		WshShell.LogEvent 0,  "Printer "& printqueue &" succesvol gekoppeld op "& portname &"."
		End If

	Err.Clear
End Function
' *****************************************************

' *****************************************************
' Return LDAP Data
Function GetDnsDomain()
Set objADSysInfo = CreateObject("ADSystemInfo")
Set objRootLDAP = GetObject("GC://" & EnvString("userdnsdomain") & "/RootDSE")
strDNSDomain = objRootLDAP.Get("DefaultNamingContext")
GetDNSDomain = StrDNSDomain
End Function
' *****************************************************

Function AddDomainSites
					Dim objConnection, objRootDSE, objRecordSet
					Dim strFilter
					Dim IntraNetSites()
					arrSize=-1
	if debugflag=True then wscript.echo "TRUST: Starting enumeration of TRUSTS and add this to the Trusted Site of IE"
			'*********************************************************************************
			' Enumerate Domains in Forest and add to Array
	if debugflag=False then On error resume next
					strFilter = "(NETBIOSName=*)"

					Set objConnection = CreateObject("ADODB.Connection")
					objConnection.Provider = "ADsDSOObject"
					objConnection.Open "Active Directory Provider"

					Set objRootDSE = GetObject("GC://" & EnvString("userdnsdomain") & "/RootDSE")
					'Set objRootDSE = GetObject("GC://" & EnvString("userdnsdomain") & "/RootDSE")

					Set objRecordSet = objConnection.Execute( _
					   "<LDAP://" & EnvString("userdnsdomain") & "/" & objRootDSE.Get("configurationNamingContext") & ">;" & _
					   strFilter & ";" & "dnsroot,ncname;subtree")


					   set WshShell = WScript.CreateObject( "WScript.Shell" )
						WshShell.LogEvent 0 ,"User " &EnvString("username") & " queried GC:" & objRootDSE.Get("dnsHostName")
						set WshShell = Nothing

					   if debugflag=True then wscript.echo "TRUST: Connected to: "& objRootDSE.Get("dnsHostName")
					   'Set objRootDSE = Nothing

					   While Not objRecordSet.EOF
							if debugflag=True then WScript.Echo "TRUST: DNSRoot: " &Join(objRecordSet.Fields("dnsroot").Value)
							arrSize = arrSize + 1
							ReDim Preserve IntraNetSites(arrSize)
							IntraNetSites(arrSize) = Join(objRecordSet.Fields("dnsroot").Value)
							if debugflag=True then WScript.Echo "TRUST: Root: " & objRecordSet.Fields("ncname").Value
							objRecordSet.MoveNext
						WEnd

			'********************************************************************************
			' Enumerate Trusts and add to Array

					'set objRoot = getobject("LDAP://" & EnvString("userdnsdomain") & "/RootDSE")
					set objRoot = objRootDSE

					defaultNC = objRoot.get("defaultNamingContext")
					if debugflag=True then wscript.echo "TRUST: Connected to  (for enumeration): "& objRoot.Get("dnsHostName")
					set cn = createobject("ADODB.Connection")
					set cmd = createobject("ADODB.Command")
					set rs = createobject("ADODB.Recordset")

					cn.open "Provider=ADsDSOObject;"
					cmd.activeconnection =cn

					cmd.commandtext = "SELECT trustPartner,trustDirection, TrustType, flatName FROM 'GC://" & EnvString("userdnsdomain") & "/CN=System," & DefaultNC & "' WHERE objectclass = 'trusteddomain'"
					if debugflag=True then wscript.echo "TRUST: "&cmd.commandtext

					set rs = cmd.execute

					while rs.eof <> true and rs.bof <> true
						select case rs("trustDirection")
							case 0
								TrustDirection = "Disabled"
							case 1
								TrustDirection = "Inbound trust"
							case 2
								TrustDirection = "Outbound trust"
							case 3
								TrustDirection = "Two-way trust"
						end select
						select case rs("trustType")
							case 1
								TrustType = "Downlevel Trust"
							case 2
								TrustType = "Windows 2000 (Uplevel) Trust"
							case 3
								TrustType = "MIT"
							case 4
								TrustType = "DCE"
						end select
						if debugflag=True then wscript.echo "TRUST: DNS DomainName: " & rs("trustPartner") & " - Netbios: " & rs("flatName") & ", " & TrustType & ", " & TrustDirection
								arrSize = arrSize + 1
							ReDim Preserve IntraNetSites(arrSize)
							IntraNetSites(arrSize) = rs("trustPartner")
						rs.movenext
					wend

					cn.close
					'wscript.echo "All Sites"
					if debugflag=True then wscript.echo "TRUST: End of trust enumeration"

					'Adding NON trust DNS Domain Names
					if debugflag=True then wscript.echo "IE-INTRANET: NonDomain Sites: Adding NON trust DNS Domain Names"
					For Each AddSite in AdditionalSites
						arrSize = arrSize + 1
						ReDim Preserve IntraNetSites(arrSize)
						if debugflag=True then wscript.echo "IE-INTRANET: NonDomain Sites: Adding *." & AddSite
						IntraNetSites(arrSize) = AddSite
					Next


			' *******************************************************************************
			' Read array and add to local intranet zones
	if debugflag=True then wscript.echo "IE-INTRANET: Read TRUST array and add to local intranet zones (also ESC)"
			Const HKEY_CURRENT_USER = &H80000001

			strComputer = "."
			Set objReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

					For Each Site in IntraNetSites
						'wscript.echo Site
						strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\" & Site
						strKeyPathEnhanced = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains\" & Site
							if debugflag=True then wscript.echo "Site: *." & Site
						objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
						objReg.CreateKey HKEY_CURRENT_USER, strKeyPathEnhanced
						'strValueName = "http"
						strValueName = "*"
						dwValue = 1
						objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, strValueName, dwValue
						objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPathEnhanced, strValueName, dwValue
					Next
if debugflag=True then wscript.echo "IE-INTRANET: End of trust enumeration and Lokal intranet zone processing"

End Function

Sub Run(ByVal sFile)
Dim shell

    Set shell = CreateObject("WScript.Shell")
    shell.Run sFile & Chr(34), 1, false
    Set shell = Nothing
End Sub


' *****************************************************
' This function checks if user is local admin.
' On Win7 this shows no network drives because the
' GPO Logon script is run elevated.
Function KB937624()

If (CheckForServer = True) then
	Set oShell = CreateObject("WScript.Shell")
	oShell.LogEvent 4, "Computer is a server. I will not change EnableLinkedConnections"
	Exit Function
End If


		Dim oShell, oExec, szStdOut
		szStdOut = ""
		Set oShell = CreateObject("WScript.Shell")
		Set oExec = oShell.Exec("whoami /groups")

		Do While (oExec.Status = cnWshRunning)
		   WScript.Sleep 100
		   if not oExec.StdOut.AtEndOfStream then
			   szStdOut = szStdOut & oExec.StdOut.ReadAll
		   end if
		Loop

		select case oExec.ExitCode
		   case 0
			   if not oExec.StdOut.AtEndOfStream then
				   szStdOut = szStdOut & oExec.StdOut.ReadAll
			   end if

			   if ((instr(szStdOut,"S-1-16-12288")) OR (instr(szStdOut,"S-1-16-8448"))) Then

				   const HKEY_LOCAL_MACHINE = &H80000002
					strComputer = "."
					Set StdOut = WScript.StdOut

					Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
					strComputer & "\root\default:StdRegProv")

					strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
					strValueName = "EnableLinkedConnections"
					dwValue = 1

					oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
					If IsNull(strValue) Then
						'Wscript.Echo "The registry key does not exist."
						oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
						oShell.LogEvent 2, "User is Local Admin (S-1-16-12288) or PowerUser (S-1-16-8448), so changing registry for EnableLinkedConnections"
					Else
						'Wscript.Echo "The registry key exists."
						oShell.LogEvent 4, "User is Local Admin (S-1-16-12288) or PowerUser (S-1-16-8448), but EnableLinkedConnections already set."
					End If

			   else
				   if instr(szStdOut,"S-1-16-8192")  Then
					   'wscript.echo "Not Elevated"
					   oShell.LogEvent 4, "User is not Local Admin (S-1-16-8192), no need for EnableLinkedConnections."
				   else
					   oShell.LogEvent 4, "User is not Local Admin or no UAC enabled OS, no need for EnableLinkedConnections."
					   oShell.LogEvent 4, szStdOut
				   end if
			   end if

				if debugflag=True then wscript.echo "whoami output " &szStdOut

		   case else
			   if not oExec.StdErr.AtEndOfStream then
				   'wscript.echo oExec.StdErr.ReadAll
			   end if
		end select

End Function
' ****************************************************

' *****************************************************
' This function checks if this computer is a server.
Function CheckForServer()
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colOperatingSystems = objWMIService.ExecQuery _
		("Select * from Win32_OperatingSystem")

	For Each objOperatingSystem in colOperatingSystems
		'Wscript.Echo objOperatingSystem.Caption
		If (inStr(lcase(objOperatingSystem.Caption),"server")) then
			CheckForServer = CBool(True)
			Else
			CheckForServer = CBool(False)
		End If

	Next

End Function
' *****************************************************

' *****************************************************
' This function controls TMG Client Settings.
Function SetTMGClient(Enable,Browserconfig)
On Error Resume Next
'Params:	TMGClient,Browserconfig
'			Enable TMG Client:		Enable or Disable
'			Enable Browser Config:	EnableBrowserConfig or DisableBrowserconfig
if debugflag=True then wscript.echo "Setting TMG Client to " &Enable & " and "& BrowserConfig

Dim myProgramFiles
Dim myProgramFilesx86
Dim myPath
Dim wshShell
dim filesys
dim Command
Set wshShell = CreateObject( "WScript.Shell" )
Set filesys = CreateObject("Scripting.FileSystemObject")
myProgramFiles = wshShell.ExpandEnvironmentStrings( "%ProgramFiles%" )
myProgramFilesx86 = wshShell.ExpandEnvironmentStrings( "%ProgramFiles(x86)%" )
'wshShell = Nothing
'filesys = Nothing
	If filesys.FileExists(myProgramFiles & "\Forefront TMG Client\FwcTool.exe") Then
		'Execute
		wshShell.Run("%comspec% /C " & CHR(34) & myProgramFiles & "\Forefront TMG Client\FwcTool.exe"& CHR(34) &" "& Enable)
		wshShell.Run("%comspec% /C " & CHR(34) & myProgramFiles & "\Forefront TMG Client\FwcTool.exe"& CHR(34) &" "& BrowserConfig)
		WshShell.LogEvent 0, "TMG Client Config Tool has been run at "& myProgramFiles & "\Forefront TMG Client\FwcTool.exe"
	ElseIf filesys.FileExists(myProgramFilesx86 & "\Forefront TMG Client\FwcTool.exe") Then
		'Execute
		'wshShell.Run("CMD /K " & CHR(34) & myProgramFilesx86 & "\Forefront TMG Client\FwcTool.exe " & Enable & CHR(34))
		wshShell.Run("%comspec% /C " & CHR(34) & myProgramFilesx86 & "\Forefront TMG Client\FwcTool.exe"& CHR(34) &" "& Enable)
		wshShell.Run("%comspec% /C " & CHR(34) & myProgramFilesx86 & "\Forefront TMG Client\FwcTool.exe"& CHR(34) &" "& BrowserConfig)
		WshShell.LogEvent 0, "TMG Client Config Tool has been run at "& myProgramFilesx86 & "\Forefront TMG Client\FwcTool.exe"
	Else
		'Report no TMG Client Found
		if debugflag=True then wscript.echo "No TMG Client Config Tool Found"
		WshShell.LogEvent 1, "No TMG Client Config Tool Found at "& myProgramFilesx86 & "\Forefront TMG Client\FwcTool.exe or at "& myProgramFiles & "\Forefront TMG Client\FwcTool.exe"
	End If
End Function

Function MsgBoxPopup(Text,Buttons,Title)
	On error Resume Next
	if debugflag=True then wscript.echo "MsgPopup: " & Title
	'Value Button
	'0 OK
	'1 OK, Cancel
	'2 Abort, Ignore, Retry
	'3 Yes, No, Cancel
	'4 Yes, No
	'5 Retry, Cancel

	'Value Icon
	'16 Critical
	'32 Question
	'48 Exclamation
	'64 Information

	Set ButtonShell = CreateObject("WScript.Shell")
	if Buttons = 64 OR Buttons = 4160 then
			intButton = ButtonShell.Popup (Text, 30, Title, Buttons)
		else
			intButton = ButtonShell.Popup (Text, 60, Title, Buttons)
	End If
	MsgBoxPopup = intButton
	if debugflag=True then wscript.echo "MsgPopup: result= " &MsgBoxPopup
	Set ButtonShell=Nothing
End Function


' *****************************************************
' This will set Environment Variable
Function SetEnvironmentVariable(envvar,envval)

On error resume next

  Err.Clear

Set wshNetwork=CreateObject("WScript.Network")
  Set WSHShell = WScript.CreateObject("WScript.Shell")
  Set WshSystemEnv = WshShell.Environment("USER")

	WshShell.LogEvent 0 ,"Setting up Environment Variable " & envvar



	if Replace(WshNetwork.username," ", "") = "" then
		MsgBox "Error in setting Environment Variables! Username Issue" & VbCrLf , vbOKOnly + vbCritical, "Environment variable error!"
		WshShell.LogEvent 1 ,"Error in setting Environment Variables! Username Issue. "& Err.Description
	End if

'  WshSystemEnv("ASML_PLM_DIR") = "C:\ptc\workspaces\"& Replace(WshNetwork.username," ", "")
'  WshSystemEnv("ASML_PLM_DIR") = "\\vim.local\vim\Data\Config_CAD_PDM\ASML\plm_prd"
  WshSystemEnv(envvar) = envval

    If Err.Number <> 0 Then
    	'wscript.echo "FOUT!"
		WshShell.LogEvent 1 ,"Error in setting Environment Variables! "& Err.Description
		MsgBox "Error in setting Environment Variables!" & VbCrLf& VbCrLf & _
		"Please log off and log in again to correct this."& VbCrLf & VbCrLf & "ErrorCode: " & _
		Err.Number & VbCrLf & VbCrLf & "Description: " & VbCrLf & Err.Description, vbOKOnly + vbCritical, "Environment Variable error!"
	End If
	Err.clear
End Function
' *****************************************************

' *****************************************************
' Function to check ping latency
Function Ping()
	'strHost = "200.200.200.222"
	'strhost = EnvString("userdnsdomain")
	'Set objDomain = GetObject("LDAP://rootDse")
	Set objDomain = GetObject("GC://" & EnvString("userdnsdomain") & "/RootDSE")
	strHost = objDomain.Get("dnsHostName")
	IF IsNull(strHost) then
		strhost = EnvString("userdnsdomain")
		if debugflag=True then wscript.echo "Latency: Error finding GC, using %USERDNSDOMAIN%"
	ELSE if debugflag=True then wscript.echo "Latency: Found GC "&strHost
	End If


	Dim oPing, oRetStatus, bReturn
    Set oPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address='" & strHost & "'")

	if debugflag=True then wscript.echo "Latency: Pinging "&strHost

    For Each oRetStatus In oPing
        If IsNull(oRetStatus.StatusCode) Or oRetStatus.StatusCode <> 0 Then
            bReturn = -1

             if debugflag=True then WScript.Echo "Latency: Status code is " & oRetStatus.StatusCode
			 'if debugflag=True then WScript.Echo "Status code description is " & GetPingStatusCode(oRetStatus.StatusCode)
        Else
            'bReturn = True

             'Wscript.Echo "Bytes = " & vbTab & oRetStatus.BufferSize
             if debugflag=True then Wscript.Echo "Latency: Time (ms) = " & vbTab & oRetStatus.ResponseTime
             'Wscript.Echo "TTL (s) = " & vbTab & oRetStatus.ResponseTimeToLive
			 bReturn = oRetStatus.ResponseTime
        End If
		if debugflag=True then WScript.Echo "Latency: Status code is " & oRetStatus.StatusCode
		if debugflag=True then WScript.Echo "Latency: Status code description is: " & GetPingStatusCode(oRetStatus.StatusCode)
        Set oRetStatus = Nothing
    Next
    Set oPing = Nothing
    Ping = bReturn
End Function

' ____________________________
Function GetPingStatusCode(intCode)

  Dim strStatus
  Select Case intCode
  case  0
    strStatus = "Success"
  case  11001
    strStatus = "Buffer Too Small"
  case  11002
    strStatus = "Destination Net Unreachable"
  case  11003
    strStatus = "Destination Host Unreachable"
  case  11004
    strStatus = "Destination Protocol Unreachable"
  case  11005
    strStatus = "Destination Port Unreachable"
  case  11006
    strStatus = "No Resources"
  case  11007
    strStatus = "Bad Option"
  case  11008
    strStatus = "Hardware Error"
  case  11009
    strStatus = "Packet Too Big"
  case  11010
    strStatus = "Request Timed Out"
  case  11011
    strStatus = "Bad Request"
  case  11012
    strStatus = "Bad Route"
  case  11013
    strStatus = "TimeToLive Expired Transit"
  case  11014
    strStatus = "TimeToLive Expired Reassembly"
  case  11015
    strStatus = "Parameter Problem"
  case  11016
    strStatus = "Source Quench"
  case  11017
    strStatus = "Option Too Big"
  case  11018
    strStatus = "Bad Destination"
  case  11032
    strStatus = "Negotiating IPSEC"
  case  11050
    strStatus = "General Failure"
  case Else
    strStatus = intCode & " - Unknown"
  End Select
  GetPingStatusCode = strStatus

End Function

' *****************************************************

' *****************************************************
' Functino to start Ping Latency Check
Function CheckLatency()

 Dim LatencyMax
 LatencyMax = 50

Set wshNetwork=CreateObject("WScript.Network")
  Set WSHShell = WScript.CreateObject("WScript.Shell")
  Set WshSystemEnv = WshShell.Environment("USER")

if debugflag=True then
	wscript.echo "Latency: Starting latency check"
	DebugLatencyMax = 150
	wscript.echo "Lowering latency to " & DebugLatencyMax & " ms (only for debugging: normally it would be " & LatencyMax & " ms.)"
	LatencyMax = DebugLatencyMax
End If
	PingReply = Ping()
	PingReply = Ping()
	PingReply = Ping()
	PingRuns=1
	If (Pingreply > LatencyMax AND NOT IsNull(PingReply)) then
		if debugflag=True then wscript.echo "Latency: Higher Latency! - Will Wait ("&Pingreply&"ms reply of "&EnvString("userdnsdomain")&")"
		WshShell.LogEvent 1 ,"Higher Latency! - Will Wait ("&Pingreply&"ms reply of "&EnvString("userdnsdomain")&")"

		Berichttekst = "Momenteel is uw netwerkverbinding traag."& VbCrLf & VbCrLf & "We zullen gedurende maximaal 1 minuut uw verbinding opnieuw testen voordat we verder gaan." &_
						VbCrLf & "Dit komt vaker voor indien een WiFi-, netwerk- of VPN-verbinding net is opgebouwd."
		if debugflag=False then MsgBoxPopup Berichttekst,vbOKOnly + vbCritical,"Slow link detected..."

		Do Until ((PingReply <50 AND NOT IsNull(PingReply)) OR PingRuns>25)
			wscript.sleep 2000
			PingReply = Ping()
			if debugflag=True then wscript.echo "Latency: Pingreply ="&PingReply
			WshShell.LogEvent 1 ,"High Latency! Run "&PingRuns&" - Will Wait ("&Pingreply&"ms reply of "&EnvString("userdnsdomain")&")"
			PingRuns=PingRuns+1
		 Loop
	Else
		if debugflag=True then Wscript.echo "Latency: Pingreply ="&PingReply & " Tried it "&PingRuns&" times..."
		WshShell.LogEvent 0 ,"Latency is good. ("&Pingreply&"ms reply of "&EnvString("userdnsdomain")&")"
	End If
End Function
' *****************************************************


' *****************************************************
' Function to set Homepage in IE
Function SetHomePage(strHomePage)
if debugflag=True then wscript.echo "IE-HOMEPAGE: Setting Homepage to "&strHomePage
	HKEY_CURRENT_USER = &H80000001
	'strHomePage = "https://portal.vdlnet.nl" 'change to desired URL <-
	strMachine = "."
	Set objReg = GetObject("winmgmts:\\" & strMachine & "\root\default:StdRegProv")'#get the proper object to open registries.
	KeyPath = "SOFTWARE\Microsoft\Internet Explorer\Main"
	objReg.CreateKey HKEY_CURRENT_USER, KeyPath
	ValueName = "Start Page"
	objReg.SetStringValue HKEY_CURRENT_USER, KeyPath, ValueName, strHomePage
if debugflag=True then wscript.echo "IE-HOMEPAGE: Ended"
End Function
' *****************************************************

' *****************************************************
'RDS Function to redirect Documents to user documents folder
Function MKLinkDocuments(Target)
	Err.Clear
	MKLinkFail = False
	Try = 0
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set objShell = WScript.CreateObject("WScript.shell")
	Set fso = CreateObject("Scripting.FileSystemObject")
	TargetFile = Target & "\desktop.ini"
	WshShell.LogEvent 0, "Making symbolic link to "&Target&" for Documents Folder under Userprofile"

	If (fso.FileExists(TargetFile)) Then
		if debugflag=True then Wscript.Echo "LINKDOCS: Found desktop.ini file at location "&Target
		Do
		'Deleting My Documents at USERPROFILE
			Try = Try + 1
			if debugflag=True then Wscript.Echo "LINKDOCS: Deleting My Documents at USERPROFILE"
			Result = objShell.run("cmd /C rmdir /S /Q ""%USERPROFILE%\Documents""",2, True)
			if debugflag=True then Wscript.Echo "LINKDOCS: Result "&Result
			IF Result > 0 then
				SMKLinkFail = True
				WshShell.LogEvent 1, "Error Deleting Documents Folder"
			End If
		Loop Until ((Result < 1) OR (Try > 3))
		If Result = 0 then WshShell.LogEvent 0, "Succeeded Removing Existing Documents link" else WshShell.LogEvent 2, "Failed Removing Existing Documents link"

		Try = 0
		Do
		'Making Symbolic Link to target
			Try = Try + 1
			if debugflag=True then Wscript.Echo "LINKDOCS: Making Symbolic Link to target " & Target
			if debugflag=True then Wscript.Echo "LINKDOCS: cmd /C mklink /D ""%USERPROFILE%\Documents"" """& Target &""""
			Result = objShell.run("cmd /C mklink /D ""%USERPROFILE%\Documents"" """& Target &"""",2, True)
			if debugflag=True then Wscript.Echo "LINKDOCS: Result "&Result
			IF Result > 0 then
				MKLinkFail = True
				WshShell.LogEvent 1, "Error Creating Link for Documents Folder to "&Target
			End If
		Loop Until ((Result < 1) OR (Try > 3))
		If Result = 0 then WshShell.LogEvent 0, "Succeeded Creating Documents Link" else WshShell.LogEvent 2, "Failed Creating Documents Link"
	Else
		WshShell.LogEvent 1, "Error Creating Link for Documents Folder to "&Target&". No desktop.ini found"
		if debugflag=True then Wscript.Echo "LINKDOCS: ERROR: No found desktop.ini file at location "&Target
		MKLinkFail = True
	End If

	If MKLinkFail = True then
		MsgBoxPopup "Er is een probleem ontstaan in de omleiding van uw Documenten Folder." & VbCrLf & "Tijdens deze sessie zult u gebruik moeten maken van de netwerkschijven als u nieuw opgeslagen documenten wilt behouden.", vbOKOnly + vbExclamation, "Fout in het omleiden van Uw Documenten Folder"
	End If

	Set objShell = Nothing
	Set WshShell = Nothing
	Set fso = Nothing
End Function
' *****************************************************

' *****************************************************
'RDS Function to redirect Desktop to user desktop folder
Function MKLinkDesktop(Target)
	Err.Clear
	MKLinkFail = False
	Try = 0
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set objShell = WScript.CreateObject("WScript.shell")
	Set fso = CreateObject("Scripting.FileSystemObject")
	TargetFile = Target & "\desktop.ini"

	WshShell.LogEvent 0, "Making symbolic link to "&Target&" for Desktop Folder under Userprofile"

	If (fso.FileExists(TargetFile)) Then
		if debugflag=True then Wscript.Echo "LINKDESK: Found desktop.ini file at location "&Target
		Do
		'Deleting Desktop at USERPROFILE
			Try = Try + 1
			if debugflag=True then Wscript.Echo "LINKDESK: Deleting Desktop at USERPROFILE"
			Result = objShell.run("cmd /C rmdir /S /Q ""%USERPROFILE%\Desktop""",2, True)
			if debugflag=True then Wscript.Echo "Result "&Result
			IF Result > 0 then
				MKLinkFail = True
				WshShell.LogEvent 1, "Error Deleting Desktop Folder"
			End If
		Loop Until ((Result < 1) OR (Try > 3))
		If Result = 0 then WshShell.LogEvent 0, "Succeeded Removing Existing Desktop Link" else WshShell.LogEvent 2, "Failed Removing Existing Desktop Link"

		Try = 0
		Do
		'Making Symbolic Link to target
			Try = Try + 1
			if debugflag=True then Wscript.Echo "LINKDESK: Making Symbolic Link to target " & Target
			if debugflag=True then Wscript.Echo "LINKDESK: cmd /C mklink /D ""%USERPROFILE%\Desktop"" """& Target &""""
			Result = objShell.run("cmd /C mklink /D ""%USERPROFILE%\Desktop"" """& Target &"""",2, True)
			if debugflag=True then Wscript.Echo "LINKDESK: Result "&Result
			IF Result > 0 then
				MKLinkFail = True
				WshShell.LogEvent 1, "Error Creating Link for Desktop Folder to "&Target
			End If
		Loop Until ((Result < 1) OR (Try > 3))
		If Result = 0 then WshShell.LogEvent 0, "Succeeded Creating Desktop Link" else WshShell.LogEvent 2, "Failed Creating Desktop Link"
	Else
		WshShell.LogEvent 1, "Error Creating Link for Desktop Folder to "&Target&". No desktop.ini found"
		if debugflag=True then Wscript.Echo "LINKDESK: ERROR: No found desktop.ini file at location "&Target
		MKLinkFail = True
	End If

	If MKLinkFail = True then
		MsgBoxPopup "Er is een probleem ontstaan in het omleiding van uw Bureaublad Folder." & VbCrLf & "Tijdens deze sessie zult u gebruik moeten maken van de netwerkschijven als u nieuw opgeslagen documenten wilt behouden.", vbOKOnly + vbExclamation, "Fout in het omleiden van Uw Bureaublad Folder"
	End If

	Set objShell = Nothing
	Set WshShell = Nothing
End Function
' *****************************************************

' *****************************************************
'Change default network printer
Function SetNetPrinterDefault(strUNCPrinter)
Dim WSHNetwork
Dim WsShell
Dim objWMIService
Dim colPrinters
Dim objPrinter

if debugflag=True then WScript.Echo "Querying printer list, searching for "&strUNCPrinter
Set WSHNetwork = CreateObject("WScript.Network")
Set WshShell = WScript.CreateObject("WSCript.shell")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")

For Each objPrinter In colPrinters
    If InStr(lcase(objPrinter.Name), lcase(strUNCPrinter)) Then
        WSHNetwork.SetDefaultPrinter objPrinter.Name
		if debugflag=True then WScript.Echo "Asked for default printer to be "&strUNCPrinter&". Therefore default printer has been set to " & objPrinter.Name
		WshShell.LogEvent 0, "Asked for default printer to be "&strUNCPrinter&". Therefore default printer has been set to " & objPrinter.Name
        Exit For
    End If
Next

Set WSHNetwork = nothing
Set objWMIService = nothing
Set colPrinters = nothing
Set WshShell = nothing
End Function
' *****************************************************

' *****************************************************
'Change default local printer
Function SetLocalPrinterDefault(strUNCPrinter)
Dim WSHNetwork
Dim WshShell
Dim objWMIService
Dim colPrinters
Dim objPrinter

if debugflag=True then WScript.Echo "Querying printer list, searching for "&strUNCPrinter
Set WSHNetwork = CreateObject("WScript.Network")
Set WshShell = WScript.CreateObject("WSCript.shell")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")

For Each objPrinter In colPrinters
    If InStr(lcase(objPrinter.Name), lcase(strUNCPrinter)) Then
        'WSHNetwork.SetDefaultPrinter objPrinter.Name
		WshShell.run "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n "& Chr(34) & objPrinter.Name & Chr(34)
		if debugflag=True then WScript.Echo "Asked for default printer to be "&strUNCPrinter&". Therefore default printer has been set to " & objPrinter.Name
		WshShell.LogEvent 0, "Asked for default printer to be "&strUNCPrinter&". Therefore default printer has been set to " & objPrinter.Name
        Exit For
    End If
Next

Set WSHNetwork = nothing
Set objWMIService = nothing
Set colPrinters = nothing
Set WshShell = nothing
End Function
' *****************************************************

' *****************************************************
' Get OU My computer account belongs to
Function GetComputerAccountOU(strDepth)
Dim WshNetwork
Dim ComputerName
Dim objADSysInfo
Dim strComputerName
Dim objComputer
Dim strOUName
Dim strOUs
Dim StrOU
Set WshNetwork = CreateObject("WScript.Network")
ComputerName = WshNetwork.ComputerName

Set objADSysInfo = CreateObject("ADSystemInfo")
strComputerName = objADSysInfo.ComputerName

Set objComputer = GetObject("GC://" & strComputerName)
strOUName = objComputer.DistinguishedName
strOUs = Split(strOUName, ",")
strOU = Split(strOUs(strDepth), "=")
GetComputerAccountOU = strOU(1)
if debugflag=True then wscript.echo "COMP OU Query:  Query for OU at depth "&strdepth&", name: " & GetComputerAccountOU
End Function
' *****************************************************

' *****************************************************
'Hide all redirected drives except C and D
'Const HKEY_CLASSES_ROOT = &H80000000
'Const HKEY_CURRENT_USER = &H80000001
'Const HKEY_LOCAL_MACHINE = &H80000002
'Const HKEY_USERS = &H80000003
'Const HKEY_CURRENT_CONFIG = &H80000005


Function HideRedirectedDrivesFromExplorer()
   Dim aSessions, sSession, sNameSpacesPath, aNamespaces, sNameSpace, sClientDriveLetterPath, sClientDriveLetterValue, sClientDriveLetter, sNameSpacePath

if debugflag=True then WScript.Echo "HIDEREDIR: Search For redirected drives and hiding them/"

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005

Const SESSIONINFO_REG_PATH = "Software\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo"


   'Returns an array of sessions that the user has on this Terminal Server
   'Users are restricted to one session only, meaning a single subkey should be returned when ReturnSubKeys() is called
   'The session restriction is configured under Terminal Services Configuration-->Server Settings
   aSessions = ReturnSubKeys(HKEY_CURRENT_USER, SESSIONINFO_REG_PATH)
   If IsArray(aSessions) Then
      For each sSession In aSessions
         sNamespacesPath = SESSIONINFO_REG_PATH & "\" & sSession & "\MyComputer\Namespace"
         if debugflag=True then WScript.Echo "HIDEREDIR: working "&sNamespacesPath
         'Returns an array of namespaces, each represented by a GUID
         'There is a Namespace for each client drive letter
         aNamespaces = ReturnSubKeys(HKEY_CURRENT_USER, sNamespacesPath)

         If IsArray(aNamespaces) Then

            'Loop through each namespace and look up the GUID under HKCU\Software\Classes\CLSID
            'The GUID key contains client drive info, and if the Target drive letter
            'is N, U, or any other drive letter in the Case statement, delete the namespace
            For Each sNamespace In aNamespaces
               sClientDriveLetterPath =  "Software\Classes\CLSID\" & sNameSpace & "\Instance\InitPropertyBag"
			   sClientDriveLetterPathWOW64 =  "Software\Classes\Wow64Node\CLSID\" & sNameSpace & "\Instance\InitPropertyBag"
			   if debugflag=True then WScript.Echo "HIDEREDIR: working "&sNamespace
			   if debugflag=True then WScript.Echo "HIDEREDIR: working "&sClientDriveLetterPath
               sClientDriveLetterValue = "Target"
               sClientDriveLetter = Ucase(Right(ReturnStringValue(HKEY_CURRENT_USER, sClientDriveLetterPath, sClientDriveLetterValue), 1))
			   if isNull(sClientDriveLetter) then sClientDriveLetter = Ucase(Right(ReturnStringValue(HKEY_CURRENT_USER, sClientDriveLetterPathWOW64, sClientDriveLetterValue), 1))
			   if isNull(sClientDriveLetter) then sClientDriveLetter = Ucase(Right(ReturnStringValue(HKEY_LOCAL_MACHINE, sClientDriveLetterPath, sClientDriveLetterValue), 1))
			   if isNull(sClientDriveLetter) then sClientDriveLetter = Ucase(Right(ReturnStringValue(HKEY_LOCAL_MACHINE, sClientDriveLetterPathWOW64, sClientDriveLetterValue), 1))
               if debugflag=True then WScript.Echo "HIDEREDIR: working drive "&sClientDriveLetter
			   sNamespacePath = sNamespacesPath & "\" & sNameSpace

			  If (sClientDriveLetter <> "C" and sClientDriveLetter <> "D") then
				if debugflag=True then WScript.Echo "HIDEREDIR: Disconnecting client drive "& sClientDriveLetter &" " & VbCrLf & sNamespacePath
				call DeleteKey(HKEY_CURRENT_USER, sNamespacePath)
			  End If

            Next
         End If
      Next
   End If

if debugflag=True then WScript.Echo "HIDEREDIR: End"

End Function

Sub DeleteKey(sRegHive, sRegPath)
   Dim sComputer, oReg
   sComputer = "."
   Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& sComputer & "\root\default:StdRegProv")
   oReg.DeleteKey sRegHive, sRegPath
   set oReg = nothing
End Sub

Function ReturnStringValue(sRegHive, sRegPath, sRegValue)
   Dim sComputer, oReg, sStringValue
   sComputer = "."
   Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& sComputer & "\root\default:StdRegProv")
   oReg.GetStringValue sRegHive, sRegPath, sRegValue, sStringValue
   if debugflag=True then WScript.Echo "HIDEREDIR: trying to fetch "&sRegHive& sRegPath& sRegValue
      set oReg = nothing
   ReturnStringValue = sStringValue
End Function

Function ReturnSubKeys(sRegHive, sRegPath)
   Dim aSubKeys, sSubKey, sComputer, oReg
   sComputer = "."
   Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& sComputer & "\root\default:StdRegProv")
   oReg.EnumKey sRegHive, sRegPath, aSubKeys
   set oReg = nothing
   ReturnSubKeys = aSubKeys
End Function
' *****************************************************

' *****************************************************
'rename mapped drive
Function RenameMappedDrive(strDrive,strNewName)

Set oShell = CreateObject("Shell.Application")
'oShell.NameSpace("U:\").Self.Name = "Home Drive"
oShell.NameSpace(strDrive).Self.Name = strNewName

End Function
' *****************************************************


' *****************************************************
' Prevent duplicate Runs of the same script
Function CheckDuplicateScriptInstance
	'Functionality supported only for winXP up
	if debugflag=True then WScript.Echo "INSTANCECHECK: Started"
	dim svc, squery, ncount,mycount,strUser,WshShell,CMDLine
	Set WshShell = WScript.CreateObject("WScript.Shell")
	'set svc=getobject("winmgmts:root\cimv2")
	strComputer = "."
	strUser = CreateObject("WScript.Network").UserName

	squery="select commandline,name from win32_process " & _
		"where (not ((commandline like ""%command.com%"") or (commandline like ""%cmd.exe%""))) and " & _
		"commandline like '%[WC]script%" & wscript.ScriptName & "%' and " & _
		"NOT commandline like '%[WC]script%/Loginscript%" & wscript.ScriptName & "%'"

	set colProcesses = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2").ExecQuery(squery)

	'set colProcesses = svc
	ncount=colProcesses.count
	'wscript.echo ncount

	mycount=ncount
	if debugflag=True then WScript.Echo "INSTANCECHECK: Instance Count: "&myCount

	For Each objProcess in colProcesses

		Return = objProcess.GetOwner(strNameOfUser)
		If Return <> 0 Then
			Wscript.Echo "Could not get owner info for process " & _
				objProcess.Name & VBNewLine _
				& "Error = " & Return
		Else
			If (strUser <> strNameOfUser) then
			'Wscript.Echo "Process " _
			'    & objProcess.Name & " is owned by " _
			'    & "\" & strNameOfUser & "."
			myCount = myCount-1
			if debugflag=True then WScript.Echo "INSTANCECHECK: Instance Count: "&myCount
			else
			CMDLine = objProcess.Commandline
			if debugflag=True then WScript.Echo "INSTANCECHECK: Instance found: "&CMDLine
			End If
		End If
	Next


	'set svc=nothing
	set colProcesses = nothing

	if mycount>1 then
	'wscript.echo("only one instance allowed")
	WshShell.LogEvent 2, "Login script is already running for user "& strUser &"!" & VbCrLf & "Script: " & wscript.ScriptFullname & VbCrLf & "Commandline: "& CMDLine & _
						VbCrLf & VbCrLf & "This instance will quit"
	if debugflag=True then WScript.Echo "INSTANCECHECK: Instance found with CMDLine: "&CMDLine
	if debugflag=True then WScript.Echo "INSTANCECHECK: Instance found! Quitting now!"
		wscript.quit
	else
	'Wscript.Echo "Script " & wscript.ScriptName & " started."
	'Do nothing and continue
	end if

	set WshShell=nothing
End Function
' *****************************************************


' *****************************************************
' Returns an array of printers located on a server
' Used in function MapSharedClientPrinter to map a shared printer on the RD Client that a user uses.
function ListPrinters(strServer)

	dim objShell, objExecObject
	dim strCommand, strResults
	dim arrResults
	dim arrPrinterResults : arrPrinterResults = Array()

	Set objShell = CreateObject("WScript.Shell")

	' Get list of printers
	strCommand = "net view \\" & strServer
	strResults=""
	Set objExecObject = objShell.Exec(strCommand)
	Do
		WScript.Sleep 100
	Loop Until objExecObject.Status <> 0
	strResults = objExecObject.StdOut.ReadAll()

	' Now parse list for printers
	Dim i
	arrResults=Split(strResults, vbCrLf)
	strResults=""
	for i=0 to UBound(arrResults)
		'if debugflag=True then wscript.echo "MAPSHAREDCLIENTPRINTER: i: "& i & " arrResult: "& arrResults(i) & " uBound(arrResults):" & UBound(arrResults)
		if Instr(1,arrResults(i),"Print")>0 then
			if debugflag=True then wscript.echo "MAPSHAREDCLIENTPRINTER: Adding " & Trim(Left(arrResults(i), InStr(1,arrResults(i),"Print")-1)) & " to Array"
			strResults=strResults & Trim(Left(arrResults(i), InStr(1,arrResults(i),"Print")-1)) & vbCrLf
	  end if
	next
	if (strResults <> "") then
		strResults=Left(strResults, Len(strResults)-2)
		arrPrinterResults=Split(strResults, vbCrLf)
		'if debugflag=True then wscript.echo "MAPSHAREDCLIENTPRINTER: returning Array"
	else
		arrResults = Array()
	end If

	ListPrinters=arrPrinterResults
end function
' *****************************************************

' *****************************************************
' Map a shared printer on the RDS Client Computer
' That matches the name %strPrinterName%
Function MapSharedClientPrinter(strPrinterName)
	Dim WshShell
	Dim RemoteClient
	Set WshShell = WScript.CreateObject("WScript.Shell")
	RemoteClient = EnvString("CLIENTNAME")
	if debugflag=True then wscript.echo "MAPSHAREDCLIENTPRINTER: Detected variable CLIENTNAME to be: "&RemoteClient
	If (RemoteClient <> "" and RemoteClient <> "%CLIENTNAME%") then
		if debugflag=True then wscript.echo "MAPSHAREDCLIENTPRINTER: This means we are running in a RD Session"
		WshShell.LogEvent 0, "MAPSHAREDCLIENTPRINTER: RD Session detected using client "&RemoteClient
		Dim strPrinterUNC, objNetwork,i
		Dim arrPrinters : arrPrinters = Array()
		Set objNetwork = CreateObject("WScript.Network")
		arrPrinters=ListPrinters(RemoteClient)
		if debugflag=True then wscript.echo "MAPSHAREDCLIENTPRINTER: "&UBound(arrPrinters)+1 & " shared printers found on client used on "&Remoteclient
		for i=0 to UBound(arrPrinters)
			if Instr( 1, ucase(arrPrinters(i)), ucase(strPrinterName), vbTextCompare ) > 0 then
				if debugflag=True then wscript.echo "MAPSHAREDCLIENTPRINTER: Printer " & arrPrinters(i) & " matches search query of "& strPrinterName
				strPrinterUNC = "\\" & remoteclient & "\" & arrPrinters(i)
				objNetwork.AddWindowsPrinterConnection strPrinterUNC
				If Err <> 0 then
					if debugflag=True then wscript.echo "MAPSHAREDCLIENTPRINTER: ERROR Mapping "&strPrinterName&" Printer at "& strPrinterUNC
					WshShell.LogEvent 3, "MAPSHAREDCLIENTPRINTER: Error "&Err.Numer&" Mapping "& strPrinterUNC & ". Description: "&Err.Description
				Else
					if debugflag=True then wscript.echo "MAPSHAREDCLIENTPRINTER: Found and mapped "&strPrinterName&" Printer at "& strPrinterUNC
					WshShell.LogEvent 0, "MAPSHAREDCLIENTPRINTER: Mapped Printer at "& strPrinterUNC & "."
				End If
			Else
				if debugflag=True then wscript.echo "MAPSHAREDCLIENTPRINTER: Printer " & arrPrinters(i) & " not a match for "& strPrinterName
			end if
		next
		set objNetwork = nothing
	Else
	if debugflag=True then wscript.echo "MAPSHAREDCLIENTPRINTER: This means we are NOT running in a RD Session"
	'WshShell.LogEvent 0, "MAPSHAREDCLIENTPRINTER: RD Session detected using client "&RemoteClient
	End If
End Function
' *****************************************************

' *****************************************************
' Function to remove any printer in the user session that matches %strPrinterName%
Function RemoveMappedPrinter(strPrinterName)
	if debugflag=False then On Error Resume Next
	Dim strComputer
	Dim colPrinters
	Dim objWMIService
	Dim objPrinter
	strComputer = "."
	 Set objWMIService = GetObject("winmgmts:" _
	 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	 Set colPrinters =  objWMIService.ExecQuery _
	 ("Select * from Win32_Printer Where DeviceID like '%"&strPrinterName&"%'")

	 For Each objPrinter in colPrinters
	 if debugflag=True then Wscript.Echo "Removing: "&objPrinter.Name
	 objPrinter.Delete_
	 Next
End Function
' *****************************************************
