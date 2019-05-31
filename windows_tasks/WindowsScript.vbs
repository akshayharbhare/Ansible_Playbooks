Option Explicit
Const strScriptNameVer = "WinAudit v1.4.3"

Dim objFSO, objEnv, objShell, objExecObject
Dim strServerName, strClientName, strDirName, strDomainRole, strDomainName
Dim strSystemInfo, strUsers, strGroups, strFilePerms, strDirPerms, _
	strRegistryVals, strServices, strHotFixes, strLogSettings, strShares, _
	strRegistryPerms, strDrives, strADTrusts, strErrorLog, strOSVersion, strGpresult

strServerName = CreateObject("WScript.Network").ComputerName

'WScript.StdOut.WriteLine "Welcome to " & strScriptNameVer & vbCRLF
'WScript.StdOut.WriteLine "Initializing Scan..."

' build the name of the output directory
Set objFSO = CreateObject("Scripting.FileSystemObject")
strDirName = strServername & "-" & FormatDateTime(NOW(),vbLongDate) &_
	" " & Replace(FormatDateTime(NOW(),vbLongTime),":",".")

' remove the directory if it already exists
If(objFSO.FolderExists(strDirName)) Then
	objFSO.DeleteFolder strDirName
End If

' Create the output directory and clear the file system object
objFSO.CreateFolder(strDirName)
Set objFSO = Nothing

' Generate the filenames for all of the output files
strSystemInfo = ".\" & strDirName & "\(" & strServerName & ")SystemInfo.xls"
strUsers = ".\" & strDirName & "\(" & strServerName & ")Users.xls"
strGroups = ".\" & strDirName & "\(" & strServerName & ")Groups.xls"
strFilePerms = ".\" & strDirName & "\(" & strServerName &_
	")FilePermissions.xls"
strDirPerms = ".\" & strDirName & "\(" & strServerName &_
	")DirectoryPermissions.xls"
strRegistryVals = ".\" & strDirName & "\(" & strServerName &_
	")RegistryValues.xls"
strRegistryPerms = ".\" & strDirName & "\(" & strServerName &_
	")RegistryPermissions.xls"
strServices = ".\" & strDirName & "\(" & strServerName & ")Services.xls"
strHotFixes = ".\" & strDirName & "\(" & strServerName & ")HotFixes.xls"
strLogSettings = ".\" & strDirName & "\(" & strServerName & ")LogSettings.xls"
strShares = ".\" & strDirName & "\(" & strServerName & ")Shares.xls"
strDrives = ".\" & strDirName & "\(" & strServerName & ")Drives.xls"
strADTrusts = ".\" & strDirName & "\(" & strServerName & ")ADTrusts.xls"
strErrorLog = ".\" & strDirName & "\(" & strServerName & ")ErrorLog.txt"
strGpresult = ".\" & strDirName & "\(" & strServerName & ")gpresult.txt"

' Write distribution limitation on the top of every output file
writeToFile strSystemInfo, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
writeToFile strUsers, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
writeToFile strGroups, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
writeToFile strFilePerms, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
writeToFile strDirPerms, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
writeToFile strRegistryVals, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
writeToFile strServices, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
writeToFile strHotFixes, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
writeToFile strLogSettings, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
writeToFile strShares, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
writeToFile strDrives, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
writeToFile strGpresult, strScriptNameVer & ": Security Assessment: " &_
	"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
	
'WScript.StdOut.WriteLine "Gathering system information..."

'Get  Date and Time of Script Execution
writeToFile strSystemInfo, strScriptNameVer & " Run on: " &_
	FormatDateTime(NOW(),vbGeneralDate) & vbCRLF & vbCRLF

'Get Operating System information (computer name, os version, service pack, etc)
writeToFile strSystemInfo, getOSVersion(strOSVersion)

'Get IP Address of host for each ethernet adapter
writeToFile strSystemInfo, getIPAddress()

'Get Domain Information
writeToFile strSystemInfo, getDomainInfo(strDomainRole, strDomainName)

'Get Current user id
writeToFile strSystemInfo, getCurrentUser()

Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("Process")

' Gpresult
' If the system  is not running Windows 2000, grab "gpresult /Z"
If(Left(strOSVersion,3) <> "5.0" And Left(strOSVersion,3) <> "4.0") Then
	'WScript.StdOut.WriteLine "Gathering gpresult information..."
	Set objExecObject = objShell.exec(objEnv.Item("SystemRoot") &_
		"\system32\gpresult.exe /Z")
	
	Do While Not objExecObject.StdOut.AtEndOfStream
		writeToFile strGpresult, objExecObject.StdOut.ReadLine() & vbCRLF
	Loop
Else
	writetoFile strGpresult, "System is running Windows 2000, no gpresult output gathered" & vbCRLF
End If

' File Permissions
'WScript.StdOut.WriteLine "Gathering file permission information..."
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\regedit.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\arp.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\at.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\attrib.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\cacls.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\cmd.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\dcpromo.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\debug.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\edit.com")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\edlin.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\eventtriggers.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\finger.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\ftp.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\gpupdate.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\ipconfig.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\nbtstat.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\net.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\net1.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\netstat.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\nslookup.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\ntbackup.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\ping.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\rcp.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\reg.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\regedt32.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\regsvr32.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\rexec.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\route.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\rsh.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\runonce.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\sc.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\secedit.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\syskey.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\systeminfo.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\telnet.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\tftp.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\tlntsvr.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\tlntsess.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\tlntadmn.exe")	
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\tlntsvrp.dll")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\tracert.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\xcopy.exe")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\config\APPEVENT.EVT")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\config\SECEVENT.EVT")
writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
	"\system32\config\SYSEVENT.EVT")
If(InStr(strDomainRole, "Domain Controller") > 0) Then
	writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
		"\system32\config\DNSEVENT.EVT")
	writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
		"\system32\config\NTDS.EVT")
	writeToFile strFilePerms, getFilePermissions(objEnv.Item("SystemRoot") &_
		"\system32\config\NTFRS.EVT")
End If

' Directory Permissions
'WScript.StdOut.WriteLine "Gathering directory permission information..."
writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemDrive") & "\")
writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemRoot"))
writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemRoot") &_
	"\repair")
writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemRoot") &_
	"\security")
writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemRoot") &_
	"\system32")
writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemRoot") &_
	"\system32\drivers")
writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemRoot") &_
	"\system32\config")
writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemRoot") &_
	"\system32\spool")
writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemRoot") &_
	"\system32\spool\printers")
	
If(InStr(strDomainRole, "Domain Controller") > 0) Then
	writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemRoot") &_
		"\SYSVOL")
	writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemRoot") &_
		"\ntds")
	writeToFile strDirPerms, getDirPermissions(objEnv.Item("SystemRoot") &_
		"\ntfrs")
End If

Set objShell = Nothing
Set objEnv = Nothing

' Registry Settings
'WScript.StdOut.WriteLine "Gathering registry settings..."
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_CURRENT_USER", _
	"Control Panel\Desktop")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\Driver Signing")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\OLE")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\OS/2 Subsystem for NT")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\OS/2 Subsystem for NT\1.0")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\OS/2 Subsystem for NT\1.0\config.sys")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\TelnetServer")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Setup\RecoveryConsole")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Microsoft\Windows NT\Terminal Services")	
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Policies\Microsoft\Cryptography")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Policies\Microsoft\Messenger\Client")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Policies\Microsoft\PCHealth\ErrorReporting\DW")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Policies\Microsoft\PCHealth\ErrorReporting")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Policies\Microsoft\Windows NT\DNSClient")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Policies\Microsoft\Windows NT\Network Connections")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Policies\Microsoft\Windows\Safer\CodeIdentifiers")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SOFTWARE\Policies\Microsoft\Windows\System")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\FileSYSTEM")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\LSA")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\LSA\MSV1_0")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\Print\Providers\LanMan Print Services\" &_
	"Servers")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\SecurePipeServers\winreg")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\SecurePipeServers\winreg\AllowedPaths")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\SecurePipeServers\winreg\" &_
	"AllowedExactPaths")	
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\Session Manager")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\Session Manager\Kernel")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\Session Manager\SubSystems")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\Terminal Server")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\LSA")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services\LanManServer\Parameters")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services\LDAP")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services\Netlogon\Parameters")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services\SimpTcp")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services\SimpTcp\Parameters")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services\SNMP")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services\SNMP\Parameters")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"SYSTEM\CurrentControlSet\Services\TlntSvr")
writeToFile strRegistryVals, enumerateRegistryValues("HKEY_USERS", _
	".DEFAULT\Control Panel\Desktop")

'DC Registry Keys
If(InStr(strDomainRole, "Domain Controller") > 0) Then
	writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
		"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System")
	writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
		"SYSTEM\CurrentControlSet\Control\LSA")
	writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
		"SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management")
	writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
		"SYSTEM\CurrentControlSet\Services\EventLog\Directory Service")
	writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
		"SYSTEM\CurrentControlSet\Services\EventLog\DNS Server")
	writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
		"SYSTEM\CurrentControlSet\Services\EventLog\File Replication Service")
	writeToFile strRegistryVals, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
		"SYSTEM\CurrentControlSet\Services\NTDS\Parameters")
End If

'Services & Status
'WScript.StdOut.WriteLine "Gathering services information..."
writeToFile strServices, listServices()
'
'Patches
'WScript.StdOut.WriteLine "Gathering hotfix information..."
writeToFile strHotFixes, listHotFixes()
writeToFile strHotFixes, vbCRLF

'Log Settings (Size of the logs)
'WScript.StdOut.WriteLine "Gathering log settings..."
writeToFile strLogSettings, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"System\CurrentControlSet\Services\EventLog\Security")
writeToFile strLogSettings, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"System\CurrentControlSet\Services\EventLog\Application")
writeToFile strLogSettings, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
	"System\CurrentControlSet\Services\EventLog\System")
If(InStr(strDomainRole, "Domain Controller") > 0) Then
	writeToFile strLogSettings, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
		"System\CurrentControlSet\Services\EventLog\Directory Service")
	writeToFile strLogSettings, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
		"System\CurrentControlSet\Services\EventLog\DNS Server")
	writeToFile strLogSettings, enumerateRegistryValues("HKEY_LOCAL_MACHINE", _
		"System\CurrentControlSet\Services\EventLog\File Replication Service")
End if

' Share Permissions
'WScript.StdOut.WriteLine "Gathering shares information..."
writeToFile strDirPerms, vbCRLF & "NTFS Permissions for Share Directories" & vbCRLF
writeToFile strShares, getShares(strDirPerms)

'Drive info
'WScript.StdOut.WriteLine "Gathering drive information..."
writeToFile strDrives, getDrives()

'Domain Controller Specific
If(InStr(strDomainRole, "Domain Controller") > 0) Then
	'Get user and groups for domain controller
	'WScript.StdOut.WriteLine "Gathering user information..."
	getADUserAccounts()
	'WScript.StdOut.WriteLine "Gathering group information..."
	getADGroups()
	' GPO - For domain controllers
	'WScript.StdOut.WriteLine "Gathering group policy information..."
	retrieveGPODirs(strDirName)
	'AD Trusts - For domain controllers
	'WScript.StdOut.WriteLine "Gathering Active Directory trust information..."
	writeToFile strADTrusts, strScriptNameVer & ": Security Assessment: " &_
		"Confidential for " &	strClientName & " use only" & vbCRLF & vbCRLF
	writeToFile strADTrusts, getADTrusts(strDomainName)
Else  'Not DC Specific
	' Get user and groups for non-domain controllers
	'WScript.StdOut.WriteLine "Gathering user information..."
	writeToFile strUsers, getUserAccounts()
	'WScript.StdOut.WriteLine "Gathering group information..."
	writeToFile strGroups, getGroups()
End If

' Audit Settings
'WScript.StdOut.WriteLine "Gathering audit settings..."
' Passwd Policies
'WScript.StdOut.WriteLine "Gathering password policies..."
' User Rights
'WScript.StdOut.WriteLine "Gathering user rights..."
analyzeAuditandUserRights(strDirName)

'MsgBox strScriptNameVer & " complete!", 0, strScriptNameVer


'######################################################
'# Functions
'######################################################

'######################################################
'# Function Name: getOSVersion
'# Parameters: 
'#	strOsVersion - Output parameter storing the OS version of the system
'# Return Value: 
'#	getOSVersion - String containing system name, caption, service pack level, and version
'#
'# Description:
'#	Queries the WMI service for the name of system, the caption (cointains general system information)
'#	the service pack level, and the numeric version of the system.
'######################################################
Function getOSVersion (ByRef strOsVersion)
	Dim strComputer, objWMIService, colOSes, objOS
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
	For Each objOS in colOSes
		getOSVersion = "Computer Name: " & objOS.CSName & vbCRLF
		getOSVersion = getOSVersion &  "Caption: " & objOS.Caption & vbCRLF 'Name
		getOSVersion = getOSVersion &  "Service Pack: " &_
			objOS.ServicePackMajorVersion & "." &	objOS.ServicePackMinorVersion &_
			vbCRLF
		getOSVersion = getOSVersion &  "Version: " & objOS.Version & vbCRLF &_
			vbCRLF
		strOsVersion = objOS.Version
	Next
	Set objWMIService = Nothing
	Set colOSes = Nothing
End Function

'######################################################
'# Function Name: getIPAddress
'# Parameters: None
'# Return Value: 
'#	getIPAddress - List of IP addresses bound to each adapter
'#
'# Description:
'#	Queries the WMI service for a list of network adapters with IP enabled.  For each adapter
'#	Returns the of IP addresses bound to each adapter
'######################################################
Function getIPAddress()
	Dim strComputer, objWMIService, objIPConfigSet, objIPConfig, intCounter
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set objIPConfigSet = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
	getIPAddress = "IP Address(es):" & vbCRLF
	For Each objIPConfig in objIPConfigSet
    If Not IsNull(objIPConfig.IPAddress) Then
        For intCounter=LBound(objIPConfig.IPAddress) to _
					UBound(objIPConfig.IPAddress)
					If(objIPConfig.IPAddress(intCounter) <> "0.0.0.0") Then
						getIPAddress = getIPAddress & objIPConfig.Description(intCounter)
						getIPAddress = getIPAddress & ": " &_
							objIPConfig.IPAddress(intCounter) & vbCRLF
					End If
        Next
    End If
	Next
	getIPAddress = getIPAddress & vbCRLF

	Set objIPConfig = Nothing
	Set objWMIService = Nothing
	Set objIPConfigSet = Nothing
End Function

Function getCurrentUser()
	Dim strComputer, objWSHNetwork, strUser
	Set objWSHNetwork = CreateObject("WScript.Network")
	strUser = objWSHNetwork.Username
	getCurrentUser = "Current User: " & strUser & vbCRLF
End Function

'######################################################
'# Function Name: getDomainInfo
'# Parameters:
'#	strDomainRole - Output parameter which stores the role of the system in the domain
'#	strDomainName - Output parameter which stores the domain which the server is a member of
'# Return Value: None
'#
'# Description:
'#	Queries the WMI service for the domain of which the server is a member of and the
'#	role of that server in the domain.
'######################################################
Function getDomainInfo(ByRef strDomainRole, ByRef strDomainName)
	Dim strComputer, objWMIService, objDomRole, objDomRoleSet, _
		strPartOfDomain, arrRoles, strElement, strRoles

	strComputer = "."

	strDomainRole = ""

	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set objDomRoleSet = objWMIService.ExecQuery _
		("Select * from Win32_ComputerSystem")

	For Each objDomRole in objDomRoleSet

		Select Case objDomRole.DomainRole
			Case 0
				strDomainRole = "Standalone Workstation"
			Case 1
				strDomainRole = "Member Workstation"
			Case 2
				strDomainRole = "Standalone Server"
			Case 3
				strDomainRole = "Member Server"
			Case 4
				strDomainRole = "Backup Domain Controller"
			Case 5
				strDomainRole = "Primary Domain Controller"
		End Select

		strDomainName = objDomRole.Domain
		arrRoles = objDomRole.Roles

		If NOT(objDomRole.DomainRole = 0 Or objDomRole.DomainRole = 2) Then
			getDomainInfo = "Domain: " & strDomainName & vbCRLF &_
				"Domain Role: " & strDomainRole & vbCRLF
		Else
			getDomainInfo = "Server not member of domain." & vbCRLF
		End If
	Next

	Set objWMIService = Nothing
	Set objDomRole = Nothing
End Function

'######################################################
'# Function Name: getUserAccounts
'# Parameters: None
'# Return Value: 
'#	getUserAccounts - list of user accounts and their attributes, tab-separated
'#
'# Description:
'#	Queries the WMI service for a list of local user accounts on the system
'######################################################
Function getUserAccounts()
	Dim strComputer, objWMIService, objAccount, strUserAccounts, _
		objAccountProp, accountProp, userAccount, objUser, _
		strLastLogin, strPwdLastChanged
		
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	' gather all user accounts (SIDType is 1) from the computer
	Set objAccount = objWMIService.ExecQuery ("Select * from Win32_Account Where SIDType=1 AND Domain='" & strServerName & "'")
	
	' set up the header for the output file	
	strUserAccounts = Chr(34) & "User Name" & Chr(34) & vbTab &_
		Chr(34) & "Full Name" & Chr(34) & vbTab &_
		Chr(34) & "Description" & Chr(34) & vbTab &_
		Chr(34) & "Account Type" & Chr(34) & vbTab &_
		Chr(34) & "SID" & Chr(34) & vbTab &_
		Chr(34) & "Password Last Changed" & Chr(34) & vbTab &_
		Chr(34) & "Domain" & Chr(34) & vbTab &_
		Chr(34) & "Password Is Changeable" & Chr(34) & vbTab &_
		Chr(34) & "Password Expires" & Chr(34) & vbTab &_
		Chr(34) & "Password Required" & Chr(34) & vbTab &_
		Chr(34) & "Account Disabled" & Chr(34) & vbTab &_
		Chr(34) & "Account Locked" & Chr(34) & vbTab &_
		Chr(34) & "Last Login" & Chr(34) & vbCrLf
	
	' Output each user account with their attributes, tab separated, one user per line	
	For Each userAccount in objAccount
		Set objUser = GetObject("WinNT://" & strComputer &_
			"/" & userAccount.Name)
		
		On Error Resume Next
		strLastLogin = objUser.LastLogin
		If (Err.Number <> 0) Then
			On Error GoTo 0
			strLastLogin = "<Never>"
		End If 
		
		strPwdLastChanged = objUser.PasswordAge
		strPwdLastChanged = strPwdLastChanged * -1
		strPwdLastChanged = DateAdd("s", strPwdLastChanged, Now)
		
		strUserAccounts = strUserAccounts & Chr(34) & userAccount.Name & Chr(34) & vbTab &_
			Chr(34) & userAccount.FullName & Chr(34) & vbTab &_
			Chr(34) & userAccount.Description & Chr(34) & vbTab &_
			Chr(34) & userAccount.AccountType & Chr(34) & vbTab &_
			Chr(34) & userAccount.SID & Chr(34) & vbTab &_
			Chr(34) & strPwdLastCHanged & Chr(34) & vbTab &_
			Chr(34) & userAccount.Domain & Chr(34) & vbTab &_
			Chr(34) & userAccount.PasswordChangeable & Chr(34) & vbTab &_
			Chr(34) & userAccount.PasswordExpires & Chr(34) & vbTab &_
			Chr(34) & userAccount.PasswordRequired & Chr(34) & vbTab &_
			Chr(34) & userAccount.Disabled & Chr(34) & vbTab &_
			Chr(34) & userAccount.Lockout & Chr(34) & vbTab &_
			Chr(34) & strLastLogin & Chr(34) & vbCRLF
	Next
	
	getUserAccounts = strUserAccounts
	Set objAccount = Nothing
	Set objWMIService = Nothing
	Set objAccountProp = Nothing
	Set objUser = Nothing
End Function

'######################################################
'# Function Name: getGroups
'# Parameters: None
'# Return Value: 
'#	getGroups - Tab-separated
'#
'# Description:
'#	Queries the WMI service for a list of local user accounts on the system
'######################################################
Function getGroups()
	On Error Resume Next
	Dim strComputer, objWMIService, colItems, objItem, strEnumedGroups
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colItems = objWMIService.ExecQuery("Select * from Win32_Group Where Domain='BUILTIN' or Domain='" & strServerName & "'")

	For Each objItem in colItems
		getGroups = getGroups &  Chr(34) & "Name: " & objItem.Name & Chr(34) & vbTab
		getGroups = getGroups &  Chr(34) & "SID: " & objItem.SID & Chr(34) & vbTab
		getGroups = getGroups &  Chr(34) & "Caption: " & objItem.Caption & Chr(34) & vbTab
		getGroups = getGroups &  Chr(34) & "Description: " & objItem.Description & Chr(34) & vbTab
		getGroups = getGroups &  Chr(34) & "Domain: " & objItem.Domain & Chr(34) & vbTab
		getGroups = getGroups & vbCRLF
		strEnumedGroups = ""
		getGroups = getGroups & getGroupMembers(objItem.Name, strEnumedGroups) &_
			vbCRLF
	Next

	Set objWMIService = Nothing
	Set colItems = Nothing
End Function

'######################################################
'# Function Name: getGroupMembers
'# Parameters:
'#	strGroupName - Input parameter containing the name of the group to be enumerated
'#	strEnumedGroups - Output parameter, not used in the function
'# Return Value: 
'#	getGroupMembers - List of the members of strGroupName, one per line
'#
'# Description:
'#	Queries the specified computer and enumerates the members of the group specified
'#	in strGroupName
'######################################################
Function getGroupMembers(ByVal strGroupName, ByRef strEnumedGroups)
	Dim strComputer, objWMIService, objAccount, strGroupMembers, objAdminGroup
	Dim strAdmin, userAccount, objAccountProp
	Dim intCheck

	intCheck = 0
	strComputer = "."

	Set objWMIService = GetObject("winmgmts:" &_
		"{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set objAdminGroup = GetObject("WinNT://" & strComputer & "/" & strGroupName)

	' Get the list of users who are a member of strGroupName, one per line
	For Each strAdmin in objAdminGroup.Members
		strGroupMembers = strGroupMembers & Chr(34) &_
			Right(strAdmin.ADsPath,Len(strAdmin.ADsPath) - _
				InStrRev(strAdmin.ADsPath, "//") - 1) & Chr(34) & vbCRLF
	Next
	getGroupMembers = strGroupMembers
	Set objWMIService = Nothing
	Set objAccount = Nothing
	Set objAdminGroup = Nothing
End Function

'######################################################
'# Function Name: getADUserAccounts
'# Parameters: None
'# Return Value: None
'#
'# Description:
'#	Gathers a list of AD user accounts and outputs this list to the users output file
'######################################################
Function getADUserAccounts()
	On Error Resume Next
	Const ADS_SCOPE_SUBTREE = 2
	
	Dim strLastLogin, objUser, objLastLogin, strComputer, objDomain, _
		intUserFlags, strPwdNeverExpires, strAcctDisabled, strPwdCannotChange, _
		strAcctLocked, strPasswordRequired, strPwdExpired, objWMIService, strSID, _
		strAccountExpirationDate, strAccountExpired, strLastLoginEnv, _
		objRootDSE, objLDAPUser, objConnection, objCommand, objRecordSet, _
		strFullName, objUser2, strPwdLastChanged

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	
	' Get Domain name from RootDSE object.
	Set objRootDSE = GetObject("LDAP://rootDSE")
	Set objDomain = GetObject("LDAP://" &_
		objRootDSE.Get("DefaultNamingContext"))
	If objDomain.[msDS-Behavior-Version] <> 2 Then
		strLastLoginEnv = "This DC Only"
	Else
		strLastLoginEnv = "Domain Within 14 Days"
	End If
	
	' Set up the headers for listing the AD user accounts	
	writeToFile strUsers, Chr(34) & "NT Name" & Chr(34) & vbTab &_
		Chr(34) & "Display Name" & Chr(34) & vbTab &_
		Chr(34) & "Description" & Chr(34) & vbTab &_
		Chr(34) & "SID" & Chr(34) & vbTab &_
		Chr(34) & "Password Last Changed" & Chr(34) & vbTab &_
		Chr(34) & "Password Expired" & Chr(34) & vbTab &_
		Chr(34) & "Password Cannot Change" & Chr(34) & vbTab &_
		Chr(34) & "Password Never Expires" & Chr(34) & vbTab &_
		Chr(34) & "Password Required" & Chr(34) & vbTab &_
		Chr(34) & "Account Disabled" & Chr(34) & vbTab &_
		Chr(34) & "Account Locked" & Chr(34) & vbTab &_
		Chr(34) & "Last Login (" & strLastLoginEnv & ")" & Chr(34) & vbTab &_
		Chr(34) & "Account Expiration Date" & Chr(34) & vbCRLF

	Set objConnection = CreateObject("ADODB.Connection")'
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection

	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

	Set objDomain = GetObject("WinNT://" & strComputer)
	objDomain.Filter = Array("User")
	
	For Each objUser In objDomain
		objCommand.CommandText = _
			"SELECT distinguishedName FROM 'LDAP://" &_
			objRootDSE.Get("DefaultNamingContext") &_
			"' WHERE sAMAccountName='" & objUser.name &_
			"' AND objectCategory='person' AND objectClass='user'"
			
		Set objRecordSet = objCommand.Execute
		objRecordSet.MoveFirst
		
		Do Until objRecordSet.EOF
			Set objUser2 = GetObject("LDAP://" & objRecordSet.Fields("distinguishedName").Value)	

			strPwdLastChanged = objUser2.PasswordLastChanged
			
			If(StrComp(strLastLoginEnv, "Domain Within 14 Days") = 0) Then
				strLastLogin = "Never"
				
				Set objLastLogin = objUser2.Get("lastLogonTimestamp")

				strLastLogin = objLastLogin.HighPart * (2^32) + objLastLogin.LowPart 
				strLastLogin = strLastLogin / (60 * 10000000)
				strLastLogin = strLastLogin / 1440
				strLastLogin = strLastLogin + #1/1/1601#
				
				Set objLastLogin = Nothing
			Else
				strLastLogin = "1/1/1970"
				strLastLogin = objUser.LastLogin
			End If

			If strLastLogin = "" Then
				strLastLogin = "Never"
			End If

			If(objUser2.PasswordRequired = True) Then
				strPasswordRequired = "Yes"
			Else
				strPasswordRequired = "No"
			End If
				
			strAccountExpirationDate = objUser2.AccountExpirationDate
			
			If(Err <> 0) Then
				strAccountExpirationDate = "Never"
			End If
			
			strAccountExpirationDate = Replace(straccountExpirationDate,"1/1/1970",_
				"Never")
			strLastLogin = Replace(FormatDateTime(strLastLogin),"1/1/1970",_
				"Never")
			strLastLogin = Replace(FormatDateTime(strLastLogin),"1/1/1601",_
				"Never")

			objRecordSet.MoveNext
		Loop
		
		strSID = convObjectSID(objUser.objectSID)		

		intUserFlags = objUser.UserFlags
		If((intUserFlags And &h00040) = 0) Then
			strPwdCannotChange = "No"
		Else
			strPwdCannotChange = "Yes"
		End If
		If((intUserFlags And &h10000) = 0) Then
			strPwdNeverExpires = "No"
		Else
			strPwdNeverExpires = "Yes"
		End If
		If(objUser.PasswordExpired = True) Then
			strPwdExpired = "No"
		Else
			strPwdExpired = "Yes"
		End If
		If(objUser.AccountDisabled = True) Then
			strAcctDisabled = "Yes"
		Else
			strAcctDisabled = "No"
		End If
		If(objUser.IsAccountLocked = True) Then
			strAcctLocked = "Yes"
		Else
			strAcctLocked = "No"
		End If

		strFullName = objUser.FullName
 
		writeToFile strUsers, Chr(34) & objUser.name & Chr(34) & vbTab &_
			Chr(34) & strFullName & Chr(34) & vbTab &_
			Chr(34) & objUser.Description & Chr(34) & vbTab &_
			Chr(34) & strSID & Chr(34) & vbTab &_
			Chr(34) & strPwdLastChanged & Chr(34) & vbTab &_
			Chr(34) & strPwdExpired & Chr(34) & vbTab &_
			Chr(34) & strPwdCannotChange & Chr(34) & vbTab &_
			Chr(34) & strPwdNeverExpires & Chr(34) & vbTab &_
			Chr(34) & strPasswordRequired & Chr(34) & vbTab &_
			Chr(34) & strAcctDisabled & Chr(34) & vbTab &_
			Chr(34) & strAcctLocked & Chr(34) & vbTab &_
			Chr(34) & strLastLogin & Chr(34) & vbTab &_
			Chr(34) & strAccountExpirationDate & Chr(34) & vbCRLF
	Next
	
	Set objDomain = Nothing
	Set objUser = Nothing
	Set objLDAPUser = Nothing
End Function

'######################################################
'# Function Name: getADGroups
'# Parameters: None
'# Return Value: None
'#
'# Description:
'#	Gathers a list of AD user accounts and outputs this list to the users output file
'######################################################
Function getADGroups()
	On Error Resume Next

	Dim strComputer, objDomain, objDomain2, objGroup, objUser, _
		objGroupItemGroups, objGroupItemUsers, objGroupEnum, arrEnumedGroups, _
		strGroup, objWMIService, colItems, strSID, objGroupSID, objItem

	strComputer = "."
	
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set objDomain = GetObject("WinNT://" & strComputer)
	Set objDomain2 = GetObject("WinNT://" & strComputer)
	objDomain.Filter = Array("Group")
	objDomain2.Filter = Array("Group")
	
	For Each objGroup In objDomain
		Set objGroupItemUsers = GetObject("WinNT://" & strComputer & "/" &_
			objGroup.Name & ",group")
		objGroupItemUsers.Members.Filter = Array("User")
		
		strSID = convObjectSID(objGroup.objectSID)

		If(strSID = "") Then
			strSID = "Local Account"
		End If
		
		writeToFile strGroups, Chr(34) & "Name: " & objGroup.Name & Chr(34) & vbTab
		writeToFile strGroups, Chr(34) & "Description: " & objGroup.Description & Chr(34) & vbTab
		writeToFile strGroups, Chr(34) & "SID: " & strSID & Chr(34) & vbCRLF

		For Each objUser In objGroupItemUsers.Members
			writeToFile strGroups, vbTab &_
				Chr(34) & Right(objUser.ADsPath,Len(objUser.ADsPath) - _
				InStrRev(objUser.ADsPath, "//") - 1) & Chr(34) & vbCRLF
		Next
		writeToFile strGroups, vbCRLF
	Next
	Set objDomain = Nothing
	Set objGroup = Nothing
	Set objUser = Nothing
	Set objGrouqpItemGroups = Nothing
	Set objGroupItemUsers = Nothing
	Set objGroupEnum = Nothing
	Set arrEnumedGroups = Nothing
End Function

'######################################################
'# Function Name: convObjectSID
'# Parameters: 
'#	strBinSID - Input parameter containing the SID to be converted
'# Return Value: 
'#	convObjectSID - Converted SID
'#
'# Description:
'#	Takes a passed in SID and converts it to the more common, readable form
'######################################################	
Function convObjectSID(ByVal strBinSID)
	Dim strChar, strSID, intLength, SID1, SID2, SID3, RID
	Dim arrSID, SID
	
	strChar = ""
	strSID = ""
	For intLength = 1 to Lenb(strBinSID)
		strSID = strSID & strChar & Right("0" & Hex(ascb(Midb(strBinSID, _
			intLength ,1))),2)
		strChar = ","
	Next
	
	arrSID = strSID
	
	SID = Split(arrSID,",")
	
	SID1 = (convHexToDec(Mid(SID(15), 1, 1))*268435456) + _
		(convHexToDec(Mid(SID(15), 2, 2))*16777216) + _
		(convHexToDec(Mid(SID(14), 1, 1))*1048576) + _
		(convHexToDec(Mid(SID(14), 2, 2))*65536) + _
		(convHexToDec(Mid(SID(13), 1, 1))*4096) + _
		(convHexToDec(Mid(SID(13), 2, 2))*256) + _
		(convHexToDec(Mid(SID(12), 1, 1))*16) + _
		convHexToDec(Mid(SID(12), 2, 2))
	SID2 = (convHexToDec(Mid(SID(19), 1, 1))*268435456) + _
		(convHexToDec(Mid(SID(19), 2, 2))*16777216) + _
		(convHexToDec(Mid(SID(18), 1, 1))*1048576) + _
		(convHexToDec(Mid(SID(18), 2, 2))*65536) + _
		(convHexToDec(Mid(SID(17), 1, 1))*4096) + _
		(convHexToDec(Mid(SID(17), 2, 2))*256) + _
		(convHexToDec(Mid(SID(16), 1, 1))*16) + _
		convHexToDec(Mid(SID(16), 2, 2))
	SID3 = (convHexToDec(Mid(SID(23), 1, 1))*268435456) + _
		(convHexToDec(Mid(SID(23), 2, 2))*16777216) + _
		(convHexToDec(Mid(SID(22), 1, 1))*1048576) + _
		(convHexToDec(Mid(SID(22), 2, 2))*65536) + _
		(convHexToDec(Mid(SID(21), 1, 1))*4096) + _
		(convHexToDec(Mid(SID(21), 2, 2))*256) + _
		(convHexToDec(Mid(SID(20), 1, 1))*16) + _
		convHexToDec(Mid(SID(20), 2, 2))
	RID = (convHexToDec(Mid(SID(27), 1, 1))*268435456) + _
		(convHexToDec(Mid(SID(27), 2, 2))*16777216) + _
		(convHexToDec(Mid(SID(26), 1, 1))*1048576) + _
		(convHexToDec(Mid(SID(26), 2, 2))*65536) + _
		(convHexToDec(Mid(SID(25), 1, 1))*4096) + _
		(convHexToDec(Mid(SID(25), 2, 2))*256) + _
		(convHexToDec(Mid(SID(24), 1, 1))*16) + _
		convHexToDec(Mid(SID(24), 2, 2))
	
	convObjectSID = "S-1-5-21-" & SID1 & "-" & SID2 & "-" & SID3 & "-" & RID
End Function 

'######################################################
'# Function Name: convHexToDec
'# Parameters: 
'#	sHex - Input parameter which contains an hex value
'# Return Value: 
'#	convHexToDec - Decimal value of passed in hex value
'#
'# Description:
'#	Takes a hex value and converts it to a decimal value
'######################################################	
Function convHexToDec(ByVal sHex)
	convHexToDec = "" & CLng("&H" & sHex)
End Function 

'######################################################
'# Function Name: getADGroupMembers
'# Parameters: 
'#	strGroupEnum - Input parameter containing the group to be enumerated
'#	arrEnumedGroups - Output parameter containing a list of previously enumerated groups
'# Return Value: 
'#	getADGroupMembers - A list of users in the specified group
'#
'# Description:
'#	Takes an AD group and enumerates all of the users in that group
'######################################################	
Function getADGroupMembers(ByVal strGroupEnum, ByRef arrEnumedGroups)
	Dim objDomain, objGroupItemUsers, strUser, intFound, strGroup,_
		strMembers, strComputer

	strComputer = "."

	Set objDomain = GetObject("WinNT://" & strComputer)
	objDomain.Filter = Array("Group")
	Set objGroupItemUsers = GetObject("WinNT://" & strComputer &_
		"/" & strGroupEnum & ",group")
	objGroupItemUsers.Members.Filter = Array("User")
	For Each strUser In objGroupItemUsers.Members
		intFound = 0
		For Each strGroup In objDomain
			If((StrComp(strUser.Name, strGroup.Name) = 0) And _
				InStr(arrEnumedGroups, strGroup.Name & ",") = 0) Then
				arrEnumedGroups = arrEnumedGroups & strGroup.Name & ","
				intFound = 1
				strMembers = strMembers & vbTab & getADGroupMembers(strGroup.Name, _
					arrEnumedGroups) & vbCRLF
			End If
		Next
		If(intFound = 0) Then
			strMembers = strMembers & vbTab & Chr(34) & strUser.Name & Chr(34) & vbCRLF
		End If
	Next
	getADGroupMembers = getADGroupMembers & strMembers
	Set objDomain = Nothing
	Set objGroupItemUsers = Nothing
	Set strGroup = Nothing
End Function

'######################################################
'# Function Name: getDirPermissions
'# Parameters: 
'#	strFolderName - Input parameter containing the group to be enumerated
'# Return Value: 
'#	getDirPermissions - The Permissions on the specified directory
'#
'# Description:
'#	Takes a specified directory and returns the permissions
'######################################################
Function getDirPermissions(ByVal strFolderName)
	Const SE_DACL_PRESENT = &h4
	Const SE_SACL_PRESENT = &h10
	Const ACCESS_ALLOWED_ACE_TYPE = &h0
	Const ACCESS_DENIED_ACE_TYPE  = &h1
	Const SUCCESS_AUDIT_TYPE = &h40
	Const FAILURE_AUDIT_TYPE = &h80

	Const FILE_ALL_ACCESS         = &h1f01ff
	Const FOLDER_ADD_SUBDIRECTORY = &h000004
	Const FILE_DELETE             = &h010000
	Const FILE_DELETE_CHILD       = &h000040
	Const FOLDER_TRAVERSE         = &h000020
	Const FILE_READ_ATTRIBUTES    = &h000080
	Const FILE_READ_CONTROL       = &h020000
	Const FOLDER_LIST_DIRECTORY   = &h000001
	Const FILE_READ_EA            = &h000008
	Const FILE_WRITE_ATTRIBUTES   = &h000100
	Const FILE_WRITE_DAC          = &h040000
	Const FOLDER_ADD_FILE         = &h000002
	Const FILE_WRITE_EA           = &h000010
	Const FILE_WRITE_OWNER        = &h080000
	
	Dim vbTab3 
	vbTab3 = vbTab & vbTab & vbTab

	Dim strDACL, objWMIService, objFolderSecuritySettings, intControlFlags
	Dim intRetVal, arrACEs, objSD, strSpecialPerms
	Dim objACE, strTemp, objFolder, strSACL, arrAEs, objAE, intAccessMask
	strDACL = strFolderName & ":" & vbCRLF & vbTab
	Set objWMIService = GetObject("winmgmts:")
	Set objFolderSecuritySettings = _
		objWMIService.ExecQuery("Select * From Win32_LogicalFileSecuritySetting" &_
			" Where Path='" &	Replace(strFolderName,"\","\\") & "'")

	For Each objFolder in objFolderSecuritySettings
		intRetVal = objFolder.GetSecurityDescriptor(objSD)

		intControlFlags = objSD.ControlFlags

		If (intControlFlags And SE_DACL_PRESENT) And (IsArray(objSD.DACL)) Then
			arrACEs = objSD.DACL

			For Each objACE in arrACEs
				strTemp = ""
				strSpecialPerms = ""
				If (objACE.AccessMask = 2032127) Then
					strTemp = strTemp & vbTab & "Full Control" & vbCRLF
				ElseIf (objACE.AccessMask = 1245631) Then
					strTemp = strTemp & vbTab & "Modify" & vbCRLF
				ElseIf (objACE.AccessMask = 1179817) Then
					strTemp = strTemp & vbTab & "Read & Execute" & vbCRLF
				ElseIf (objACE.AccessMask = 1179785) Then
					strTemp = strTemp & vbTab & "Read" & vbCRLF
				ElseIf (objACE.AccessMask = 1048854) Then
					strTemp = strTemp & vbTab & "Write" & vbCRLF
				ElseIf (objACE.AccessMask = 1180063) Then
					strTemp = strTemp & vbTab & "Read & Write" & vbCRLF
				ElseIf (objACE.AccessMask = 1180095) Then
					strTemp = strTemp & vbTab & "Read, Write & Execute" & vbCRLF
				Else
					strSpecialPerms = "Special Permissions: " & vbCRLF
					If objACE.AccessMask And FOLDER_ADD_SUBDIRECTORY Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Create Folders" &_
							vbCRLF
					End If
					If objACE.AccessMask And FILE_DELETE Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Delete" & vbCRLF
					End If
					If objACE.AccessMask And FILE_DELETE_CHILD Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Delete Subfolders " &_
							"And Files" & vbCRLF
					End If
					If objACE.AccessMask And FOLDER_TRAVERSE Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Traverse Folder" &_
							vbCRLF
					End If
					If objACE.AccessMask And FILE_READ_ATTRIBUTES Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Read Attributes" &_
							vbCRLF
					End If
					If objACE.AccessMask And FILE_READ_CONTROL Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Read Permissions" &_
							vbCRLF
					End If
					If objACE.AccessMask And FOLDER_LIST_DIRECTORY Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "List Folder" & vbCRLF 
					End If
					If objACE.AccessMask And FILE_READ_EA Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Read Extended " &_
							"Attributes" & vbCRLF
					End If
					If objACE.AccessMask And FILE_WRITE_ATTRIBUTES Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Write Attributes" &_
							vbCRLF
					End If
					If objACE.AccessMask And FILE_WRITE_DAC Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Change Permissions" &_
							vbCRLF
					End If
					If objACE.AccessMask And FOLDER_ADD_FILE Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Create Files" & vbCRLF
					End If
					If objACE.AccessMask And FILE_WRITE_EA Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Write Extended " &_
							"Attributes" & vbCRLF
					End If
					If objACE.AccessMask And FILE_WRITE_OWNER Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Take Ownership" &_
							vbCRLF
					End If
					strSpecialPerms = strSpecialPerms & vbTab
				End If
				
				If(StrComp(strSpecialPerms, "Special Permissions: " &_
					vbCRLF & vbTab) = 0) Then
					strSpecialPerms = ""
				End If
				
				If (strTemp <> "" Or strSpecialPerms <> "") Then
					strDACL = strDACL & objACE.Trustee.Domain & "\" &_
						objACE.Trustee.Name
					If objACE.AceType = ACCESS_ALLOWED_ACE_TYPE Then
						strDACL = strDACL & vbTab & "Allowed: "
					ElseIf objACE.AceType = ACCESS_DENIED_ACE_TYPE Then
						strDACL = strDACL & vbTab & "Denied: "
					End If
					strDACL = strDACL & strTemp & vbTab & strSpecialPerms
				End If
			Next
		Else
			strDACL = strDACL & "No DACL present in security descriptor" & vbCRLF
		End If

		If (intControlFlags And SE_SACL_PRESENT) And (IsArray(objSD.SACL)) Then
			arrAEs = objSD.SACL
			strSACL = "Auditing: " & vbCRLF

			For Each objAE in arrAEs
				strTemp = ""
				
				If(objAE.AceFlags And SUCCESS_AUDIT_TYPE) Then
					strTemp = vbTab & vbTab & "Success: " & vbCRLF
				End If
				If(objAE.AceFlags And FAILURE_AUDIT_TYPE) Then
					strTemp = vbTab & vbTab & "Failure: " & vbCRLF
				End If
			
				If objAE.AccessMask And FOLDER_ADD_SUBDIRECTORY Then
					strTemp = strTemp & vbTab3 & "Create Folder" & vbCRLF
				End If
				If objAE.AccessMask And FILE_DELETE Then
					strTemp = strTemp & vbTab3 & "Delete" & vbCRLF
				End If
				If objAE.AccessMask And FILE_DELETE_CHILD Then
					strTemp = strTemp & vbTab3 & "Delete Subfolders and " &_
						"Files" & vbCRLF
				End If
				If objAE.AccessMask And FOLDER_TRAVERSE Then
					strTemp = strTemp & vbTab3 & "Traverse Folder" & vbCRLF
				End If
				If objAE.AccessMask And FILE_READ_ATTRIBUTES Then
					strTemp = strTemp & vbTab3 & "Read Attributes" & vbCRLF
				End If
				If objAE.AccessMask And FILE_READ_CONTROL Then
					strTemp = strTemp & vbTab3 & "Read Permissions" & vbCRLF
				End If
				If objAE.AccessMask And FOLDER_LIST_DIRECTORY Then
					strTemp = strTemp & vbTab3 & "List Folder" & vbCRLF
				End If
				If objAE.AccessMask And FILE_READ_EA Then
					strTemp = strTemp & vbTab3 & "Read Extended Attributes" & vbCRLF
				End If
				If objAE.AccessMask And FILE_WRITE_ATTRIBUTES Then
					strTemp = strTemp & vbTab3 & "Write Attributes" & vbCRLF
				End If
				If objAE.AccessMask And FILE_WRITE_DAC Then
					strTemp = strTemp & vbTab3 & "Change Permissions" & vbCRLF
				End If
				If objAE.AccessMask And FOLDER_ADD_FILE Then
					strTemp = strTemp & vbTab3 & "Create Files" & vbCRLF
				End If
				If objAE.AccessMask And FILE_WRITE_EA Then
					strTemp = strTemp & vbTab3 & "Write Extended Attributes" & vbCRLF
				End If
				If objAE.AccessMask And FILE_WRITE_OWNER Then
					strTemp = strTemp & vbTab3 & "Take Ownership" & vbCRLF
				End If

				If strTemp <> "" Then
					strSACL = strSACL & vbTab & objAE.Trustee.Domain & "\" &_
						objAE.Trustee.Name & vbCRLF
					strSACL = strSACL & strTemp
				End If
			Next
		Else
			strSACL = strSACL & "(No auditing)" & vbCRLF
		End If
	Next

	getDirPermissions = strDACL & strSACL & vbCRLF
	Set objWMIService = Nothing
	Set arrACEs = Nothing
	Set objSD = Nothing
	Set objACE = Nothing
	Set objFolderSecuritySettings = Nothing
	Set objFolder = Nothing
End Function

'######################################################
'# Sub-Function Name: getSubDirPerms
'# Parameters: 
'#	strPath - Input parameter of the path to enumerate from
'#	strDirPerms - Input parameter of the file to write output to
'# Return Value: 
'#	None
'#
'# Description:
'#	Takes a root directory, and enumerates the permissions on subdirectories
'######################################################
' Subroutine for Custom Directory Permissions
Sub getSubDirPerms(ByVal strPath, ByVal strDirPerms)
	'On Error Resume Next
	Dim strFolderPath, objFSO, objFolder, SubFolder, strErrorMsg
	strFolderPath = strPath

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.getFolder(strFolderPath)

	If ( InStr(strFolderPath, "'") ) Then 
		'record error, move on
		strErrorMsg = "*** WARNING: Directory Path Contains A Single Quote - Skipping! ***"
		writeToFile strDirPerms, strFolderPath & vbCRLF & vbTab & strErrorMsg & vbCRLF & vbCRLF
	Else
		'Get the root dir's permissions	
		writeToFile strDirPerms, getDirPermissions(strPath)
		'Get the subdirectory's permissions
		For Each SubFolder In objFolder.subFolders
			getSubDirPerms SubFolder, strDirPerms
		Next
	End If
	
	
	Set objFSO = Nothing
	Set objFolder = Nothing
	Set SubFolder = Nothing
	
End Sub

'######################################################
'# Function Name: getFilePermissions
'# Parameters: 
'#	strFileName - Input parameter containing the filename to evaluate permissions on
'# Return Value: 
'#	getFilePermissions - Permissions on the specified file
'#
'# Description:
'#	Takes the filename and returns the permissions on the file
'######################################################
Function getFilePermissions(strFileName)
	Const SE_DACL_PRESENT = &h4
	Const ACCESS_ALLOWED_ACE_TYPE = &h0
	Const ACCESS_DENIED_ACE_TYPE  = &h1

	Const FILE_ALL_ACCESS       = &h1f01ff
	Const FILE_APPEND_DATA      = &h000004	
	Const FILE_DELETE           = &h010000	
	Const FILE_DELETE_CHILD     = &h000040	
	Const FILE_EXECUTE          = &h000020	
	Const FILE_READ_ATTRIBUTES  = &h000080	
	Const FILE_READ_CONTROL     = &h020000	
	Const FILE_READ_DATA        = &h000001	
	Const FILE_READ_EA          = &h000008	
	Const FILE_WRITE_ATTRIBUTES = &h000100	
	Const FILE_WRITE_DAC        = &h040000	
	Const FILE_WRITE_DATA       = &h000002	
	Const FILE_WRITE_EA         = &h000010	
	Const FILE_WRITE_OWNER      = &h080000	

	Dim vbTab3 
	vbTab3 = vbTab & vbTab & vbTab
	
	Dim objWMIService, objAE,  intControlFlags, intRetVal, _
		objFileSecuritySettings, strTemp, objSD, arrACEs
	Dim objFile, strSpecialPerms
	
	getFilePermissions = strFileName & ":" & vbCRLF & vbTab

	Set objWMIService = GetObject("winmgmts:")

	Set objFileSecuritySettings = _
		objWMIService.ExecQuery("Select * From Win32_LogicalFileSecuritySetting" &_
			" Where Path='" & Replace(strFileName,"\","\\") & "'")
	
	For Each objFile in objFileSecuritySettings
		intRetVal = objFile.GetSecurityDescriptor(objSD)

		intControlFlags = objSD.ControlFlags

		If intControlFlags And SE_DACL_PRESENT Then
			arrACEs = objSD.DACL

			For Each objAE in arrACEs
				strTemp = ""
				strSpecialPerms = ""
				If (objAE.AccessMask = 2032127) Then
					strTemp = strTemp & vbTab & "Full Control" & vbCRLF
				ElseIf (objAE.AccessMask = 1245631) Then
					strTemp = strTemp & vbTab & "Modify" & vbCRLF
				ElseIf (objAE.AccessMask = 1179817) Then
					strTemp = strTemp & vbTab & "Read & Execute" & vbCRLF
				ElseIf (objAE.AccessMask = 1179785) Then
					strTemp = strTemp & vbTab & "Read" & vbCRLF
				ElseIf (objAE.AccessMask = 1048854) Then
					strTemp = strTemp & vbTab & "Write" & vbCRLF
				ElseIf (objAE.AccessMask = 1180063) Then
					strTemp = strTemp & vbTab & "Read & Write" & vbCRLF
				ElseIf (objAE.AccessMask = 1180095) Then
					strTemp = strTemp & vbTab & "Read, Write & Execute" & vbCRLF
				Else
					strSpecialPerms = strSpecialPerms & "Special Permissions: " & vbCRLF
					If objAE.AccessMask And FILE_APPEND_DATA Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Append Data" & vbCRLF
					End If
					If objAE.AccessMask And FILE_DELETE Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Delete" & vbCRLF
					End If
					If objAE.AccessMask And FILE_EXECUTE Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Execute File" &_
							vbCRLF
					End If
					If objAE.AccessMask And FILE_READ_ATTRIBUTES Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Read Attributes" &_
							vbCRLF
					End If
					If objAE.AccessMask And FILE_READ_CONTROL Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Read Permissions" &_
							vbCRLF
					End If
					If objAE.AccessMask And FILE_READ_DATA Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Read Data" &_
							vbCRLF
					End If
					If objAE.AccessMask And FILE_READ_EA Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Read Attributes" &_
							vbCRLF
					End If
					If objAE.AccessMask And FILE_WRITE_ATTRIBUTES Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Write Attributes" &_
							vbCRLF
					End If
					If objAE.AccessMask And FILE_WRITE_DAC Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Change " &_
							"Permissions" & vbCRLF
					End If
					If objAE.AccessMask And FILE_WRITE_DATA Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Write Data" & vbCRLF
					End If
					If objAE.AccessMask And FILE_WRITE_EA Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Write Extended " &_
							"Attributes" & vbCRLF
					End If
					If objAE.AccessMask And FILE_WRITE_OWNER Then
						strSpecialPerms = strSpecialPerms & vbTab3 & "Take Ownership" &_
							vbCRLF
					End If
					strSpecialPerms = strSpecialPerms & vbTab
				End If
				
				If(StrComp(strSpecialPerms, "Special Permissions: " &_
					vbCRLF & vbTab) = 0) Then
					strSpecialPerms = ""
				End If
				
				If(strTemp <> "" Or strSpecialPerms <> "") Then
					getFilePermissions = getFilePermissions & objAE.Trustee.Domain &_
						"\" & objAE.Trustee.Name
					If objAE.AceType = ACCESS_ALLOWED_ACE_TYPE Then
						getFilePermissions = getFilePermissions & vbTab & "Allowed: "
					ElseIf objAE.AceType = ACCESS_DENIED_ACE_TYPE Then
						getFilePermissions = getFilePermissions & vbTab & "Denied: "
					End If
					getFilePermissions = getFilePermissions & strTemp & vbTab &_
						strSpecialPerms
				End If
			Next
			getFilePermissions = getFilePermissions & vbCRLF
		Else
			getFilePermissions = getFilePermissions & "No DACL present in " &_
				"security descriptor" & vbCRLF
		End If
	Next
	
	If(StrComp(getFilePermissions, strFileName & ":" & vbCRLF & vbTab) = 0) Then
		getFilePermissions = getFilePermissions & "File does not exist!" &_
			vbCRLF & vbCRLF
	End If
	
	Set objAE = Nothing
	set arrACEs = Nothing
	Set objWMIService = Nothing
	Set objFileSecuritySettings = Nothing
	Set objFile = Nothing
End Function

'######################################################
'# Function Name: listHotFixes
'# Parameters: 
'#	None
'# Return Value: 
'#	listHotFixes - A list of applied hotfixes
'#
'# Description:
'#	Returns a list of applied hotfixes to the system
'######################################################
Function listHotFixes()
	Dim strComputer, objWMIService, colOperatingSystems, objOperatingSystem,_
		colQuickFixes, objQuickFix

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")

	For Each objOperatingSystem in colOperatingSystems
    listHotFixes = listHotFixes & "Service Pack: " &_
			objOperatingSystem.ServicePackMajorVersion  &_
			"." & objOperatingSystem.ServicePackMinorVersion & vbCRLF & vbCRLF
  Next

	Set colQuickFixes = objWMIService.ExecQuery _
    ("Select * from Win32_QuickFixEngineering")

	For Each objQuickFix in colQuickFixes
		If(StrComp(objQuickFix.HotFixID, "File 1") <> 0) Then
			listHotFixes = listHotFixes & "Description: " &_
				objQuickFix.Description & vbTab
			listHotFixes = listHotFixes & "Hot Fix ID: " &_
				objQuickFix.HotFixID & vbCRLF
		End If
	Next
	Set objWMIService = Nothing
	Set colOperatingSystems = Nothing
	Set objQuickFix = Nothing
	Set colQuickFixes = Nothing
End Function

'######################################################
'# Function Name: enumerateRegistryValues
'# Parameters: 
'#	strHive - Input parameter containing the registry hive to enumerate
'#	strKeyPath - Input parameter containing the registry key to enumerate
'# Return Value: 
'#	enumerateRegistryValues - The registry values and data in the key specified
'#
'# Description:
'#	Takes an AD group and enumerates all of the users in that group
'######################################################
Function enumerateRegistryValues(ByVal strHive, ByVal strKeyPath)
	Const HKEY_CURRENT_USER = &h80000001
	Const HKEY_LOCAL_MACHINE = &h80000002
	Const HKEY_USERS = &h80000003
	Const REG_SZ = 1
	Const REG_EXPAnd_SZ = 2
	Const REG_BINARY = 3
	Const REG_DWORD = 4
	Const REG_MULTI_SZ = 7

	Dim strHiveTrans, strAllValues, strValue, strComputer, objReg, strTemp, _
		arrValueTypes, arrValueNames, intCounter, intTemp, arrValues, _
		intRetVal, counter

	strComputer = "."

	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
			strComputer & "\root\default:StdRegProv")

	'Convert Text Name to HEX
	Select Case strHive
		Case "HKEY_LOCAL_MACHINE"
			strHiveTrans = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"
			strHiveTrans = HKEY_USERS
		Case "HKEY_CURRENT_USER"
			strHiveTrans = HKEY_CURRENT_USER
	End Select

	'Enumerate the Value names and Types from the Registry Key (strHive)
	intRetVal = objReg.EnumValues(strHiveTrans, strKeyPath, arrValueNames, _
		arrValueTypes)
	strAllValues = "Registry Key: " & strHive & "\" & strKeyPath & ":"

	If(intRetVal = 0) Then
		If(IsArray(arrValueNames)) Then
			For intCounter=0 To UBound(arrValueNames)
				strAllValues = strAllValues & vbCRLF & "Value Name: " &_
					arrValueNames(intCounter)

				Select Case arrValueTypes(intCounter)
					Case REG_SZ
						objReg.GetStringValue strHiveTrans, strKeypath, _
							arrValueNames(intCounter), strValue
						strAllValues = strAllValues & vbTab & "Data: "
						
						For counter = 1 to Len(strValue) 
							strTemp = strTemp & Chr(Asc(Mid(strValue, counter, _
								1)))
						Next
						strValue = strTemp
						strTemp = ""
						
						strAllValues = strAllValues & strValue
					Case REG_EXPAnd_SZ
						objReg.GetExpandedStringValue strHiveTrans, strKeypath, _
							arrValueNames(intCounter), strValue
						strAllValues = strAllValues & vbTab & "Data: "
						strAllValues = strAllValues & strValue
					Case REG_BINARY
						objReg.GetBinaryValue strHiveTrans, strKeypath, _
							arrValueNames(intCounter), strValue
						strAllValues = strAllValues & vbTab & "Data: "
						If(IsArray(strValue)) Then
							For intTemp = lBound(strValue) To uBound(strValue)
								strAllValues = strAllValues &  strValue(intTemp) & " "
							Next
						Else
								strAllValues = strAllValues &  strValue & " "
						End If
					Case REG_DWORD
						objReg.GetDWORDValue strHiveTrans, strKeypath, _
							arrValueNames(intCounter), strValue
						strAllValues = strAllValues & vbTab & "Data: "
						If(StrComp(arrValueNames(intCounter), "MaxSize") = 0) Then
							strValue = strValue / 1024
						ElseIf(StrComp(arrValueNames(intCounter), "Retention") = 0) Then
							strValue = strValue / 86400
						End If
						strAllValues = strAllValues & strValue
					Case	 REG_MULTI_SZ
						objReg.GetMultiStringValue strHiveTrans, strKeypath, _
							arrValueNames(intCounter), arrValues
						strAllValues = strAllValues & vbTab & "Data: "
						For Each strValue In arrValues
							strAllValues = strAllValues & " " & strValue
						Next
				End Select
			Next
			strAllValues = strAllValues & vbCRLF & vbCRLF
			If(Len(strAllValues) = 0) Then
				strAllValues = strAllValues & vbCRLF & "No Registry Values" &_
					vbCRLF & vbCRLF
			End If
		Else
			strAllValues = strAllValues & vbCRLF & "No Registry Values" &_
				vbCRLF & vbCRLF
		End If
	Else
		strAllValues = strAllValues & vbCRLF & "Registry Key Does Not Exist" &_
			vbCRLF & vbCRLF
	End If

	'Return the list of Values and Data
	enumerateRegistryValues = strAllValues
	Set objReg = Nothing
End Function

'######################################################
'# Function Name: enumerateSubKeys
'# Parameters: 
'#	strKeyPath - Input parameter containing the key to enumerate
'#	strHive - Input parameter containing the hive to enumerate
'# Return Value: 
'#	enumerateSubKeys - A list of subkeys of the key specified
'#
'# Description:
'#	Enumerats all of the subkeys under the specified key path
'######################################################
Function enumerateSubKeys(ByVal strHive, ByVal strKeyPath)
	Const HKEY_CURRENT_USER = &h80000001
	Const HKEY_LOCAL_MACHINE = &h80000002
	Const HKEY_USERS = &h80000003
	Const REG_SZ = 1
	Const REG_EXPAnd_SZ = 2
	Const REG_BINARY = 3
	Const REG_DWORD = 4
	Const REG_MULTI_SZ = 7

	Dim arrSubKeys, strHiveTrans

	'Convert Text Name to HEX
	Select Case strHive
		Case "HKEY_LOCAL_MACHINE"
			strHiveTrans = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"
			strHiveTrans = HKEY_USERS
		Case "HKEY_CURRENT_USER"
			strHiveTrans = HKEY_CURRENT_USER
	End Select

	strComputer = "."
	enumerateSubKeys = strHive & "\" & strKeyPath & ":" & vbCRLF

	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
    strComputer & "\root\default:StdRegProv")

	objReg.EnumKey strHiveTrans, strKeyPath, arrSubKeys

	For Each objSubkey In arrSubKeys
		enumerateSubKeys = enumerateSubKeys & objSubkey & vbCRLF
	Next
	Set objReg = Nothing
End Function

'######################################################
'# Function Name: listServices
'# Parameters: 
'#	None
'# Return Value: 
'#	listServices - A list of services and corresponding information
'#
'# Description:
'#	Returns all services and services' information
'######################################################
Function listServices()
	Dim strComputer, objWMIService, colListOfServices, objService, _
		counter, strTemp

	listServices = _
    "Service Name" & vbTab &_
		"Service State" & vbTab &_
		"Caption" & vbTab & "Description" & vbTab & "Can Interact with Desktop" &_
			vbTab &_
		"Display Name" & vbTab & "Error Control" & vbTab &_
		"Executable Path Name" & vbTab &_
		"Service Started" & vbTab &_
    "Start Mode" & vbTab & "Account Name " & vbCRLF

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colListOfServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service")

	For Each objService in colListOfServices
		listServices = listServices & Chr(34) & objService.Name & _
			Chr(34) & vbTab
		listServices = listServices & Chr(34) & objService.State & _
			Chr(34) & vbTab
		listServices = listServices & Chr(34) & objService.Caption & _
			Chr(34) & vbTab
		listServices = listServices & Chr(34) & objService.Description & _
			Chr(34) & vbTab
		listServices = listServices & Chr(34) & objService.DesktopInteract & _
			Chr(34) & vbTab
		listServices = listServices & Chr(34) & objService.DisplayName & _
			Chr(34) & vbTab
		listServices = listServices & Chr(34) & objService.ErrorControl & _
			Chr(34) & vbTab
		listServices = listServices & Chr(34) & objService.PathName & _
			Chr(34) & vbTab
		listServices = listServices & Chr(34) & objService.Started & _
			Chr(34) & vbTab
		listServices = listServices & Chr(34) & objService.StartMode & _
			Chr(34) & vbTab
		listServices = listServices & Chr(34) & objService.StartName & _
			Chr(34) & vbTab
		listServices = listServices & vbCRLF
		
		For counter = 1 to Len(listServices) 
			strTemp = strTemp & Chr(Asc(Mid(listServices, counter, 1)))
		Next
		listServices = strTemp
		strTemp = ""
		
	Next
	
	Set objWMIService = Nothing
	Set colListOfServices = Nothing
	Set objService = Nothing
End Function

'######################################################
'# Function Name: getShares
'# Parameters: 
'#	strDirPerms - Intput parameter containing the output file path
'# Return Value: 
'#	None
'#
'# Description:
'#	Returns the shares and share information
'######################################################
Function getShares(ByVal strDirPerms)
	On Error Resume Next

	Dim strComputer, objWMIService, colShares, objSecurityDescriptor, strMask
	Dim objShare, objShareSecSetting, objAE, arrDACL, strTrustee, objTrustee, _
		strPerms, strPermType, strAceType
	Dim intFirstTime

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")

	getShares = "Name" & vbTab & "Path" & vbTab & "Caption" & vbTab & "Type" &_
		vbTab &	"Trustee" & vbTab & "ACE Type" & vbTab & "Permissions" & vbCRLF

	For Each objShare in colShares
		strPerms = ""
		If(objShare.Type = 0) Then
			Set objShareSecSetting = GetObject( _
				"winmgmts:Win32_LogicalShareSecuritySetting.Name='" &_
				objShare.Name & "'")
			intFirstTime = objShareSecSetting.GetSecurityDescriptor _
				(objSecurityDescriptor)
			arrDACL = objSecurityDescriptor.DACL

			intFirstTime = 1

			For Each objAE In arrDACL
				If(intFirstTime = 0) Then
					strPerms = strPerms & vbCRLF & vbTab & vbTab & vbTab &_
						vbTab
				Else
					intFirstTime = 0
				End If

				strMask = objAE.AccessMask
				strAceType = objAE.AceType

				If(strAceType = 0) Then
					strPermType = "Allow"
				Else
					strPermType = "Deny"
				End If

				Select Case strMask
					Case 1179817
						strPermType = strPermType & vbTab & "Read"
					Case 1245631
						strPermType = strPermType & vbTab & "Change"
					Case 2032127
						strPermType = strPermType & vbTab & "Full Conrol"
				End Select

				' Get Win32_Trustee object from ACE
				Set objTrustee = objAE.Trustee
				strTrustee = objTrustee.Domain & "\" & objTrustee.Name
				strPerms = strPerms & strTrustee & vbTab & strPermType
			Next
		Else
			strPerms = "No DACL Found"
		End If

		getShares = getShares & objShare.Name & vbTab &_
			objShare.Path & vbTab & objShare.Caption & vbTab &_
			objShare.Description & vbTab & strPerms & vbTab & vbCRLF
		If (StrComp(objShare.Path, "") <> 0) And (Not IsNull(objShare.Path)) And _
			(Not IsEmpty(objShare.Path)) Then
			writeToFile strDirPerms, "Share Name: " & objShare.Name &_
				vbTab & "Share Path: "
			writeToFile strDirPerms, getDirPermissions(objShare.Path)
		End If
	
	Next
	getShares = getShares & vbCRLF
	Set objWMIService = Nothing
	Set colShares = Nothing
End Function

'######################################################
'# Function Name: getDrives
'# Parameters: 
'#	None
'# Return Value: 
'#	getDrives - A list of drives and corresponding information
'#
'# Description:
'#	Enumerates all drives and returns drive information
'######################################################
Function getDrives()
	On Error Resume Next
	Dim objFSO, Drive, strDriveType, Dtot, Dfree, Dused, Dpct, Dserial
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	getDrives = "Drive Letter" & vbTab & "Total Size" & vbTab &_
		"Free Space" & vbTab & "Used Space" & vbTab & "Percent Free" &_
		vbTab & "Volume Name" & vbTab & "Path" & vbTab & "Drive Type" &_
		vbTab & "Serial No." & vbTab & "File System" & vbTab & vbCRLF
	
	For Each Drive in objFSO.Drives
		
		Select Case Drive.DriveType
			Case "0" strDriveType = "Unknown drive type"
			Case "1" strDriveType = "Removable drive"
			Case "2" strDriveType = "Fixed drive"
			Case "3" strDriveType = "Network drive"
			Case "4" strDriveType = "CD/DVD drive"
			Case "5" strDriveType = "RAM Disk"
		End Select
		
		If Drive.IsReady Then	
			If Drive.DriveType=4 Then
				Dfree="N/A"
			ElseIf Drive.FreeSpace<1024^2 Then
				Dfree=FormatNumber(Drive.FreeSpace/1024,0) & " KB"
			ElseIf Drive.FreeSpace<10240^2 Then
				Dfree=FormatNumber(Drive.FreeSpace/(1024^2),2) & " MB"
			Else
				Dfree=FormatNumber(Drive.FreeSpace/(1024^2),0) & " MB"
			End If
  		  
			If Drive.TotalSize<1024^2 Then
				Dtot=FormatNumber(Drive.TotalSize/1024,0) & " KB"
			ElseIf Drive.TotalSize<10240^2 Then
				Dtot=FormatNumber(Drive.TotalSize/(1024^2),2) & " MB"
			Else
				Dtot=FormatNumber(Drive.TotalSize/(1024^2),0) & " MB"
			End If
    		
			Dused=Drive.TotalSize-Drive.FreeSpace
			If Dused<1024^2 Then
				Dused=FormatNumber(Dused/1024,0) & " KB"
			ElseIf Dused<10240^2 Then
				Dused=FormatNumber(Dused/(1024^2),2) & " MB"
			Else
				Dused=FormatNumber(Dused/(1024^2),0) & " MB"
			End If
   		 	
			If Drive.DriveType=4 Then
				Dpct="N/A"
			Else
				Dpct=FormatPercent(Drive.FreeSpace / Drive.TotalSize,1)
			End If
    		
			If Drive.DriveType=5 Then
				Dserial="N/A"
			Else
				Dserial=Hex(Drive.SerialNumber)
			End If		

			getDrives = getDrives & Drive.DriveLetter & vbTab & Dtot &_
				vbTab & Dfree & vbTab & Dused & vbTab & Dpct & vbTab &_
				Drive.VolumeName & vbTab & Drive.Path & vbTab &_
				strDriveType & vbTab & Dserial & vbTab & Drive.FileSystem &_
				vbCRLF
		Else
			getDrives = getDrives & Drive.DriveLetter & vbTab & " " & vbTab &_
				" " & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & " " &_
				vbTab & strDriveType & vbTab & " " & vbTab & " " & vbCRLF
		End If	
	Next
	
	Set objFSO = Nothing
	Set Drive = Nothing
End Function

'######################################################
'# Function Name: getADTrusts
'# Parameters: 
'#	strDomainName - Input parameter containing the domain name
'# Return Value: 
'#	getADTrusts - A list of trust relationships
'#
'# Description:
'#	Takes a domain name and returns the trust relationships
'######################################################
Function getADTrusts(ByVal strDomainName)
	Dim objTrustDirectionHash, objTrustTypeHash, objTrustAttrHash, objRootDSE, _
		objTrust, objTrusts, strFlag, strTrustInfo, boolTrustFound
		
	Set objTrustDirectionHash = CreateObject("Scripting.Dictionary")
	objTrustDirectionHash.Add "DIRECTION_DISABLED", 0
	objTrustDirectionHash.Add "DIRECTION_INBOUND",  1
	objTrustDirectionHash.Add "DIRECTION_OUTBOUND", 2
	objTrustDirectionHash.Add "DIRECTION_BIDIRECTIONAL", 3
	
	Set objTrustTypeHash = CreateObject("Scripting.Dictionary")
	objTrustTypeHash.Add "TYPE_DOWNLEVEL", 1
	objTrustTypeHash.Add "TYPE_UPLEVEL", 2
	objTrustTypeHash.Add "TYPE_MIT", 3
	objTrustTypeHash.Add "TYPE_DCE", 4
	
	Set objTrustAttrHash = CreateObject("Scripting.Dictionary")
	objTrustAttrHash.Add "ATTRIBUTES_NON_TRANSITIVE", 1
	objTrustAttrHash.Add "ATTRIBUTES_UPLEVEL_ONLY", 2
	objTrustAttrHash.Add "ATTRIBUTES_QUARANTINED_DOMAIN", 4
	objTrustAttrHash.Add "ATTRIBUTES_FOREST_TRANSITIVE", 8
	objTrustAttrHash.Add "ATTRIBUTES_CROSS_ORGANIZATION", 16
	objTrustAttrHash.Add "ATTRIBUTES_WITHIN_FOREST", 32
	objTrustAttrHash.Add "ATTRIBUTES_TREAT_AS_EXTERNAL", 64
	
	Set objRootDSE = GetObject("LDAP://" & strDomainName & "/RootDSE")
	Set objTrusts  = GetObject("LDAP://cn=System," & _
		objRootDSE.Get("defaultNamingContext") )
	objTrusts.Filter = Array("trustedDomain")
	getADTrusts = getADTrusts & "Trusts for " & strDomainName & ":" & vbCRLF
	
	boolTrustFound = False
	For Each objTrust In objTrusts
		strTrustInfo = ""
		boolTrustFound = True
		For Each strFlag In objTrustDirectionHash.Keys
			If objTrustDirectionHash(strFlag) = objTrust.Get("trustDirection") Then
				strTrustInfo = strTrustInfo & strFlag & " "
			End If
		Next
	
		For Each strFlag In objTrustTypeHash.Keys
			If objTrustTypeHash(strFlag) = objTrust.Get("trustType") Then 
				strTrustInfo = strTrustInfo & strFlag & " "
			End If
		Next
	
		For Each strFlag In objTrustAttrHash.Keys
			If objTrustAttrHash(strFlag) = objTrust.Get("trustAttributes") Then 
				strTrustInfo = strTrustInfo & strFlag & " "
			End If
		Next
	
		getADTrusts = getADTrusts & " " & objTrust.Get("trustPartner") &_
			" : " & strTrustInfo & vbCrLf
	Next
	
	If Not boolTrustFound Then
		getADTrusts = getADTrusts & "No trusts found!"
	End If
	
	Set objTrustDirectionHash = Nothing
	Set objTrustTypeHash = Nothing
	Set objTrustAttrHash = Nothing
	Set objRootDSE = Nothing
	Set objTrust = Nothing
	Set objTrusts = Nothing
End Function

'######################################################
'# Function Name: retrieveGPODirs
'# Parameters: 
'#	strDir - Input parameter containing the path of the output files
'# Return Value: 
'#	None
'#
'# Description:
'#	Takes the path of the output files and copies the GPOs to the specified path
'######################################################
Sub retrieveGPODirs(ByVal strDir)
	On Error Resume Next
	Dim objFSO, objFolder
	Dim objShell, objEnv

	Set objShell = CreateObject("WScript.Shell")
	Set objEnv = objShell.Environment("Process")

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If(objFSO.FolderExists(objFSO.GetFolder(objEnv.Item("SystemRoot") &_
		"\SYSVOL\sysvol\"))) Then
		Set objFolder = objFSO.GetFolder(objEnv.Item("SystemRoot") &_
			"\SYSVOL\sysvol\")
		objFolder.Copy(".\" & strDir)
	Else
		MsgBox "ERROR! Group Policy Objects / Templates are not in default" &_
			" location (" & objEnv.Item("SystemRoot") & "\SYSVOL\sysvol\)." &_
			vbCRLF & vbCRLF & "Please manually retrieve GPOs.", _
			0, strScriptNameVer
		writeToFile strErrorLog, "Error retrieving Group Policy Objects " &_
			"directory. Directory not in default location (" &_
			objEnv.Item("SystemRoot") & "\SYSVOL\sysvol\)." &_
			"Manually retrieve Group Policy Objects / Templates if required." &_
			vbCRLF & vbCRLF
	End If

	Set objEnv = Nothing
	Set objShell = Nothing
	Set objFSO = Nothing
	Set objFolder = Nothing
End Sub

'######################################################
'# Function Name: analyzeAuditandUserRights
'# Parameters: 
'#	strDirName - Input parameter containing the path of the output files
'# Return Value: 
'#	None
'#
'# Description:
'#	Takes the path of the output files and analyzes the secedit.sdb file copied
'######################################################
Sub analyzeAuditandUserRights(ByVal strDirName)
	On Error Resume Next
	Dim objShell, intRetVal
	Set objShell = CreateObject("WScript.Shell")
	intRetVal = objShell.Run("%comspec% /c cd """ & strDirName &_
		""" && secedit /analyze /cfg ../AuditandUserRights.inf " &_
		"/db AuditandUserRights.sdb", 7, True)

	'WScript.StdOut.WriteLine "Run return value " & intRetValue
	
	Set objShell = Nothing
End Sub

'######################################################
'# Function Name: writeToFile
'# Parameters: 
'#	strPathFileName - Input parameter containing path of the file to write to
'#      strToWrite - Input parameter containing the text to write
'# Return Value: 
'#	None
'#
'# Description:
'#	Writes the specified text to the specified output file
'######################################################
Sub writeToFile(ByVal strPathFileName, ByVal strToWrite)
	On Error Resume Next
	Const ForAppending = 8
	Dim objFSO, objLogFile

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objLogFile = objFSO.OpenTextFile(strPathFileName, ForAppending, True)
	If(Err.Number = vbEmpty) Then
		objLogFile.Write(strToWrite)
	Else
		objLogfile.Write("Unrecognized Character")
	End If
	objLogFile.Close
	Set objFSO = Nothing
	Set objLogFile = Nothing
End Sub