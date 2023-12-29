Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003

strComputer = "."

Set objShell = CreateObject("WScript.Shell")
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

'ini untuk disable NBT-NS
strTargetkey = ""
oReg.EnumKey HKEY_USERS, strTargetkey, arrSubkeys

If IsArray(arrSubkeys) Then 
	For Each strSubkey In arrSubkeys
		'ini untuk disableRegistryTools
		oReg.CreateKey HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies"
		oReg.CreateKey HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies\System"
		oReg.SetDWORDValue HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 1	
		'ini untuk disableCMD		
		oReg.CreateKey HKEY_USERS, strSubkey & "\Software\Policies\Microsoft\Windows"
		oReg.CreateKey HKEY_USERS, strSubkey & "\Software\Policies\Microsoft\Windows\System"
		oReg.SetDWORDValue HKEY_USERS, strSubkey & "\Software\Policies\Microsoft\Windows\System", "DisableCMD", 1
		'ini untuk disable run from File Explorer, tidak efektif kalau user mengaktifkan dari CMD
		'Computer\HKEY_USERS\S-1-5-21-1521604063-934003188-285078316-1001\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun
		oReg.CreateKey HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
		oReg.SetDWORDValue HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisallowRun", 1
		oReg.CreateKey HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun"
		oReg.SetStringValue HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "1", "powershell.exe"
		oReg.SetStringValue HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "2", "powershell_ise.exe"
		oReg.SetStringValue HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "3", "pwsh.exe"
		oReg.SetStringValue HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "4", "psexec.exe"
		oReg.SetStringValue HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "4", "psexec64.exe"
		oReg.SetStringValue HKEY_USERS, strSubkey & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "4", "wmic.exe"
		WScript.echo strSubkey
	Next
End If
