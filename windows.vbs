'VB Script ini dibuat untuk memudahkan proses harderning level server maupun client di CDN
'1. Non-aktifkan LLMNR
'2. Non-aktifkan NBT-NS
'3. Disable Windows Default Share $ADMIN
'4. Disable Registry Editor
'5. Disable CMD
'6. Disable Run beberapa file
'7. Disable USBSTOR
'8. Disable WiFi
'9. Hapus seBackupPrivileges an seRestorePrivileges dari Backup Operators
'10. Setting Password Policy
'11. Disable Administrator
'12. Aktifkan Software Restriction Policy

'https://infinitelogins.com/2020/11/23/disabling-llmnr-in-your-network/

Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003

strComputer = "."
strKeyPath = "Installer\Products"

Set objShell = CreateObject("WScript.Shell")
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

'Ini untuk disable LLMNR
oReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\DNSClient", "EnableMulticast", "1"

'ini untuk disable NBT-NS
strTargetkey = "SYSTEM\CurrentControlSet\services\NetBT\Parameters\Interfaces"
oReg.EnumKey HKEY_LOCAL_MACHINE, strTargetkey, arrSubkeys
If IsArray(arrSubkeys) Then 
	For Each strSubkey In arrSubkeys
		'WScript.echo strSubkey
		oReg.SetStringValue HKEY_LOCAL_MACHINE, strTargetkey &  "\" &  strSubkey, "NetbiosOptions", "2"
	Next
End If

'ini untuk disable WinHttpAutoProxySvc (4=Disable)
oReg.SetDWORDValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\WinHttpAutoProxySvc", "Start", 4

'ini untuk disable Windows Default $ADMIN Shares to block psexec.exe->PSEXESVC remote
oReg.SetDWORDValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters", "AutoShareServer", 0
oReg.SetDWORDValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\LanManServer", "AutoShareWks", 0

'net share Admin$ /delete
'net share IPC$ /delete
'net share C$ /delete
'net share D$ /delete

'ini untuk disableRegeditor
oReg.CreateKey HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
oReg.SetDWORDValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 1


'ini untuk disableCMD
oReg.CreateKey HKEY_CURRENT_USER, "SOFTWARE\Policies\Microsoft\Windows\System"
oReg.SetDWORDValue HKEY_CURRENT_USER, "SOFTWARE\Policies\Microsoft\Windows\System", "DisableCMD", 1

'ini untuk disable run from File Explorer, tidak efektif kalau user mengaktifkan dari CMD
oReg.CreateKey HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"
oReg.CreateKey HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun"
oReg.SetStringValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "1", "powershell.exe"
oReg.SetStringValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "2", "powershell_ise.exe"
oReg.SetStringValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "3", "pwsh.exe"
oReg.SetStringValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "4", "psexec.exe"
oReg.SetStringValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "4", "psexec64.exe"
oReg.SetStringValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", "4", "wmic.exe"

'ini untuk disable USBSTOR (4), dan enable kembali (3)
oReg.CreateKey HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\USBSTOR"
oReg.SetDWORDValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\USBSTOR", "Start", 4

'ini untuk disable WIFI (Win10 dan Win11)
'https://www.top-password.com/blog/turn-on-or-off-wifi-in-windows-11/
oReg.CreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\PolicyManager\default\Wifi\AllowWiFi"
oReg.SetDWORDValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\PolicyManager\default\Wifi\AllowWiFi", "Value", 0

WScript.echo "Jangan lupa remove seBackupPrivileges secara manual"
'1. Aktifkan secpol.msc
'2. Navigasi ke Security Settings -> Local Policies -> User Rights Assignment
'3. Pada Back up files and directories -> hapus Backup Operators
'4. Pada Restore files and directories -> hapus Backup Operators

WScript.echo "Jangan lupa set Password Policy secara manual"
'1. Aktifkan secpol.msc
'2. Navigasi ke Account Policies -> Password Policy
'3. - Maximum password age=42
'4. - Minimum password length=10
'5. - Minimum password length audit=10
'6. - Password must meet complexity requirement=Enable
'7. Navigasi ke Account Policies -> Account Lockout Policy
'8. - Account lockout duration= 1 minutes
'9. - Account lockout threshold= 5 invalid logon attempts'
'10.- Reset account lockout counter afer=1 minutes

WScript.echo "Jangan lupa set Disable Administrator secara manual"
'aktifkan ke command prompt
'Net User Administrator /Active:No

'https://blogs.manageengine.com/corporate/general/2018/10/25/application-whitelisting-using-software-restriction-policies.html
'Professfional Edition
WScript.echo "Jangan lupa aktifkan Software Restriction Policies secara manual"
'aktifkan ke secpol
'1. Navigasi ke Software Restriction Policies
'2. Klik kanan dan pilih New Software Restriction Policies
'3. Double klik pada Enforcement
'4. - Apply software restriction policies to the following=All software FileSystemObject.BuildPath
'5. - Apply software restriction policies to the following users=All users
'6. Double klik Security Level
'7. - Disallow -> Set as default
'8. Double klik pada Additional Rules
'9. - Tambahkan path kalau ada.