$Computer = [ADSI]"WinNT://$Env:COMPUTERNAME,Computer"
$LocalAdmin = $Computer.Create("User", "<Username>")
$LocalAdmin.SetPassword("<Password>")
$LocalAdmin.SetInfo()
$LocalAdmin.FullName = "<user Description Name>"
$LocalAdmin.SetInfo()
$LocalAdmin.Description = "Wintel 24x7"
$LocalAdmin.SetInfo()
$LocalAdmin.UserFlags = 65536 #
$LocalAdmin.SetInfo()
NET LOCALGROUP "Administrators" "<username>" /add 
