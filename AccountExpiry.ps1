#Up until recently a bunch of our AZ accounts in DXC managed domains had no expiry on them. It seems change went through recently which has now set all accounts to have an expiry. If you had an account that you thought previously had no expiry date I would advise you to check. An easy way to check your accounts via HPSA:
#
#Expand Device Groups > Public > CC_Midrange_Server > DC_NODE_TWOS
#Run the following Ad hoc  PowerShell script on the servers listed in this group and copy output to EXCEL or whatever is convenient for you:
#For anyone interested, here are bunch of parameters you can enter to add to the script to get further information on an account in AD, just put then with the curly brackets after the $Props = [ordered]@{ section :
#
#'Domain' = $Dom;
#'Name' = $UserExists.Name;
#'DisplayName' = $UserExists.DisplayName;
#'User ID' = $Usr;
#'SAM Account Name' = $UserExists.SamAccountName;
#'SID' = $UserExists.objectSid;
#'OU Location' = $UserExists.DistinguishedName;
#'Created' = $UserExists.Created;
#'Last Modified' = $UserExists.Modified;
#'Description' = $UserExists.Description;
#'Enabled' = $UserExists.Enabled;
#'Account Expiration Date' = $UserExists.AccountExpirationDate;
#'LastLogonDate' = $UserExists.LastLogonDate;
#'LockedOut' = $UserExists.LockedOut;
#'AccountLockoutTime' = $UserExists.AccountLockoutTime;
#'BadLogonCount' = $UserExists.BadLogonCount;
#'Password Expired' = $UserExists.PasswordExpired;
#'LastBadPasswordAttempt' = $UserExists.LastBadPasswordAttempt;
#'badPwdCount' = $UserExists.badPwdCount;
#'PasswordLastSet' = $UserExists.PasswordLastSet;
#'PasswordNeverExpires' = $UserExists.PasswordNeverExpires;
#'CannotChangePassword' = $UserExists.CannotChangePassword;
#'MemberOf' = $GrpMembership.Name;


$UserID = @("INSERT ID HERE")

$Comptr = gwmi win32_computersystem -ComputerName localhost
$Dom = $Comptr.Domain
ForEach ($Usr in $UserID) {
Try {
$UserExists = Get-AdUser -Identity $Usr -Properties *
}
Catch {
$UserExists = $null
}
if ($UserExists) {
$GrpMembership = Get-ADPrincipalGroupMembership -Identity $Usr
$Props = [ordered]@{
'Domain' = $Dom;
'Account Expiration Date' = $UserExists.AccountExpirationDate;
}
$obj = New-Object -TypeName PSObject -Property $props
Write-Output $obj | Format-List
}
Else {
Write-Host "The User or Service Account '$Usr' from the Domain '$Dom' Does Not Exist"
}
}

