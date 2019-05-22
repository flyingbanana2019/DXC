Import-Module ActiveDirectory

function Import-Excel
{
  param (
    [string]$FileName,
    [string]$WorksheetName,
    [bool]$DisplayProgress = $true
  )

  if ($FileName -eq "") {
    throw "Please provide path to the Excel file"
    Exit
  }

  if (-not (Test-Path $FileName)) {
    throw "Path '$FileName' does not exist."
    exit
  }

  $FileName = Resolve-Path $FileName
  $excel = New-Object -com "Excel.Application"
  $excel.Visible = $false
  $workbook = $excel.workbooks.open($FileName)

  if (-not $WorksheetName) {
    Write-Warning "Defaulting to the first worksheet in workbook."
    $sheet = $workbook.ActiveSheet
  } else {
    $sheet = $workbook.Sheets.Item($WorksheetName)
  }
  
  if (-not $sheet)
  {
    throw "Unable to open worksheet $WorksheetName"
    exit
  }
  
  $sheetName = $sheet.Name
  $columns = $sheet.UsedRange.Columns.Count
  $lines = $sheet.UsedRange.Rows.Count
  
  Write-Warning "Worksheet $sheetName contains $columns columns and $lines lines of data"
  
  $fields = @()
  
  for ($column = 1; $column -le $columns; $column ++) {
    $fieldName = $sheet.Cells.Item.Invoke(1, $column).Value2
    if ($fieldName -eq $null) {
      $fieldName = "Column" + $column.ToString()
    }
    $fields += $fieldName
  }
  
  $line = 2
  
  
  for ($line = 2; $line -le $lines; $line ++) {
    $values = New-Object object[] $columns
    for ($column = 1; $column -le $columns; $column++) {
      $values[$column - 1] = $sheet.Cells.Item.Invoke($line, $column).Value2
    }  
  
    $row = New-Object psobject
    $fields | foreach-object -begin {$i = 0} -process {
      $row | Add-Member -MemberType noteproperty -Name $fields[$i] -Value $values[$i]; $i++
    }
    $row
    $percents = [math]::round((($line/$lines) * 100), 0)
    if ($DisplayProgress) {
      Write-Progress -Activity:"Importing from Excel file $FileName" -Status:"Imported $line of total $lines lines ($percents%)" -PercentComplete:$percents
    }
  }
  $workbook.Close()
  $excel.Quit()
}


$Users = import-excel  "C:\Userlists.xlsx" 

#start of ad creation
       
foreach ($User in $Users)    
{

    $Displayname = $User.'Firstname' + " " + $User.'Lastname'            
    $UserFirstname = $User.'Firstname'            
    $UserLastname = $User.'Lastname'            
    $OU = $User.'OU'            
    $SAM = $User.'SAM'            
    $UPN = $User.'Firstname' + "." + $User.'Lastname' + "@" + $User.'Maildomain'            
    $Description = $User.'Description'            
    $Password = $User.'Password'  
    $Logon = $User.'Logon'
    $HomeDir = $User.'Homedirectory'
    $Dept = $User.'Department/CostCenter'
    $Company = $User.'Company' 
    $Office = $User.'Office' 
    $Division = $User.'Division'
    $OfficePhone = $User.'Telephone Number'  
    $Manager = $User.'Manager'
    $MobilePhone = $User.'MobilePhone'
    $Manager = $User.'Manager'
    $StreetAddress = $User.'StreetAddress'
    $POBox = $User.'POBox'
    $City = $User.'City'
    $State = $User.'State'
    $PostalCode = $User.'PostalCode'
    $Country = $User.'Country'
    $hdir = $user.'hdir'
    $User.'Notes'
    $User.'FixedGroup' | Out-File C:\fixedgroups.txt
    $User.'AddGroup'  | Out-File C:\groups.txt
    
    
    New-ADUser `
    -Name "$Displayname" `
    -DisplayName "$Displayname" `
    -SamAccountName $SAM `
    -UserPrincipalName $UPN `
    -GivenName "$UserFirstname" `
    -Surname "$UserLastname" `
    -Description "$Description" `
    -Title "$Description" `
    -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Enabled $true `
    -ChangePasswordAtLogon $true `
    -HomeDrive "H:" `
    -HomeDirectory $HomeDir `
    -Company $Company `
    -Department $Dept `
    -Division $Division `
    -Office $Office `
    -Manager $Manager `
    -OfficePhone $OfficePhone `
    -HomePhone $OfficePhone `
    -MobilePhone $MobilePhone `
    -StreetAddress $StreetAddress `
    -PostalCode $PostalCode `
    -POBox $POBox `
    -City $City `
    -State $State `
    -Country $Country `
    -Server downergroup.internal `
    -Path "$OU" 
    

 Set-ADUser $SAM -ScriptPath $Logon 
 Set-ADUser $SAM -Replace @{info = $User.'Notes'}
  

#Fixed Security/Distribution Group

 $User1 = Get-ADUser -Identity "$SAM"
 Add-ADGroupMember -Identity "" -Member "$User1"
 Add-ADGroupMember -Identity "" -Member "$User1"
  
 $Groups = Get-Content C:\groups.txt
 $FixedGroups = Get-Content C:\fixedgroups.txt
   
  foreach ($group in $Groups)
  {
   Add-ADGroupMember -Identity $group -Member $SAM
  }
  
  foreach ($fgroup in $FixedGroups)
  {
   Add-ADGroupMember -Identity $fgroup -Member $SAM
  }
    
   mkdir "$hdir$($SAM)"
   icacls "$hdir$($SAM)" /grant "$($SAM):(OI)(CI)F"


}
      

  
    

   



