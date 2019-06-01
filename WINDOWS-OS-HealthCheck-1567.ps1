#
# WINDOWS-OS-HealthCheck-1567.ps1
#
# A powershell script has been developed to do health checks of multiple Windows servers, there are 2 versions one for doing healthcheck through HPSA, and other to be run on a jumphost.
#
# The script will check for following things..
#
# 1. Is Server pingable (not included in HPSA version)
# 2. Is Server able to do RDP (not included in HPSA version)
# 3. Check and report if any automatic services are not running.
# 4. Check uptime of server.
# 5. Report Drives of the server with its free space %.
# 6. Check and Report Network card status.
# 7. Number of Patches installed in past 24 hours.
# 8. Report the hotfix name installed in past 24 hours.
#
# Developer: Naveen Maheshwara

$array = @()

$row = New-Object -TypeName PSObject

$status = @()

$service = Get-WmiObject -Class Win32_service | select DisplayName, StartMode, State | Where-Object {$_.StartMode -eq 'Auto'}

foreach ($item in $service){
    $getstatus = if($item.StartMode -eq 'Auto' -and $item.State -eq 'Running') {'Running'} else {$item.DisplayName}
    $status += $getstatus
    }

$notrunning = $status | Where-Object {$_ -notcontains 'Running'}
$finaloutput = if ($status -ne 'Running') {$notrunning} else {'All services Running'}

$row | Add-Member -MemberType NoteProperty -Name Server_Service -Value ($finaloutput -Join ", ")

$boot = Get-WmiObject -Class win32_operatingsystem
$bootime = $boot.ConvertToDateTime($boot.LocalDateTime) - $boot.ConvertToDateTime($boot.LastBootUpTime)
$days = $bootime.Days
$hours = $bootime.Hours
$minute = $bootime.Minutes
$uptime = Write-Output ("$days" + "days" + " " + "$hours" + "hours"+ " " + "$minute" + "minutes" )

$row |Add-Member -MemberType NoteProperty -Name Uptime -Value $uptime

$drive = Get-WmiObject Win32_logicaldisk | Where-Object {$_.drivetype -eq 3}

$drive_free =  $drive.freespace
$drive_total = $drive.size

$drive_space = @()
$drive_percent = foreach ($a in $drive_free) {
$ind = [array]::IndexOf($drive_free, $a)
$pen = New-Object -TypeName PSObject
$volume = Write-Output (($a/$drive_total[$ind])*100)
$percent = [math]::round($volume)

$pen | Add-Member -MemberType NoteProperty -Name percent -Value $percent

$drive_space += $pen
}
$drive_space = $drive_space.percent

$drive_out = $drive.deviceID
$freespace = @()
    foreach($i in $drive_out){
    $index = [array]::IndexOf($drive_out, $i)
    $column = New-Object -TypeName PSObject
    $value = Write-Output $i, $drive_space[$index]
    $column | Add-Member -MemberType NoteProperty -Name output -Value $value

    $freespace += $column
    }
    $freespace.output

$row | Add-Member -MemberType NoteProperty -Name Drive_Check -Value ($freespace.output -join " ")

$date = (Get-Date).AddHours(-24)

$hotfix = gwmi Win32_QuickFixEngineering | Where-Object {$_.InstalledOn -gt $date}

$count = $hotfix | Measure-Object
$hotfixID = $hotfix.hotfixid

$row | Add-Member -MemberType NoteProperty -Name Patch_Installed_Past_24hours -Value $count.Count
$row | Add-Member -MemberType NoteProperty -Name Hotfix_Name -Value ($hotfixID -join ", ")

$network = Get-WmiObject win32_networkadapter
$n_name = $network.netconnectionid | Where-Object {$_ -ne $null}

$n_status = $network.netconnectionstatus | Where-Object {$_ -ne  $null}

$state = foreach($net in $n_status) {if ($net -eq 2) {'-CONNECTED,'} else {'-CHECK_STATUS,'}}

$n_array = @()
foreach ($n in $n_name){
$n_index = [array]::IndexOf($n_name, $n)
$n_output = Write-Output $n, $state[$n_index]

$n_obj = New-Object -TypeName PSObject
$n_obj | Add-Member -MemberType NoteProperty -Name State -Value $n_output
$n_array += $n_obj
}

$row | Add-Member -MemberType NoteProperty -Name Network_Status -Value ($n_array.state -join " ")

$array += $row
$array
