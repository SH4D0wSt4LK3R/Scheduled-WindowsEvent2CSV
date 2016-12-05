
$starttime = get-date
$filename = $env:computername + "-" + $starttime
$filename = $filename.replace("/","-").replace(":","-")
[string]$applasttime = "AppLastTimestamp.txt"
[string]$apptimepath = "C:\Event\" + $applasttime
[string]$seclasttime = "SecLastTimestamp.txt"
[string]$sectimepath = "C:\Event\" + $seclasttime
[string]$syslasttime = "SysLastTimestamp.txt"
[string]$systimepath = "C:\Event\"+ $syslasttime

#Check for and create last run file if it doesn't exist for app logs. Populate default values.
If((Test-Path -path $apptimepath) -ne $true){New-Item $apptimepath -type file -force -value "01/01/2000 00:00:00 AM"}
ELSEIF((Get-Content $apptimepath) -eq $Null){[datetime]$appoldlastrun = "01/01/2000 00:00:00 AM"}
ELSE{[DateTime]$appoldlastrun = Get-Content -Path $apptimepath}
If($error.count -ne 0){
    [DateTime]$appoldlastrun = "01/01/2000 00:00:00 AM"
    $error.clear()
}

#Check for and create last run file if it doesn't exist for sec logs. Populate default values.
If((Test-Path -path $sectimepath) -ne $true){New-Item $sectimepath -type file -force -value "01/01/2000 00:00:00 AM"}
ELSEIF((Get-Content $sectimepath) -eq $Null){[datetime]$secoldlastrun = "01/01/2000 00:00:00 AM"}
ELSE{[DateTime]$secoldlastrun = Get-Content -Path $sectimepath}
If($error.count -ne 0){
    [DateTime]$secoldlastrun = "01/01/2000 00:00:00 AM"
    $error.clear()
}

#Check for and create last run file if it doesn't exist for sys logs. Populate default values.
If((Test-Path -path $systimepath) -ne $true){New-Item $systimepath -type file -force -value "01/01/2000 00:00:00 AM"}
ELSEIF((Get-Content $systimepath) -eq $Null){[datetime]$sysoldlastrun = "01/01/2000 00:00:00 AM"}
ELSE{[DateTime]$sysoldlastrun = Get-Content -Path $systimepath}
If($error.count -ne 0){
    [DateTime]$sysoldlastrun = "01/01/2000 00:00:00 AM"
    $error.clear()
}

#Access and grab the Application event logs
$applog = get-winevent -filterhashtable @{Logname="application";StartTime=$appoldlastrun;EndTime=$starttime}
$applastrun = $applog[0].timecreated

#Access and grab the Security event logs
$seclog = get-winevent -filterhashtable @{Logname="security";StartTime=$secoldlastrun;EndTime=$starttime}
$seclastrun = $seclog[0].timecreated

#Access and grab the System event logs
$syslog = get-winevent -filterhashtable @{Logname="system";StartTime=$sysoldlastrun;EndTime=$starttime}
$syslastrun = $syslog[0].timecreated

#Set new applastrun value
ri $apptimepath
Add-Content $applastrun -path $apptimepath

#Set new seclastrun value
ri $sectimepath
Add-Content $seclastrun -path $sectimepath

#Set new syslastrun value
ri $systimepath
Add-Content $syslastrun -path $systimepath

#Create CSV for Application event log
$appout = "C:\Event\" + $filename + "-app.csv"
$applog | Export-csv -path $appout -notypeinformation

#Create CSV for Security event log
#$secout = "C:\Event\" +$filename + "-sec.csv"
#$seclog | Export-csv -path $secout -notypeinformation

#Create CSV for System event log
$sysout = "C:\Event\" + $filename + "-sys.csv"
$syslog |Export-Csv -path $sysout -NoTypeInformation
