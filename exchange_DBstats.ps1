<#
.SYNOPSIS
    Script to export all DB and Disk size statistics to a csv and sending a html-based mail report (can be opt out).
    This csv output can be used for Excel import, Exchange storage growth reviews, move-mailbox actions and sizing forecasts.

.PARAMETER NoMail
    <optional> An email report will NOT be sent

.EXAMPLE
    .\exchange_DBstats.ps1 [-NoMail]

.VERSIONS
    V1.0 03.06.2025 - Initial Version
    V1.1 04.06.2025 - Change archive collection method to save a lot of time
    V1.2 05.06.2025 - no parameter CSVFileName anymore
    V1.3 05.06.2025 - take into account of UNLIMITED quotas
    V1.5 15.07.2025 - -in instead of -eq, Archive count corrected
    V1.6 29.07.2025 - adding SUM line
    V1.7 30.07.2025 - changed HTML code, bold TOTAL line
    V2.0 23.09.2025 - collecting disk sizes to be compared with disk thresholds, added table of critical disks and/or dbs, if there are some, so move-mailbox sources can be found easily

.AUTHOR/COPYRIGHT:
    Steffen Meyer
    Cloud Solution Architect
    Microsoft Deutschland GmbH
#>

[CmdletBinding()]
Param(
     [parameter(Position=0,HelpMessage='No MailReport')]
     [switch]$NoMail
     )

$version = "V2.0_23.09.2025"

$now = Get-Date -Format G

Function Set-HighlightErrors
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory,ValueFromPipeline)]
        [string]$Line,

        [Parameter(Mandatory)]
        [string]$CSSErrorClass,

        [Parameter(Mandatory)]
        [string]$CSSWarnClass,

        [Parameter(Mandatory)]
        [string]$CSSPassClass,

        [Parameter(Mandatory)]
        [string]$CSSInfoClass,

        [Parameter(Mandatory)]
        [string]$ERRORValue,

        [Parameter(Mandatory)]
        [string]$WARNValue,

        [Parameter(Mandatory)]
        [string]$PASSValue,
        
        [Parameter(Mandatory)]
        [string]$INFOValue
    )
    Process
    {
        $Line = $Line.Replace("<td>$INFOValue</td>","<td class=""$CSSInfoClass"">$INFOValue</td>")
        $Line = $Line.Replace("<td>$PASSValue</td>","<td class=""$CSSPassClass"">$PASSValue</td>")
        $Line = $Line.Replace("<td>$WARNValue</td>","<td class=""$CSSWarnClass"">$WARNValue</td>")
        $Line = $Line.Replace("<td>$ERRORValue</td>","<td class=""$CSSErrorClass"">$ERRORValue</td>")
        Return $Line
    }
}

try
{
    $ScriptPath = Split-Path -parent $MyInvocation.MyCommand.Path -ErrorAction Stop
}
catch
{
    Write-Host "`nDo not forget to save the script!" -ForegroundColor Red
}

write-host "`n-------------------------------------------------------------------------------------" -foregroundcolor green
write-host   " This report collects important informations of all Exchange Databases in this       " -foregroundcolor green
write-host   " organization, e.g. where it is currently mounted, Disk sizes, Disk types, DB sizes, " -foregroundcolor green
write-host   " DB whitespaces (root tables), Mailbox counts, Archive counts, DB quotas, paths,     " -foregroundcolor green
write-host   " circular logging and more to a common *.csv-file. Additionally, it sends an HTML-   " -ForegroundColor green
write-host   " based email report (can be prevented) with all important numbers. In this report,   " -ForegroundColor green
write-host   " exceeded numbers are highlighted (based on your thresholds in settings.cfg).        " -ForegroundColor green
write-host   "-------------------------------------------------------------------------------------" -foregroundcolor green

Write-Host "`nScriptversion: $version"
Write-Host "Script started: $now"

#settings.cfg
if (Test-Path -Path "$ScriptPath\settings.cfg")
{
    $Config = Get-Content -Path "$ScriptPath\settings.cfg"
}
else
{
    write-host "`nThe required file SETTINGS.CFG is missing. Add the file to ensure a working SCRIPT/MAIL REPORT." -ForegroundColor Magenta
    Return
}

#Company/Environmentname
$Company = ($config | Where-Object {$_.StartsWith("Company")}).split('=',2)[1]
    
#Thresholds for Highlighting
$CritFreeSpace = ($config | Where-Object {$_.StartsWith("CritFreeSpace")}).split('=',2)[1]
$WarnFreeSpace = ($config | Where-Object {$_.StartsWith("WarnFreeSpace")}).split('=',2)[1]
$CritFreePercent = ($config | Where-Object {$_.StartsWith("CritFreePercent")}).split('=',2)[1]
$WarnFreePercent = ($config | Where-Object {$_.StartsWith("WarnFreePercent")}).split('=',2)[1]
$CritDBSize = ($config | Where-Object {$_.StartsWith("CritDBSizeInGB")}).split('=',2)[1]
$WarnDBSize = ($config | Where-Object {$_.StartsWith("WarnDBSizeinGB")}).split('=',2)[1]
$CritMBXCount = ($config | Where-Object {$_.StartsWith("CritMailboxCountperDB")}).split('=',2)[1]
$WarnMBXCount = ($config | Where-Object {$_.StartsWith("WarnMailboxCountperDB")}).split('=',2)[1]

#Mailoptions
$MailServer = ($config | Where-Object {$_.StartsWith("MailServer")}).split('=',2)[1]
$MailFrom = ($config | Where-Object {$_.StartsWith("MailFrom")}).split('=',2)[1]
$Recipients = ($config | Where-Object {$_.StartsWith("Recipients")}).split('=',2)[1]
[string[]]$MailTo = $Recipients.Split(',')
$MailSubject = "Exchange Database(s) & Disk(s) Report - $($Company) - $now"

#CsvFileName/OutputFile
$CsvFileName = $Company + '_' + $(get-date -f yyyyMMdd) + '.csv'
$OutputFile = Join-Path $ScriptPath -ChildPath $CsvFileName

#HTML/CSS
$Style =   "<html>
           <head>
           <style type=$("text/css")>
           table {border-collapse:collapse; border-spacing:0; margin:0}
           div, td {padding:0;}
           div {margin:0 !important;}
           BODY{font-family: Arial, sans-serif ;font-size: 8px;}
           H1{font-size: 16px;}
           H2{font-size: 14px;}
           H3{font-size: 12px;}
           H4{font-size: 10px;}
           H5{font-size: 9px;}
           TABLE{font-size: 8pt;}
           TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
           TD{border: 1px solid black; padding: 5px; }
           td.pass{background: #7FFF00;}
           td.warn{background: #FFE600;}
           td.fail{background: #FF0000; color: #ffffff;}
           td.info{background: #85D4FF;}
           </style>
           </head>"

#HTML Body headline
$Mailbody = "<body>
           <H2 align=""left"">Exchange Database(s) & Disk(s) Report - $Company - $(Get-Date $now -Format 'dd.MM.yyyy HH:mm')</H2>"

#HTML/CSS highlighting
$Pass = "OK"
$Warn = "WARNING"
$Fail = "CRITICAL"
$Info = "SWITCHED"

#Start script

#Check if Exchange SnapIn is available and load it
if (!(Get-PSSession).ConfigurationName -eq "Microsoft.Exchange")
{
    if ((Get-PSSnapin -Registered).name -contains "Microsoft.Exchange.Management.PowerShell.SnapIn")
    {
        Write-Host "`nLoading the Exchange Powershell SnapIn..." -ForegroundColor Yellow
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
        . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
        Connect-ExchangeServer -auto -AllowClobber
    }
    else
    {
        write-host "`nExchange Management Tools are not installed. Run the script on a different machine." -ForegroundColor Red
        Return
    }
}

#Detect, where the script is executed
if (!(Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue))
{
    write-host "`nATTENTION: Script is executed on a non-Exchangeserver..." -ForegroundColor Cyan
}

#Fetching Archives
$Archives = $null
Set-ADServerSettings -viewentireforest $true

Write-Host "`nFetching all ARCHIVE mailboxes directly at the beginning (to save time), but this may take a while..." -foregroundcolor Cyan

#Define special types of mailboxes (Archives)
$types = "-Archive"

ForEach ($type in $types)
{
    try
    {
        $Archives += Invoke-Expression "Get-Mailbox $type -resultsize unlimited -ignoredefaultscope -ErrorAction Stop -WarningAction SilentlyContinue"
    }
    catch
    {
        Write-Host "We found no mailboxes with parameter $($type -replace ('-','')).`n" -ForegroundColor Red
    }
}
Write-Host "`nWe found $(($Archives).count) archives..." -ForegroundColor White

#Fetching databases
Write-Host "`nFetching all databases..." -foregroundcolor Cyan
$Databases = Get-MailboxDatabase -status | sort name

#Counter for Progressbar
$DBCount = ($Databases).count

Write-Host "`nWe found $($DBCount) databases..." -ForegroundColor White

#Collecting DB stats
$Results = @()
$CritResults = @()
$Count = 0
$DBSizeTotalInGB = $null
$WhiteSpaceTotalInGB = $null
$NetCapaTotalInGB = $null
$MBXTotal = $null
$ARCHTotal = $null
$PFMBXTotal = $null

foreach ($Database in $Databases)
{
    #Reset to null
    $ActPref = $null
    $DiskInfo = @()
    $DBSizeInGB = $null
    $WhiteSpaceInGB = $null
    $NetCapaInGB = $null
    $MBX = $null
    $ARCH = $null
    $PFMBX = $null
    $SUM = $null
     
    #ProgressBar
    $Count++
    $Activity = "Analyzing Databases... [$($Count)/$($DBCount)]"
    $Message = "Getting Statistics for Database: $($Database)"
    Write-Progress -Status $Message -Activity $Activity -PercentComplete (($Count/$DBCount)*100)
            
    #Database stats
    $MountedOn = (($Database).MountedOnServer -split "\.")[0]
    $ActPref = (($Database).activationpreference | where key -in (($Database).MountedOnServer -split "\.")[0]).value
    $EDBFilePath = $Database.EdbFilepath
    
    #Read out volume and its sizes remotely
    $ScriptBlock = {
        Param($EDBFilePath)
        
        $Folder = Split-Path $EDBFilePath -Parent

        try
        {
            $clusvols = (Get-Volume -FilePath $Folder -ErrorAction Stop).UniqueId
            foreach ($clusvol in $clusvols)
            {
                $clusvol = $clusvol.replace('\', '\\')

                if (Get-CimInstance -ClassName Win32_Volume -ComputerName $database.MountedOnServer -Filter "DeviceID='$clusvol'")
                {
                    [PSCustomObject]@{
                    Vol = Get-CimInstance -ClassName Win32_Volume -ComputerName $database.MountedOnServer -Filter "DeviceID='$clusvol'" -ErrorAction Stop
                    }
                    
                }
            }
        }
        catch
        {
        Write-Host "`nWe couldn't collect disk values for disk ""$root"" ($($database.name)) on server $($MountedOn)." -ForegroundColor Red
        }
    }
        
    $DiskInfo = Invoke-Command -ComputerName $Database.MountedOnServer -ScriptBlock $ScriptBlock -ArgumentList $EDBFilePath
    
    #Collect disk values from object
    $Root = $Diskinfo.Vol.Label
    $PathType = if ($Diskinfo.Vol.Name -match '^[A-Z]:\\$') { "DriveLetter" } else { "MountPoint" }
    $TotalGB = [math]::round(($DiskInfo.Vol.Capacity / 1GB),1)
    $FreeGB = [math]::round(($DiskInfo.Vol.FreeSpace / 1GB),1)
    $FreePercent = [math]::round((($DiskInfo.Vol.FreeSpace * 100 / $DiskInfo.Vol.Capacity)),1)
    
    #Collect DB values
    $DBSizeInGB = $Database.databasesize.toGB()
    $WhiteSpaceInGB = $Database.availablenewmailboxspace.toGB()
    $NetCapaInGB = $DBSizeInGB - $WhiteSpaceInGB
    
    #Count Mailboxes per DB   
    try
    {
        $MBX = (Get-Mailbox -resultsize unlimited -database $database.name -ErrorAction Stop -WarningAction SilentlyContinue).count
    }
    catch
    {
        Write-Host "`nWe couldn't collect Mailboxes for database $($database.name)."
    }
    
    #Count archive mailboxes per DB
    $ARCH = $Archives.archivedatabase.Name -like $Database.Name

    if ($ARCH)
    {
        $ARCH = $ARCH.Count
    }
    else
    {
        $ARCH = "0"
    }

    #Public Folder mailboxes per DB
    try
    {
        $PFMBX = (Get-Mailbox -resultsize unlimited -publicfolder -database $database.name -ErrorAction Stop -WarningAction SilentlyContinue).count
    }
    catch
    {
        Write-Host "`nWe couldn't collect PublicFolder mailboxes for database $($database.name)."
    }
    
    $SUM = $MBX + $ARCH + $PFMBX

    #Filling up a sorted array with all values    
    $data = [ordered] @{
        Database = $Database.Name
        DAG = $Database.MasterServerOrAvailabilityGroup.Name
        MountedOn = $MountedOn       
        "On ActPref 1" = if ($ActPref -gt 1) {"SWITCHED"} elseif ($ActPref -lt 1) {"CRITICAL"} else {"OK"}
        DiskLabel = $Root
        PathType = $PathType
        "TotalSize in GB" = $TotalGB
        "FreeSpace in GB" = $FreeGB
        "FreeSpace %" = $FreePercent
        FreeSpace = if ($FreeGB -lt $CritFreeSpace -or $FreePercent -lt $CritFreePercent) {"CRITICAL"} elseif ($FreeGB -lt $WarnFreeSpace -or $FreePercent -lt $WarnFreePercent) {"WARNING"} else {"OK"}            
        "Gross DBSize in GB" = $DBSizeInGB
        "Whitespace in GB" = $WhiteSpaceInGB
        "Net DBSize in GB" = $NetCapaInGB
        DBSize = if ($NetCapaInGB -gt $CritDBSize) {"CRITICAL"} elseif ($NetCapaInGB -gt $WarnDBSize) {"WARNING"} else {"OK"}            
        Mailboxes = $MBX
        Archives = $ARCH
        "PF Mailboxes" = $PFMBX             
        MBXperDB = if ($SUM -gt $CritMBXCount) {"CRITICAL"} elseif ($SUM -gt $WarnMBXCount) {"WARNING"} else {"OK"}
        CircLog = $Database.CircularLoggingEnabled
        LastFullBK = $Database.LastFullBackup
        LastIncBK = $Database.LastIncrementalBackup
        IsRecoveryDB = $Database.Recovery
        ExclfromProvi = $Database.IsExcludedFromProvisioning
        IssueWarInGB = if ($Database.IssueWarningQuota.isunlimited -eq $True) {"UNLIMITED"} else {$Database.IssueWarningQuota.value.toGB()}
        ProhSendInGB = if ($Database.ProhibitSendQuota.isunlimited -eq $True) {"UNLIMITED"} else {$Database.ProhibitSendQuota.value.toGB()}
        ProhSendRecInGB = if ($Database.ProhibitSendReceiveQuota.isunlimited -eq $True) {"UNLIMITED"} else {$Database.ProhibitSendReceiveQuota.value.toGB()}
        RecItemsWaInGB = if ($Database.RecoverableItemsWarningQuota.isunlimited -eq $True) {"UNLIMITED"} else {$Database.RecoverableItemsWarningQuota.value.toGB()}
        RecItemsInGB = if ($Database.RecoverableItemsQuota.isunlimited -eq $True) {"UNLIMITED"} else {$Database.RecoverableItemsQuota.value.toGB()}
        DelItemRetInDays = $Database.DeletedItemRetention.Totaldays
        MBXRetInDays = $Database.MailboxRetention.Totaldays
        MountDial = $Database.AutoDatabaseMountDial
        DBPath = $Database.EdbFilePath.Pathname
        LogPath = $Database.LogFolderPath.Pathname
    }
    
    #Creating "all databases" object and adding array values
    $Results += New-Object -TypeName PSObject -Property $data

    #Creating "critical db & disk" object to report separately
    if ($data.FreeSpace -eq "CRITICAL" -or $data.DBSize -eq "CRITICAL" -or $data.MBXperDB -eq "CRITICAL")
    {
        $CritResults += New-Object -TypeName PSObject -Property $data
    }

    #TOTAL numbers    
    $DBSizeTotalInGB += ($DBSizeInGB | measure -Sum).Sum
    $WhiteSpaceTotalInGB += ($WhiteSpaceInGB | measure -Sum).Sum
    $NetCapaTotalInGB += ($NetCapaInGB | measure -Sum).Sum
    $MBXTotal += ($MBX | measure -Sum).Sum
    $ARCHTotal += ($ARCH | measure -Sum).Sum
    $PFMBXTotal += ($PFMBX | measure -Sum).Sum
}
Write-Progress -Activity $Activity -Completed

#Export to CSVFile
try
{
    $Results | Export-Csv -Path $OutputFile -Encoding UTF8 -Delimiter ";" -NoTypeInformation -ErrorAction Stop

    write-host "`n--------------------------------------------------------------------------------------------------------------"
    write-host "Exchange Database statistics were successfully exported to ""$($OutputFile)""." -ForegroundColor Green
    write-host "--------------------------------------------------------------------------------------------------------------"
}
catch
{
    write-host "`n--------------------------------------------------------------------------------------------------------------"
    write-host "Exchange Database statistics couldn't be exported to ""$($OutputFile)""." -ForegroundColor Red
    write-host "--------------------------------------------------------------------------------------------------------------"
}

#Adding SUM line for HTML Output
$Results += New-Object PSObject -Property @{"Database" = "TOTAL"; "Gross DBSize in GB" = "DB Size`r`n"  + "$DBSizeTotalinGB" + " GB"; "Whitespace in GB" = "Whitespace`r`n" + "$WhiteSpaceTotalInGB" + " GB"; "Net DBSize in GB" = "DB NetCapacity`r`n" + "$NetCapaTotalInGB" + " GB"; "Mailboxes" = "$MBXTotal" + "`r`nMailboxes"; "Archives" = "$ARCHTotal" + "`r`nArchives" ; "PF Mailboxes" = "$PFMBXTotal" + "`r`nPF Mailboxes"}

#Optional: Send HTML based email report
if (!($NoMail))
{
    try
    {
        #if there are critical results, show separate table first
        if ($CritResults)
        {
            #Adding headline for "Critical disks and databases table"
            $Mailbody += "<body>
            <H3 align=""left""><font color=""#FF0000"">Critical Exchange Database(s) & Disk Size(s) Table</H3></Body>"
            
            $CritTable = $CritResults | select-object -Property "Database","DAG","MountedOn","On ActPref 1","DiskLabel","PathType","TotalSize in GB","FreeSpace in GB","FreeSpace %","FreeSpace","Gross DBSize in GB","Whitespace in GB","Net DBSize in GB","DBSize","Mailboxes","Archives","PF Mailboxes","MBXperDB" | ConvertTo-Html -Fragment | Set-HighlightErrors -CSSErrorClass fail -CSSWarnClass warn -CSSPassClass pass -CSSInfoClass info -ERRORValue $Fail -WARNValue $Warn -PASSValue $Pass -INFOValue $Info
        
            #Add CritTable to Mailbody
            $Mailbody += $CritTable
        }
        else
        {
            $Mailbody += "<body>
            <H3 align=""left""><font color=""#008000"">No Critical Exchange Database(s) & Disk Size(s) found.</H3></Body>"
        }
        
        #Adding BOLD letter/numbers in last line in HTML table, including CR
        $HtmlTable = $Results | select-object -Property "Database","DAG","MountedOn","On ActPref 1","DiskLabel","PathType","TotalSize in GB","FreeSpace in GB","FreeSpace %","FreeSpace","Gross DBSize in GB","Whitespace in GB","Net DBSize in GB","DBSize","Mailboxes","Archives","PF Mailboxes","MBXperDB" | ConvertTo-Html -Fragment | Set-HighlightErrors -CSSErrorClass fail -CSSWarnClass warn -CSSPassClass pass -CSSInfoClass info -ERRORValue $Fail -WARNValue $Warn -PASSValue $Pass -INFOValue $Info
        $HtmlTable = $HtmlTable -replace "`r`n","<br>"
        $Lines = $HtmlTable -split "`n"
        $lastTrIndex = ($Lines | Select-String -Pattern "<tr>" | Select-Object -Last 1).Linenumber - 1
        $Lines[$lastTrIndex] = $Lines[$LastTrIndex] -replace '<td>(.*?)</td>','<td><b>$1</b></td>'
        $FinalHtmlTable = $Lines -join "`n"
        
        #Adding headline for "All databases"
        $Mailbody += "<body>
            <H3 align=""left"">All Exchange Database(s) Table</H3></Body>"

        #Adding "All databases" table to Mailbody
        $Mailbody += $FinalHtmlTable

        #Finalize HTML Mailbody
        $FinalHtml = "$Style`n" +
        "$Mailbody`n" +
        "</body>`n" +
        "</html>"
        
        #Send Mail with HTML body and CSV Outputfile as attachment
        Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -BodyAsHtml $FinalHtml -Attachments $OutputFile -SmtpServer $MailServer -Encoding UTF8 -ErrorAction Stop
        
        Write-Host "`nNOTICE: Mail report was sent to $($Recipients) successfully."
    }
    catch
    {
        Write-Host "`nNOTICE: Mail report couldn't be sent." -ForegroundColor Red
    }
}
else
{
    Write-Host "`nNOTICE: An email report wasn't sent (Parameter -NoMail)." -ForegroundColor Yellow
}
#END