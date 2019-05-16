﻿# Purpose: gather, analyze, and export statistics based on data generated by corresponding iPerf script
# Author: Andy Jackson (andy@ku.edu)
# Date: 2/19/2019

#####
##### Edit this section
#####

# UNC path to the root folder where the iPerf reports are stored
$Path = "\\server.domain.com\share\iPerf"

# UNC path to the airwave python script
$AirwaveScript = "C:\AirWave.py"

# UNC path to the building details excel spreadsheet
$BuildingInfoPath = "C:\Building List.xlsx"

#####
##### Do not edit this section
#####

$DateTime = Get-date -Format yyyy-MM-dd_hh_mm_ss_tt
$ExportRoot = $Path + '\_Stats'
$ExportStatsReportPath = $ExportRoot + '\iPerfReport_' + $DateTime + '.xlsx'

#####
##### Helper functions
#####

function Get-AirWaveAPs
{
    param($ScriptPath,$BuildingDetails)
    $Gather = python.exe $ScriptPath
    $APs = @()
    foreach ($Device in $Gather)
    {
        if ($Device.split(',')[0].split('-')[1][0] -eq 'w'){$DeviceFunction = "Wireless AP"}
        $BuildingName = ($BuildingDetails.where({$_.ABBR -eq $Device.split('-')[0]})).Name
        switch ($Device.split(',')[0].split('-')[1][1])
        {
            A {$DeviceType = 'Wireless Access Point'}
            O {$DeviceType = 'Outdoor Wireless Access Point'}
        }
        $Properties = @{
            "Name" = $Device.split(',')[0] ;
            "IPAddress" = $Device.split(',')[1] ;
            "MAC-BGN" = $Device.split(',')[2] ;
            "MAC-AC" = $Device.split(',')[3] ;
            "LocationID" = $Device.split(',')[0].split('-')[0] ;
            "BuildingName" = $BuildingName ;
            "DeviceFunction" = $DeviceFunction ;
            "DeviceType" = $DeviceType ;
            "SubLocationID" = $Device.split(',')[0].split('-')[2] ;
            "Enumerator" = $Device.split(',')[0].split('-')[3]
        }
        $AP = New-Object pscustomobject -Property $Properties
        $APs += $AP
    }
    Return $APs
}

function Get-BuildingInfo
{
    param($Path)
    $Import = Import-Excel -Path $Path | select -Property Name,ABBR
    $Buildings = @()
    foreach ($Entry in $Import)
    {
        $Name = $Entry.name.replace('\','_')
        $Name = $Name.Replace('/','_')
        $Name = $Name.Replace(' ','_')
        $Obj = New-Object PSCUstomObject
        $Obj | Add-Member -MemberType NoteProperty -Name 'Name' -Value $Name
        $Obj | Add-Member -MemberType NoteProperty -Name 'ABBR' -Value $Entry.ABBR
        $Buildings += $Obj
    }
    Return $Buildings
}

function Get-iPerfEntryConnectionType
{
    param ($Entry)
    if ($Entry.BSSID -eq ''){$iPerfEntryConnectionType = 'Ethernet'}
    elseif ($Entry.'Adapter Name' -match 'Apple'){$connectionType = 'Ethernet'}
    elseif (($Entry.'Client Reverse Receive Rate (Mb/s)' -gt $Entry.'Link Speed (Mb/s)') -or ($Entry.'Client Send Rate (Mb/s)' -gt $Entry.'Link Speed (Mb/s)')){$iPerfEntryConnectionType = 'Ethernet'}
    elseif ($Entry.'Adapter Name' -match 'cisco'){$iPerfEntryConnectionType = 'VPN'}
    elseif ($Entry.'Adapter Name' -match 'Virtual'){$iPerfEntryConnectionType = 'Virtual'}
    elseif ($Entry.'Adapter Name' -match 'Hyper-V'){$iPerfEntryConnectionType = 'Virtual'}
    elseif ($Entry.'Adapter Name' -match 'Adapter Name'){$iPerfEntryConnectionType = 'Unknown'}
    else {$iPerfEntryConnectionType = 'WiFi'}
    return $iPerfEntryConnectionType
}

function Get-iPerfConnectionVendor
{
    param ($Entry)
    if (($entry.'Source IP'.split('.')[0] -eq '10') -and (($entry.'Source IP'.split('.')[1] -ge '104') -and ($entry.'Source IP'.split('.')[1] -le '107'))){$iPerfConnectionVendor = 'Aruba'}else{$iPerfConnectionVendor = 'Non-Aruba'}
    return $iPerfConnectionVendor
}

function Get-iPerfEntryStatus
{
    param ($Entry)
    $GoodLinkPercent = '50'
    $DownloadPercentOfLink = [math]::Round((($Entry.'Client Reverse Receive Rate (Mb/s)' / $Entry.'Link Speed (Mb/s)')*100),2)
    if ($DownloadPercentOfLink -ge $GoodLinkPercent){$DownState = 'Good'} else {$DownState = 'Bad'}
    $UploadPercentOfLink = [math]::Round((($Entry.'Client Send Rate (Mb/s)' / $Entry.'Link Speed (Mb/s)')*100),2)
    if ($UploadPercentOfLink -ge $GoodLinkPercent){$UpState = 'Good'} else {$UpState = 'Bad'}
    $iPerfEntryStatus = New-Object pscustomobject
    $iPerfEntryStatus | Add-Member -MemberType NoteProperty -Name 'DownloadStatus' -Value $DownState
    $iPerfEntryStatus | Add-Member -MemberType NoteProperty -Name 'DownloadPercentOfLink' -Value $DownloadPercentOfLink
    $iPerfEntryStatus | Add-Member -MemberType NoteProperty -Name 'UploadStatus' -Value $UpState
    $iPerfEntryStatus | Add-Member -MemberType NoteProperty -Name 'UploadPercentOfLink' -Value $UploadPercentOfLink

    # determine if in ideal state
    if 
    (
        ($Entry.Band -eq '5GHz') -and 
        ($Entry.'Wireless Adapter AC Power Setting' -eq 'Maximum Performance') -and 
        ($Entry.'AC Power Status' -eq 'Online') -and 
        ($Entry.SSID -eq "JAYHAWK"))
    {
        $iPerfEntryStatus | Add-Member -MemberType NoteProperty -Name 'Ideal' -Value $true
    } 
    else 
    {
        $iPerfEntryStatus | Add-Member -MemberType NoteProperty -Name 'Ideal' -Value $false
    }
    return $iPerfEntryStatus
}

function Export-iPerfStats
{
    param($Stats,$Path,$ExportPath,$WorksheetName)
    if (!(Test-path -Path ($Path + '\_Stats')))
    {
        $CreateFolder = New-Item -Path $Path -Name '_Stats' -ItemType Directory -Force
    }

    $OverallStats = New-Object PSCustomObject
    $OverallStats | Add-Member -MemberType NoteProperty -Name 'Scope' -Value 'Overall'
    $OverallStats | Add-Member -MemberType NoteProperty -Name 'Entry Count' -Value $Stats.count
    $OverallStats | Add-Member -MemberType NoteProperty -Name 'Unique Computer Count' -Value ($Stats.'Computer Name' | Select -Unique).count
    $OverallStats | Add-Member -MemberType NoteProperty -Name '% Unique Computers With AC Card' -Value ((($Stats.where({$_.entry.'adapter name' -match 'ac' -and $_.entry.'adapter name' -notmatch 'surface'}).'computer name' | select -Unique).count / ($stats.'computer name' | select -Unique).count)*100)
    $OverallStats | Add-Member -MemberType NoteProperty -Name '% Unique Computers in Ideal State' -Value ((($Stats.where({$_.entry.'ideal' -eq 'true'}).'computer name' | select -Unique).count / ($Stats.'computer name' | select -Unique).count)*100)
    $OverallStats | Add-Member -MemberType NoteProperty -Name 'Mean Upload (Mb/s)' -Value ($Stats.entry.'Client Send Rate (Mb/s)' | Measure-Object -Average).Average
    $OverallStats | Add-Member -MemberType NoteProperty -Name 'Mean Download (Mb/s)' -Value ($Stats.entry.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Average).Average
    $OverallStats | Add-Member -MemberType NoteProperty -Name 'Percent Of Time Upload Entries Are Good' -Value (($Stats.Entry.where({$_.'Upload Status' -eq 'Good'}).count / $Stats.Entry.count)*100)
    $OverallStats | Add-Member -MemberType NoteProperty -Name 'Percent Of Time Download Entries Are Good' -Value (($Stats.Entry.where({$_.'Download Status' -eq 'Good'}).count / $Stats.Entry.count)*100)
    Export-Excel -Path $ExportPath -TargetData $OverallStats -WorksheetName $WorksheetName -AutoSize -AutoFilter

    foreach ($AP in $Stats | Group 'AP Name')
    {
        $Stats = New-Object PSCustomObject
        $Stats | Add-Member -MemberType NoteProperty -Name 'Scope' -Value $AP.Name
        $Stats | Add-Member -MemberType NoteProperty -Name 'Entry Count' -Value $AP.Group.Count
        $Stats | Add-Member -MemberType NoteProperty -Name 'Unique Computer Count' -Value ($AP.Group.'Computer Name' | Select -Unique).count
        $Stats | Add-Member -MemberType NoteProperty -Name '% Unique Computers With AC Card' -Value ((($AP.group.where({$_.entry.'adapter name' -match 'ac' -and $_.entry.'adapter name' -notmatch 'surface'}).'computer name' | select -Unique).count / ($ap.group.'computer name' | select -Unique).count)*100)
        $Stats | Add-Member -MemberType NoteProperty -Name '% Unique Computers in Ideal State' -Value ((($AP.group.where({$_.entry.'ideal' -eq 'true'}).'computer name' | select -Unique).count / ($AP.group.'computer name' | select -Unique).count)*100)
        $Stats | Add-Member -MemberType NoteProperty -Name 'Mean Upload (Mb/s)' -Value ($AP.Group.Entry.'Client Send Rate (Mb/s)' | Measure-Object -Average).Average
        $Stats | Add-Member -MemberType NoteProperty -Name 'Mean Download (Mb/s)' -Value ($AP.Group.Entry.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Average).Average
        $Stats | Add-Member -MemberType NoteProperty -Name 'Percent Of Time Upload Entries Are Good' -Value (($AP.group.where({$_.'Upload Status' -eq 'Good'}).count / $AP.group.count)*100)
        $Stats | Add-Member -MemberType NoteProperty -Name 'Percent Of Time Download Entries Are Good' -Value (($AP.group.where({$_.'Download Status' -eq 'Good'}).count / $AP.group.count)*100)
        Export-Excel -Path $ExportPath -TargetData $Stats -WorksheetName $WorksheetName -AutoSize -AutoFilter -Append
    }
}

#####
##### Primary Functions
#####

function Get-iPerfEntries
{
    param($Path)
    $Folders = Get-ChildItem -Path $Path -Directory | Where-Object {($_.Name -ne "Errors") -and ($_.Name -ne "_Archive") -and ($_.name -ne '_Stats')}
    $Entries = @()
    [int]$i = '0'
    Foreach ($Folder in $Folders)
    {
        Write-Progress -Id '1' -Activity 'Getting iPerf Departments' -Status 'Getting Department' -CurrentOperation $Folder.Name -PercentComplete (($i / $Folders.count)*100)
        $Files = Get-ChildItem -Path $Folder.Fullname -File
        [int]$i2 = '0'
        Foreach ($File in $Files)
        {
            Write-Progress -Id '2' -Activity 'Getting iPerf Reports' -Status 'Getting Report' -CurrentOperation $File.Name -PercentComplete (($i2 / $Files.count)*100)
            $CSV = Import-Csv -Path $File.fullname
            $ComputerName = $File.fullname.split('\')[-1].Replace('_iPerf.csv','')
            $Department = $File.fullname.split('\')[-1].Replace('_iPerf.csv','').Split('-')[0]
            [int]$i3 = '0'
            foreach ($Line in $CSV)
            {
                Write-Progress -Id '3' -Activity 'Getting iPerf Report Entries' -Status 'Getting Entry' -CurrentOperation $i3 -PercentComplete (($i3 / $CSV.count)*100)
                $ConnectionType = Get-iPerfEntryConnectionType -Entry $Line
                $ConnectionVendor = Get-iPerfConnectionVendor -Entry $Line
                $EntryStatus = Get-iPerfEntryStatus -Entry $Line
                $APName = $APDetails.where({($_.'MAC-BGN' -eq $Line.BSSID) -or ($_.'MAC-AC' -eq $Line.BSSID)}).Name
                $APIPAddress = $APDetails.where({($_.'MAC-BGN' -eq $Line.BSSID) -or ($_.'MAC-AC' -eq $Line.BSSID)}).IPAddress
                $LocationID = $APDetails.where({($_.'MAC-BGN' -eq $Line.BSSID) -or ($_.'MAC-AC' -eq $Line.BSSID)}).LocationID
                $BuildingName = $APDetails.where({($_.'MAC-BGN' -eq $Line.BSSID) -or ($_.'MAC-AC' -eq $Line.BSSID)}).BuildingName
                $DeviceFunction = $APDetails.where({($_.'MAC-BGN' -eq $Line.BSSID) -or ($_.'MAC-AC' -eq $Line.BSSID)}).DeviceFunction
                $DeviceType = $APDetails.where({($_.'MAC-BGN' -eq $Line.BSSID) -or ($_.'MAC-AC' -eq $Line.BSSID)}).DeviceType
                $SubLocationID = $APDetails.where({($_.'MAC-BGN' -eq $Line.BSSID) -or ($_.'MAC-AC' -eq $Line.BSSID)}).SubLocationID
                $Enumerator = $APDetails.where({($_.'MAC-BGN' -eq $Line.BSSID) -or ($_.'MAC-AC' -eq $Line.BSSID)}).Enumerator

                [datetime]$Timestamp = $line.'time stamp'
                $entryHashtable = @{
                    "Computer Name" = $ComputerName;
                    "Department" = $Department;
                    "Timestamp" = $Timestamp;
                    "Year" = $Timestamp.Year;
                    "Month" = $Timestamp.Month;
                    "Client Send Rate (Mb/s)" = ([math]::Round(($line.'Client Send Rate (Mb/s)'),2));
                    "Client Reverse Receive Rate (Mb/s)" = ([math]::round(($line.'Client Reverse Receive Rate (Mb/s)'),2));
                    "Link Speed (Mb/s)" = $line.'Link Speed (Mb/s)';
                    "SSID" = $line.SSID;
                    "Radio Type" = $line.'Radio Type';
                    "Channel" = $line.Channel;
                    "Band" = $line.Band;
                    "Signal Strength" = $line.'Signal Strength';
                    "BSSID" = $line.BSSID;
                    "MAC Address" = $line.'MAC Address';
                    "Source IP" = $line.'Source IP';
                    "Adapter Name" = $line.'Adapter Name';
                    "Driver Version" = $line.'Driver Version';
                    "Driver Provider" = $line.'Driver Provider';
                    "Driver Date" = $line.'Driver Date';
                    "OS Name" = $line.'OS Name';
                    "OS Version" = $line.'OS Version';
                    "Power Plan" = $line.'Power Plan';
                    "AC Power Status" = $line.'AC Power Status';
                    "Wireless Adapter AC Power Setting" = $line.'Wireless Adapter AC Power Setting';
                    "Wireless Adapter DC Power Setting" = $line.'Wireless Adapter DC Power Setting';
                    "Username" = $line.Username;
                    "iPerf Server" = $line.'iPerf Server';
                    "Model" = $line.Model;
                    "Chassis Type" = $line.'Chassis Type';
                    "AP Name" = $APName ;
                    "AP IP Address" = $APIPAddress ;
                    "LocationID" = $LocationID ;
                    "Building Name" = $BuildingName ;
                    "DeviceFunction" = $DeviceFunction ;
                    "DeviceType" = $DeviceType ;
                    "SubLocationID" = $SubLocationID ;
                    "Enumerator" = $Enumerator ;
                    "Connection Type" = $ConnectionType ;
                    "Connection Vendor" = $ConnectionVendor ;
                    "Download Status" = $EntryStatus.DownloadStatus ;
                    "Upload Status" = $EntryStatus.UploadStatus ;
                    "Download Percent of Link" = $EntryStatus.DownloadPercentOfLink ;
                    "Upload Percent of Link" = $EntryStatus.UploadPercentOfLink ;
                    "Ideal" = $EntryStatus.Ideal
                }
                $iPerfEntry = New-Object PSCustomObject
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'Computer Name' -Value $ComputerName
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'Department Name' -Value $Department
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'AP Name' -Value $APName
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'Building Name' -Value $BuildingName
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'SubLocationID' -Value $SubLocationID
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'Connection Type' -Value $ConnectionType
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'Connection Vendor' -Value $ConnectionVendor
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'Download Status' -Value $EntryStatus.DownloadStatus
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'Upload Status' -Value $EntryStatus.UploadStatus
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'Download Percent of Link' -Value $EntryStatus.DownloadPercentOfLink
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'Upload Percent of Link' -Value $EntryStatus.UploadPercentOfLink
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'Ideal' -Value $EntryStatus.Ideal
                $iPerfEntry | Add-Member -MemberType NoteProperty -Name 'Entry' -Value $EntryHashtable
                $Entries += $iPerfEntry
                $i3++
            }
            $i2++
        }
        $i++
    }
    Return $Entries
}

#####
##### Main Process
#####

# Gather building details
Write-Progress -id '1' -Activity 'Getting Building Details'
$BuildingDetails = Get-BuildingInfo -Path $BuildingInfoPath

# Gather Airwave AP details
Write-Progress -id '1' -Activity 'Getting AirWave AP Details'
$APDetails = Get-AirWaveAPs -ScriptPath $AirwaveScript -BuildingDetails $BuildingDetails

# Gather iPerf entries
$iPerfEntries = Get-iPerfEntries -Path $Path

# Select just the entires that are on wifi, on aruba, and has a building name
$EntriesInScope = $iPerfEntries | where ({
    $_.entry.'connection type' -eq 'wifi' -and 
    $_.entry.'connection vendor' -eq 'aruba' -and
    $_.'building name' -ne $null
})

# Determine all years/months included in the reports
$TimeLine = @()
foreach ($year in $EntriesInScope.entry.year | select -Unique)
{
    $YearEntry = @()
    foreach ($month in $EntriesInScope.entry.where({$_.year -eq $year}).month | select -Unique)
    {
        $TimeEntry = New-Object PSCustomObject
        $TimeEntry | Add-Member -MemberType NoteProperty -Name 'Year' -Value $year
        $TimeEntry | Add-Member -MemberType NoteProperty -Name 'Month' -Value $month
        $YearEntry += $TimeEntry
    }
    $YearEntry = $YearEntry | Sort -Property 'Month'
    $Timeline += $YearEntry
}
$Timeline = $TimeLine | Sort -Property 'Year'

foreach ($building in ($EntriesInScope | group 'building name'))
{
    $ExportPath = $ExportRoot + '\' + $building.name + '_iPerfReport_' + $DateTime + '.xlsx'
    Export-iPerfStats -Stats $building.group -Path $path -ExportPath $ExportPath -WorksheetName 'All'

    foreach ($Entry in $TimeLine)
    {
        $WorksheetName = $Entry.Year.ToString() + " " + $Entry.Month.ToString()
        $TimeLimitedEntries = $Building.group.where({$_.entry.year -eq $Entry.Year -and $_.entry.month -eq $Entry.Month})
        Export-iPerfStats -Stats $TimeLimitedEntries -Path $Path -ExportPath $ExportPath -WorksheetName $WorksheetName
    }
}