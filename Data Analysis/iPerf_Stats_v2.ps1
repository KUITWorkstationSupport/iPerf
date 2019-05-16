#####################
### edit this section
#####################

$VerbosePreference = 'Continue'
#$VerbosePreference = 'SilentlyContinue'

#$path = "C:\users\Andy Jackson\OneDrive - The University of Kansas\Desktop\iPerf"
#$path = "C:\Users\ajacks88\OneDrive - The University of Kansas\Desktop\iperf"
$path = "\\cfs.home.ku.edu\it_general\units\ss\itt2\iperf"

# minumum number of days a report must have before it will be considered
[int]$minReportDaySpan = '5'

# number to divide the link speed by to determine what is "good"
[int]$linkDivisor = '2'

# path to folder to export stats
$exportFolderPath = 'C:\users\ajacks88\OneDrive - The University of Kansas\Desktop'
#$exportFolderPath = 'C:\users\Andy Jackson\OneDrive - The University of Kansas\Desktop'

# name of export folder
$exportFolderName = 'iPerf_Stats'

####################
### end edit section
####################

####################
### helper functions
####################

function find-mean 
{
    param
    (
        $list
    )
    [int]$total = '0'
    foreach ($entry in $list)
    {
        $total += $entry
    }
    return ($total / $list.count)
}

function new-obj 
{
    param
    (
        $inObj,
        $computerName
    )

    $obj = New-Object system.object
    $obj | Add-Member -MemberType NoteProperty -Name "Computer Name" -Value $computerName
    $obj | Add-Member -MemberType NoteProperty -Name "Time Stamp" -Value $inObj.'Time Stamp'
    $obj | Add-Member -MemberType NoteProperty -Name "Client Send Rate (Mb/s)" -Value $inObj.'Client Send Rate (Mb/s)'
    $obj | Add-Member -MemberType NoteProperty -Name "Client Reverse Receive Rate (Mb/s)" -Value $inObj.'Client Reverse Receive Rate (Mb/s)'
    $obj | Add-Member -MemberType NoteProperty -Name "Link Speed (Mb/s)" -Value $inObj.'Link Speed (Mb/s)'
    $obj | Add-Member -MemberType NoteProperty -Name "SSID" -Value $inObj.SSID
    $obj | Add-Member -MemberType NoteProperty -Name "Radio Type" -Value $inObj.'Radio Type'
    $obj | Add-Member -MemberType NoteProperty -Name "Channel" -Value $inObj.Channel
    $obj | Add-Member -MemberType NoteProperty -Name "Band" -Value $inObj.Band
    $obj | Add-Member -MemberType NoteProperty -Name "Signal Strength" -Value $inObj.'Signal Strength'
    $obj | Add-Member -MemberType NoteProperty -Name "BSSID" -Value $inObj.BSSID
    $obj | Add-Member -MemberType NoteProperty -Name "MAC Address" -Value $inObj.'MAC Address'
    $obj | Add-Member -MemberType NoteProperty -Name "Source IP" -Value $inObj.'Source IP'
    $obj | Add-Member -MemberType NoteProperty -Name "Adapter Name" -Value $inObj.'Adapter Name'
    $obj | Add-Member -MemberType NoteProperty -Name "Driver Version" -Value $inObj.'Driver Version'
    $obj | Add-Member -MemberType NoteProperty -Name "Driver Provider" -Value $inObj.'Driver Provider'
    $obj | Add-Member -MemberType NoteProperty -Name "Driver Date" -Value $inObj.'Driver Date'
    $obj | Add-Member -MemberType NoteProperty -Name "OS Name" -Value $inObj.'OS Name'
    $obj | Add-Member -MemberType NoteProperty -Name "OS Version" -Value $inObj.'OS Version'
    $obj | Add-Member -MemberType NoteProperty -Name "Power Plan" -Value $inObj.'Power Plan'
    $obj | Add-Member -MemberType NoteProperty -Name "AC Power Status" -Value $inObj.'AC Power Status'
    $obj | Add-Member -MemberType NoteProperty -Name "Wireless Adapter AC Power Setting" -Value $inObj.'Wireless Adapter AC Power Setting'
    $obj | Add-Member -MemberType NoteProperty -Name "Wireless Adapter DC Power Setting" -Value $inObj.'Wireless Adapter DC Power Setting'
    $obj | Add-Member -MemberType NoteProperty -Name "Username" -Value $inObj.Username
    $obj | Add-Member -MemberType NoteProperty -Name "iPerf Server" -Value $inObj.'iPerf Server'
    $obj | Add-Member -MemberType NoteProperty -Name "Model" -Value $inObj.Model
    $obj | Add-Member -MemberType NoteProperty -Name "Chassis Type" -Value $inObj.'Chassis Type'

    return $obj
}

####################
### create folders
####################

# get filename friendly timestamp
$timestamp = Get-date -Format yyyy-MM-dd_hh.mm.ss_tt

# set export folder path
$exportFolderName = $exportFolderName + "_" + $timestamp
$exportFullPath = $exportFolderPath + "\" + $exportFolderName

# create export folders
$createFolder = New-Item -Path $exportFolderPath -Name $exportFolderName -ItemType directory
$createFolder = New-Item -Path ($exportFolderPath + '\' + $exportFolderName) -Name "ap bad lists" -ItemType directory
$createFolder = New-Item -Path ($exportFolderPath + '\' + $exportFolderName) -Name "radio mismatch" -ItemType directory
$createFolder = New-Item -Path ($exportFolderPath + '\' + $exportFolderName) -Name "computer bad lists" -ItemType directory
$createFolder = New-Item -Path ($exportFolderPath + '\' + $exportFolderName) -Name "_good reports" -ItemType directory
$createFolder = New-Item -Path ($exportFolderPath + '\' + $exportFolderName) -Name "_bad reports" -ItemType directory

####################
### create variables
####################

# count # of only ethernet reports
[int]$onlyEthCount = '0'

# count # 100% good reports
[int]$onlyGoodCount = '0'

# keep track of what report number we are on
[int]$reportNumber = '0'

# find the told number of reports imported
[int]$reportTotal = '0'

# array of good report numbers
$goodReportList = @()

# array of all download speeds
$downList = @()

# array of all upload speeds
$upList = @()

# array of what % of the reports download speeds are "good"
$percentGoodDownList = @()

# array of what % of the reports upload speeds are "good"
$percentGoodUpList = @()

# array of bad reports
$badList = @()

# array of good reports
$goodList = @()

# array of radio mismatch entries
$radioMismatchList = @()

# array of inconsistant download variations by more than -10%
$inconsistantDown10List = @()

# array of inconsistant download variations by more than -20%
$inconsistantDown20List = @()

# array of inconsistant download variations by more than -50%
$inconsistantDown50List = @()

# array of inconsistant upload variations by more than -10%
$inconsistantUp10List = @()

# array of inconsistant upload variations by more than -20%
$inconsistantUp20List = @()

# array of inconsistant upload variations by more than -50%
$inconsistantUp50List = @()

####################
### loop through each report
####################

####################
### gather data
####################

$folders = Get-ChildItem -Path $path -Directory | Where-Object {$_.Name -ne "Errors"} | Where-Object {$_.Name -ne "_Archive"}

foreach ($folder in $Folders)
{
    # gather reports
    $reports = Get-ChildItem -Path $folder.fullname -File
    Write-Verbose ("# of reports = " + $reports.count)
    $reportTotal += $reports.count

    foreach ($report in $reports)
    {
        # create csv object of current report
        $csv = Import-Csv -Path $report.FullName
        Write-Verbose ("starting report: " + $reportNumber + ", # of entries = " + $csv.count)

        # create counter variables
        [int]$ethCount = '0'
        [int]$entryCount = '0'
        [int]$goodDownCount = '0'
        [int]$goodUpCount = '0'
        [int]$entryTotal = $csv.count
        [int]$linkSpeedChangeCount = '0'

        #[int]$inconsistantDown10Count = '0'
        #[int]$inconsistantDown20Count = '0'
        #[int]$inconsistantDown50Count = '0'
        #[int]$inconsistantUp10Count = '0'
        #[int]$inconsistantUp20Count = '0'
        #[int]$inconsistantUp50Count = '0'

        $inconsistantDown10 = @()
        $inconsistantDown20 = @()
        $inconsistantDown50 = @()
        $inconsistantUp10 = @()
        $inconsistantUp20 = @()
        $inconsistantUp50 = @()

        # check to see if report meets minimum day span requirement
        if ((New-TimeSpan -Start $csv[0].'Time Stamp' -End $csv[($csv.count -1)].'Time Stamp').Days -lt $minReportDaySpan)
        {
            # report doesnt meet mimimum day span requirement, skip to next report
            Write-Verbose ("Report does not span at least " + $minReportDaySpan + " days, skipping")

            # decrease report total by 1 for accuracy in statistics
            $reportTotal--
        }
        else
        {
            # loop through each entry in report
            foreach ($entry in $csv)
            {
                Write-Verbose ("report: " + $reportNumber + ", entry: " + $entryCount)

                # check to see if entry is ethernet or wifi traffic
                if (($entry.BSSID -eq ''))
                {
                    # entry is ethernet, ignore
                    Write-Verbose ("No BSSID detected, this entry is ethernet, adding 1 to ethCount")
                    $ethCount++
                    # decrease entry total by 1 for accuracy in statistics
                    $entryTotal--
                }
                # check to see if entry is VPN traffic
                elseif ($entry.'Adapter Name' -match 'AnyConnect')
                {
                    # entry is VPN traffic, ignore
                    Write-Verbose ("Adapter name matches AnyConnect, this entry is VPN, ignoring")
                    # decrease entry total by 1 for accuracy in statistics
                    $entryTotal--
                }
                # check to see if link speed is way too high indicating an error
                elseif ($entry.'Link Speed (Mb/s)' -gt '800')
                {
                    # link speed is way too high, something went wrong, ignore entry
                    Write-Verbose ("Link speed greater than 800 on wifi, something went wrong, ignoring this entry")
                    # decrease entry total by 1 for accuracy in statistics
                    $entryTotal--
                }
                elseif ($entry.SSID -ne 'JAYHAWK')
                {
                    Write-Verbose ("SSID is not JAYHAWK, ignoring this entry")
                    $entryTotal--
                }
                # entry is good, continue to analyze
                else
                {
                    # check to see if radio mismatch
                    if (($entry.'Adapter Name' -match 'ac') -and ($entry.'Radio Type' -notmatch 'ac'))
                    {
                        $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                        $radioMismatchList += $obj
                    }

                    $downList += $entry.'Client Reverse Receive Rate (Mb/s)'
                    # determine if download speed is good or bad
                    if ($entry.'Client Reverse Receive Rate (Mb/s)' -gt ($entry.'Link Speed (Mb/s)' / $linkDivisor))
                    {
                        # download speed is good
                        Write-Host "down:" $entry.'Client Reverse Receive Rate (Mb/s)' "link:" $entry.'Link Speed (Mb/s)' -ForegroundColor Green
                        $goodDownCount++
                        $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                        $goodList += $obj
                    }
                    else
                    {
                        # download speed is bad
                        Write-Host "down:" $entry.'Client Reverse Receive Rate (Mb/s)' "link:" $entry.'Link Speed (Mb/s)' -ForegroundColor Red
                        $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                        $badList += $obj
                    }

                    $upList += $entry.'Client Send Rate (Mb/s)'
                    # check to see if upload speed is good or bad
                    if ($entry.'Client Send Rate (Mb/s)' -gt ($entry.'Link Speed (Mb/s)' / $linkDivisor))
                    {
                        # upload speed is good
                        Write-Host "up:" $entry.'Client Send Rate (Mb/s)' "link:" $entry.'Link Speed (Mb/s)' -ForegroundColor Green
                        $goodUpCount++
                        $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                        $goodList += $obj
                    }
                    else
                    {
                        # upload speed is bad
                        Write-Host "up:" $entry.'Client Send Rate (Mb/s)' "link:" $entry.'Link Speed (Mb/s)' -ForegroundColor Red
                        $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                        $badList += $obj
                    }

                    if ($entryCount -eq '0')
                    {
                
                    }
                    elseif ($entry.BSSID -eq $csv[($entryCount -1)].bssid)
                    {
                        $downDifferencePercent = (($entry.'Client Reverse Receive Rate (Mb/s)' - $csv[($entryCount -1)].'Client Reverse Receive Rate (Mb/s)') / $csv[($entryCount -1)].'Client Reverse Receive Rate (Mb/s)')*100
                        $upDifferencePercent = (($entry.'Client Send Rate (Mb/s)' - $csv[($entryCount -1)].'Client Send Rate (Mb/s)') / $csv[($entryCount -1)].'Client Send Rate (Mb/s)')*100
                        Write-Host "down difference = " $downDifferencePercent -ForegroundColor Yellow
                        Write-Host "up difference = " $upDifferencePercent -ForegroundColor Yellow

                        if ($downDifferencePercent -lt '-10')
                        {
                            #$inconsistantDown10Count++
                            $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                            $inconsistantDown10 += $obj
                        }
                        if ($downDifferencePercent -lt '-20')
                        {
                            #$inconsistantDown20Count++
                            $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                            $inconsistantDown20 += $obj
                        }
                        if ($downDifferencePercent -lt '-50')
                        {
                            #$inconsistantDown50Count++
                            $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                            $inconsistantDown50 += $obj
                        }
                        if ($upDifferencePercent -lt '-10')
                        {
                            #$inconsistantUp10Count++
                            $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                            $inconsistantUp10 += $obj
                        }
                        if ($upDifferencePercent -lt '-20')
                        {
                            #$inconsistantUp20Count++
                            $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                            $inconsistantUp20 += $obj
                        }
                        if ($upDifferencePercent -lt '-50')
                        {
                            #$inconsistantUp50Count++
                            $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                            $inconsistantUp50 += $obj
                        }
                    }
                }
                $entryCount++
            }

            # check to see if report is 100% ethernet traffic
            if ($ethCount -eq $csv.count)
            {
                # report is 100% ethernet, ignore
                Write-Verbose ("ethCount = entryCount, this whole report was ethernet, adding 1 to onlyEthCount")
                $onlyEthCount++
            }
            # check to see if report is 100% good
            elseif (($goodDownCount -eq $entryTotal) -and ($goodUpCount -eq $entryTotal))
            {
                # report is 100% good
                $onlyGoodCount++
                $goodReportList += $reportNumber
                foreach ($entry in $csv)
                {
                    $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                    export-csv -InputObject $obj -Path ($exportFullPath + "\_good reports\" + $report.Name) -Append -NoTypeInformation
                }
            }
            else
            {
                # report is less than 100%
                Write-Verbose ("goodDownCount = " + $goodUpCount + ", goodUpCount = " + $goodUpCount + ", total entries = " + $entryTotal)
                foreach ($entry in $csv)
                {
                    $obj = new-obj -inObj $entry -computerName $report.Name.replace('_iPerf.csv','')
                    export-csv -InputObject $obj -Path ($exportFullPath + "\_bad reports\" + $report.Name) -Append -NoTypeInformation
                }
            }
            # check to see if entry total is 0, meaning this report is 100% ethernet
            if ($entryTotal -eq '0')
            {
                # report is 100% ethernet
            }
            else
            {
                # report is not 100% ethernet
                $percentGoodDownList += (($goodDownCount / $entryTotal) * 100)
                $percentGoodUpList += (($goodUpCount / $entryTotal) * 100)
            }

            if ($inconsistantDown10.count -gt 1)
            {
                $inconsistantDown10List += $inconsistantDown10
            }
            if ($inconsistantDown20.count -gt 1)
            {
                $inconsistantDown20List += $inconsistantDown20
            }
            if ($inconsistantDown50.count -gt 1)
            {
                $inconsistantDown50List += $inconsistantDown50
            }
            if ($inconsistantUp10.count -gt 1)
            {
                $inconsistantUp10List += $inconsistantUp10
            }
            if ($inconsistantUp20.count -gt 1)
            {
                $inconsistantUp20List += $inconsistantUp20
            }
            if ($inconsistantUp50.count -gt 1)
            {
                $inconsistantUp50List += $inconsistantUp50
            }
        }
        $reportNumber++
    }
}

####################
### export data
####################

# group goodList by BSSID
$goodListGroupedBSSID = $goodList | Group-Object -Property BSSID
# group badlist by BSSID and export
$badListGroupedBSSID = $badList | Group-Object -Property BSSID
foreach ($AP in $badListGroupedBSSID)
{
    $obj = New-Object System.Object
    $obj | Add-Member -MemberType NoteProperty -Name "AP MAC Address" -Value $AP.name
    $obj | Add-Member -MemberType NoteProperty -Name "# Computers" -Value ($AP.Group | Group-Object -Property "Computer Name").name.count
    $obj | Add-Member -MemberType NoteProperty -Name "Total Bad Entries" -Value $AP.group.count
    $obj | Add-Member -MemberType NoteProperty -Name "Total Good Entries" -Value ($goodListGroupedBSSID | where {$_.name -eq $AP.name}).count
    $obj | Add-Member -MemberType NoteProperty -Name "% bad of total entries" -Value (($AP.group.count / ($AP.group.count + ($goodListGroupedBSSID | where {$_.name -eq $AP.name}).count)) * 100)
    Export-csv -InputObject $obj -Path ($exportFullPath + "\ap_bad_list.csv") -Append -NoTypeInformation

    foreach ($entry in $AP.Group)
    {
        $obj = new-obj -inObj $entry -computerName $entry."computer name"
        export-csv -InputObject $obj -Path ($exportFullPath + "\ap bad lists\" + ($AP.name).Replace(":","-") + ".csv") -Append -NoTypeInformation
    }
}

# group goodList by computer name
$goodListGroupedComputerName = $goodList | Group-Object -Property 'computer name'
# group badlist by computer name and export
$badListGroupedComputerName = $badList | Group-Object -Property "computer name"
foreach ($Computer in $badListGroupedComputerName)
{
    $obj = New-Object System.Object
    $obj | Add-Member -MemberType NoteProperty -Name "computer name" -Value $Computer.name
    $obj | Add-Member -MemberType NoteProperty -Name "# APs" -Value ($Computer.Group | Group-Object -Property "BSSID").name.count
    $obj | Add-Member -MemberType NoteProperty -Name "Total Bad Entries" -Value $Computer.group.count
    $obj | Add-Member -MemberType NoteProperty -Name "Total Good Entries" -Value ($goodListGroupedComputerName | where {$_.name -eq $computer.name}).count
    $obj | Add-Member -MemberType NoteProperty -Name "% bad of total entries" -Value (($computer.group.count / ($computer.group.count + ($goodListGroupedComputerName | where {$_.name -eq $computer.name}).count)) * 100)
    Export-csv -InputObject $obj -Path ($exportFullPath + "\computer_bad_list.csv") -Append -NoTypeInformation

    foreach ($entry in $Computer.Group)
    {
        $obj = new-obj -inObj $entry -computerName $entry."computer name"
        export-csv -InputObject $obj -Path ($exportFullPath + "\computer bad lists\" + ($Computer.name).Replace(":","-") + ".csv") -Append -NoTypeInformation
    }
}

# group mismatch by computer name and export
$radioMismatchListGrouped = $radioMismatchList | Group-Object -Property "computer name"
foreach ($group in $radioMismatchListGrouped)
{
    foreach ($entry in $group.group)
    {
        $obj = new-obj -inObj $entry -computerName $entry."computer name"
        export-csv -InputObject $obj -Path ($exportFullPath + "\radio mismatch\" + ($group.name).Replace(":","-") + ".csv") -Append -NoTypeInformation
    }
}

# export badlist
foreach ($entry in $badList)
{
    $obj = new-obj -inObj $entry -computerName $entry."computer name"
    export-csv -InputObject $obj -Path ($exportFullPath + "\badlist.csv") -Append -NoTypeInformation
}

# do some math
$meanDown = find-mean -list $downList
$meanUp = find-mean -list $upList
$meanPercentGoodDown = find-mean -list $percentGoodDownList
$meanPercentGoodUp = find-mean -list $percentGoodUpList

$percentInconsistant1Down10 = (($inconsistantDown10list | group 'computer name').count / $reportTotal) * 100
$percentInconsistant1Down20 = (($inconsistantDown20list | group 'computer name').count / $reportTotal) * 100
$percentInconsistant1Down50 = (($inconsistantDown50list | group 'computer name').count / $reportTotal) * 100
$percentInconsistant1Up10 = (($inconsistantUp10list | group 'computer name').count / $reportTotal) * 100
$percentInconsistant1Up20 = (($inconsistantUp20list | group 'computer name').count / $reportTotal) * 100
$percentInconsistant1Up50 = (($inconsistantUp50list | group 'computer name').count / $reportTotal) * 100

$percentInconsistant5Down10 = ((($inconsistantDown10list | group 'computer name').where({$_.count -ge 5})).count / $reportTotal) * 100
$percentInconsistant5Down20 = ((($inconsistantDown20list | group 'computer name').where({$_.count -ge 5})).count / $reportTotal) * 100
$percentInconsistant5Down50 = ((($inconsistantDown50list | group 'computer name').where({$_.count -ge 5})).count / $reportTotal) * 100
$percentInconsistant5Up10 = ((($inconsistantUp10list | group 'computer name').where({$_.count -ge 5})).count / $reportTotal) * 100
$percentInconsistant5Up20 = ((($inconsistantUp20list | group 'computer name').where({$_.count -ge 5})).count / $reportTotal) * 100
$percentInconsistant5Up50 = ((($inconsistantUp50list | group 'computer name').where({$_.count -ge 5})).count / $reportTotal) * 100

# export more data
foreach ($computer in (($inconsistantDown10list | group 'computer name').where({$_.count -ge 5})))
{
    $obj = New-Object System.Object
    $obj | Add-Member -MemberType NoteProperty -Name "computer name" -Value $Computer.name
    $obj | Add-Member -MemberType NoteProperty -Name "Total Down Inconsistancies" -Value $Computer.count
    Export-csv -InputObject $obj -Path ($exportFullPath + "\5_inconsistant_download_10_percent.csv") -Append -NoTypeInformation
}

foreach ($computer in (($inconsistantDown20list | group 'computer name').where({$_.count -ge 5})))
{
    $obj = New-Object System.Object
    $obj | Add-Member -MemberType NoteProperty -Name "computer name" -Value $Computer.name
    $obj | Add-Member -MemberType NoteProperty -Name "Total Down Inconsistancies" -Value $Computer.count
    Export-csv -InputObject $obj -Path ($exportFullPath + "\5_inconsistant_download_20_percent.csv") -Append -NoTypeInformation
}

foreach ($computer in (($inconsistantDown50list | group 'computer name').where({$_.count -ge 5})))
{
    $obj = New-Object System.Object
    $obj | Add-Member -MemberType NoteProperty -Name "computer name" -Value $Computer.name
    $obj | Add-Member -MemberType NoteProperty -Name "Total Down Inconsistancies" -Value $Computer.count
    Export-csv -InputObject $obj -Path ($exportFullPath + "\5_inconsistant_download_50_percent.csv") -Append -NoTypeInformation
}

foreach ($computer in (($inconsistantUp10list | group 'computer name').where({$_.count -ge 5})))
{
    $obj = New-Object System.Object
    $obj | Add-Member -MemberType NoteProperty -Name "computer name" -Value $Computer.name
    $obj | Add-Member -MemberType NoteProperty -Name "Total Down Inconsistancies" -Value $Computer.count
    Export-csv -InputObject $obj -Path ($exportFullPath + "\5_inconsistant_upload_10_percent.csv") -Append -NoTypeInformation
}

foreach ($computer in (($inconsistantUp20list | group 'computer name').where({$_.count -ge 5})))
{
    $obj = New-Object System.Object
    $obj | Add-Member -MemberType NoteProperty -Name "computer name" -Value $Computer.name
    $obj | Add-Member -MemberType NoteProperty -Name "Total Down Inconsistancies" -Value $Computer.count
    Export-csv -InputObject $obj -Path ($exportFullPath + "\5_inconsistant_upload_20_percent.csv") -Append -NoTypeInformation
}

foreach ($computer in (($inconsistantUp50list | group 'computer name').where({$_.count -ge 5})))
{
    $obj = New-Object System.Object
    $obj | Add-Member -MemberType NoteProperty -Name "computer name" -Value $Computer.name
    $obj | Add-Member -MemberType NoteProperty -Name "Total Down Inconsistancies" -Value $Computer.count
    Export-csv -InputObject $obj -Path ($exportFullPath + "\5_inconsistant_upload_50_percent.csv") -Append -NoTypeInformation
}

# write general stats
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("good is defined by if the speed is greater than the link speed divied by " + $linkDivisor)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("total reports considered : " + $reportTotal)
#Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("total 100% ethernet reports : " + $onlyEthCount)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("total reports with 0 bad entries : " + $onlyGoodCount)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("mean % good down : " + $meanPercentGoodDown + "%")
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("mean % good up : " + $meanPercentGoodUp + "%")
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("mean down speed : " + $meanDown + "(Mb/s)")
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("mean up speed : " + $meanUp + "(Mb/s)")
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("computers with atleast 1 radio mismatch = " + $radioMismatchListGrouped.count)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value " "
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 1 inconsistant download (-10%), same BSSID = " + $percentInconsistant1Down10)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 1 inconsistant download (-20%), same BSSID = " + $percentInconsistant1Down20)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 1 inconsistant download (-50%), same BSSID = " + $percentInconsistant1Down50)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 1 inconsistant upload (-10%), same BSSID = " + $percentInconsistant1Up10)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 1 inconsistant upload (-20%), same BSSID = " + $percentInconsistant1Up20)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 1 inconsistant upload (-50%), same BSSID = " + $percentInconsistant1Up50)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value " "
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 5 inconsistant download (-10%), same BSSID = " + $percentInconsistant5Down10)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 5 inconsistant download (-20%), same BSSID = " + $percentInconsistant5Down20)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 5 inconsistant download (-50%), same BSSID = " + $percentInconsistant5Down50)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 5 inconsistant upload (-10%), same BSSID = " + $percentInconsistant5Up10)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 5 inconsistant upload (-20%), same BSSID = " + $percentInconsistant5Up20)
Add-Content -Path ($exportFullPath + "\stats.txt") -Value ("% of time report containts more than 5 inconsistant upload (-50%), same BSSID = " + $percentInconsistant5Up50)