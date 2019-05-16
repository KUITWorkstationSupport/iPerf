###################
# edit this section 
###################

# verbose logging | 'continue' will allow, 'SilentlyContinue' will supress
$verbosePreference = 'Continue'
#$verbosePreference = 'SilentlyContinue'

# path to source of iperf reports
#$sourcePath = 'C:\users\Andy Jackson\OneDrive - The University of Kansas\Desktop\iperf'
#$sourcePath = 'C:\users\ajacks88\OneDrive - The University of Kansas\Desktop\iperf'
$sourcePath = '\\cfs.home.ku.edu\it_general\units\ss\itt2\iperf'

# path to the location in which results will be exported
#$exportPath = 'C:\users\Andy Jackson\OneDrive - The University of Kansas\Desktop'
$exportPath = 'C:\users\ajacks88\OneDrive - The University of Kansas\Desktop'

# name of the folder that will be created to store anything that is exported
$exportFolderName = 'iPerf_Stats'

# minimum days of reports needed before those reports will be analyzed
[int]$minReportDaySpan = '5'

# what percent of the link speed is the line between good/bad
[int]$goodLinkPercent = '50'

# what percent drop in speed is considered significant and should be documented
[int]$speedDropPercent = '-25'

# what ssid name should be targeted
$targetSSID = 'JAYHAWK'

###################
# don't edit below
###################

###################
# helper functions
###################

function Export-Stats
{
    param
    (
        $property,
        $measure,
        $nonIdeal,
        $ideal
    )
    $obj = New-Object System.Object
    $obj | Add-Member -MemberType NoteProperty -Name ' ' -Value $property
    $obj | Add-Member -MemberType NoteProperty -Name 'Non-Ideal' -Value $nonIdeal
    $obj | Add-Member -MemberType NoteProperty -Name 'Ideal' -Value $ideal
    $obj | Add-Member -MemberType NoteProperty -Name 'Diff' -Value (([math]::round(((($ideal - $nonIdeal) / $nonIdeal)*100),2)) + '%')
    Export-csv -InputObject $obj -Path ($exportDir + '\stats.csv') -Append -NoTypeInformation
}

###################
# global variables
###################

# timestamp in filename friendly format
$timestamp = Get-Date -Format yyyy-MM-hh_hh-mm-ss_tt

# append timestamp to end of export folder name
$exportFolderName = $exportFolderName + '_' + $timestamp

# root export directory
$exportDir = $exportPath + '\' + $exportFolderName

# create an array to store all the data
$reports = @()

# create an array to store ideal data
$idealReports = @()

# create an array to store nonideal data
$nonIdealReports = @()

###################
# gather data
###################

# gather all the folders that arent errors or archive
$folders = Get-ChildItem -Path $sourcePath -Directory | Where-Object {$_.Name -ne "Errors"} | Where-Object {$_.Name -ne "_Archive"}

# loop through all the folders
foreach ($folder in $folders)
{
    # create an array to store all current department's reports
    $department = @()

    # create an array to store all current department's ideal reports
    $idealDepartment = @()

    # create an array to store all current department's nonideal reports
    $nonIdealDepartment = @()

    # get all the report names and filepaths
    $gci = Get-ChildItem -Path $folder.fullname -File

    # loop through each report
    foreach ($report in $gci)
    {
        # create csv object of the report
        $csv = Import-Csv -Path $report.fullname

        # determine computername based on report name
        $computerName = $report.name.Replace('_iPerf.csv','')

        # create an array to store current report
        $report = @()

        # create an array to store current ideal report
        $idealReport = @()

        # create an array to store current nonideal report
        $nonIdealReport = @()

        # entry counter
        [int]$entryCount = '0'

        # loop through each entry in the csv
        foreach ($entry in $csv)
        {
            # check to see if either down or up speeds are 0 or null
            if (($entry.'Client Reverse Receive Rate (Mb/s)' -eq '0') -or ($entry.'Client Send Rate (Mb/s)' -eq '0')){break}

            # find the percent of link speed for both download and upload
            $downPercent = [math]::Round((($entry.'Client Reverse Receive Rate (Mb/s)' / $entry.'Link Speed (Mb/s)')*100),2)
            $upPercent = [math]::Round((($entry.'Client Send Rate (Mb/s)' / $entry.'Link Speed (Mb/s)')*100),2)

            # set a variable for both down and up to classify as "good" or "bad"
            if ($downPercent -ge $goodLinkPercent){$downState = 'good'} else {$downState = 'bad'}
            if ($upPercent -ge $goodLinkPercent){$upState = 'good'} else {$upState = 'bad'}

            # check to see if client is connected to an aruba AP via the ip address
            if (($entry.'Source IP'.split('.')[0] -eq '10') -and (($entry.'Source IP'.split('.')[1] -ge '104') -and ($entry.'Source IP'.split('.')[1] -le '107'))){$aruba = $true}else{$aruba = $false}

            # check to see if current entry is in ideal state
            if (($entry.Band -eq '5GHz') -and ($entry.'Wireless Adapter AC Power Setting' -eq 'Maximum Performance') -and ($entry.'AC Power Status' -eq 'Online') -and ($entry.SSID -eq $targetSSID)){$ideal = $true} else {$ideal = $false}

            # check for radio mismatch
            if (($entry.'Adapter Name' -match 'AC') -and ($entry.'Radio Type' -notmatch 'AC')){$radioMismatch = $true} else {$radioMismatch = $false}

            # check to see what type of connection
            if ($entry.BSSID -eq ''){$connectionType = 'Ethernet'}
            elseif (($entry.'Client Reverse Receive Rate (Mb/s)' -gt $entry.'Link Speed (Mb/s)') -or ($entry.'Client Send Rate (Mb/s)' -gt $entry.'Link Speed (Mb/s)')){$connectionType = 'Ethernet'}
            elseif ($entry.'Adapter Name' -match 'cisco'){$connectionType = 'VPN'}
            elseif ($entry.'Adapter Name' -match 'Virtual'){$connectionType = 'Virtual'}
            elseif ($entry.'Adapter Name' -match 'Hyper-V'){$connectionType = 'Virtual'}
            elseif ($entry.'Adapter Name' -match 'Realtek'){$connectionType = 'Unknown'}
            elseif ($entry.'Adapter Name' -match 'Adapter Name'){$connectionType = 'Unknown'}
            else {$connectionType = 'WiFi'}

            # if previous entry on same BSSID and WiFi
            if (($entry.BSSID -eq $csv[($entryCount -1)].BSSID) -and ($connectionType -eq 'WiFi') -and ($entryCount -ne '0'))
            {
                # calculate difference in down and up speed
                $downDifferencePercent = [math]::round(((($entry.'Client Reverse Receive Rate (Mb/s)' - $csv[($entryCount -1)].'Client Reverse Receive Rate (Mb/s)') / $csv[($entryCount -1)].'Client Reverse Receive Rate (Mb/s)')*100),2)
                $upDifferencePercent = [math]::round(((($entry.'Client Send Rate (Mb/s)' - $csv[($entryCount -1)].'Client Send Rate (Mb/s)') / $csv[($entryCount -1)].'Client Send Rate (Mb/s)')*100),2)

                # check to see if either down or up difference is below target speed drop percent
                if (($downDifferencePercent -lt $speedDropPercent) -or ($upDifferencePercent -lt $speedDropPercent)){$inconsistant = $true}else{$inconsistant = $false}
            }
            else {$inconsistant = $false}

            # create an hashtable of the entry's data
            $entryHashtable = @{
                "Computer Name" = $computerName;
                "Timestamp" = $entry.'Time Stamp';
                "Client Send Rate (Mb/s)" = ([math]::Round(($entry.'Client Send Rate (Mb/s)'),2));
                "Client Reverse Receive Rate (Mb/s)" = ([math]::round(($entry.'Client Reverse Receive Rate (Mb/s)'),2));
                "Link Speed (Mb/s)" = $entry.'Link Speed (Mb/s)';
                "SSID" = $entry.SSID;
                "Radio Type" = $entry.'Radio Type';
                "Channel" = $entry.Channel;
                "Band" = $entry.Band;
                "Signal Strength" = $entry.'Signal Strength';
                "BSSID" = $entry.BSSID;
                "MAC Address" = $entry.'MAC Address';
                "Source IP" = $entry.'Source IP';
                "Adapter Name" = $entry.'Adapter Name';
                "Driver Version" = $entry.'Driver Version';
                "Driver Provider" = $entry.'Driver Provider';
                "Driver Date" = $entry.'Driver Date';
                "OS Name" = $entry.'OS Name';
                "OS Version" = $entry.'OS Version';
                "Power Plan" = $entry.'Power Plan';
                "AC Power Status" = $entry.'AC Power Status';
                "Wireless Adapter AC Power Setting" = $entry.'Wireless Adapter AC Power Setting';
                "Wireless Adapter DC Power Setting" = $entry.'Wireless Adapter DC Power Setting';
                "Username" = $entry.Username;
                "iPerf Server" = $entry.'iPerf Server';
                "Model" = $entry.Model;
                "Chassis Type" = $entry.'Chassis Type';

                "Department" = $computerName.Split('-')[0];

                "Down State" = $downState;
                "Up State" = $upState;

                "Percent Of Link (Up)" = $upState;
                "Percent Of Link (Down)" = $downPercent;

                "Up Percent" = $upPercent;
                "Down Percent" = $downPercent;

                "Ideal" = $ideal;

                "Radio Mismatch" = $radioMismatch;

                "Connection Type" = $connectionType;

                "Inconsistant" = $inconsistant;
                "Down Difference (%)" = $downDifferencePercent;
                "Up Difference (%)" = $upDifferencePercent
            }
            
            # add this entry to the computer's report array
            $report += [PSCustomObject]$entryHashtable

            # check if type is wifi
            if($connectionType -eq 'WiFi')
            {
                # check if aruba is true
                if ($aruba -eq $true)
                {
                    # check if ideal is true
                    if ($ideal -eq $true){$idealReport += [PSCustomObject]$entryHashtable}
                    else {$nonIdealReport += [PSCustomObject]$entryHashtable}
                }
            }

            # increase entry count by 1
            $entryCount++
        }

        # check to make sure current report has at least x days of entries
        if ((New-TimeSpan -Start $report[0]."timestamp" -End $report[-1]."TimeStamp").Days -ge $minReportDaySpan)
        {
            # calculate overall percent good for down/up
            $reportDownPercent = [math]::round(((($report.'down state'.where({$_ -eq 'good'})).count / $report.Count)*100),2)
            $reportUpPercent = [math]::round(((($report.'up state'.where({$_ -eq 'good'})).count / $report.Count)*100),2)

            # calculate percent inconsistant
            $inconsistantPercent = [math]::round(((($report.'inconsistant'.where({$_ -eq $true})).count / $report.inconsistant.Count)*100),2)

            $downMax = ($report.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Maximum).Maximum
            $downMin = ($report.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Minimum).Minimum
            $upMax = ($report.'Client Send Rate (Mb/s)' | Measure-Object -Maximum).Maximum
            $upMin = ($report.'Client Send Rate (Mb/s)' | Measure-Object -Minimum).Minimum
            $downSpread = $downMax - $downMin
            $upSpread = $upMax - $upMin

            # create an object and add it to the larger object
            $computer = New-Object pscustomobject
            $computer | Add-Member -MemberType NoteProperty -Name 'Name' -Value $computerName
            $computer | Add-Member -MemberType NoteProperty -Name 'Percent Good (Down)' -Value $reportDownPercent
            $computer | Add-Member -MemberType NoteProperty -Name 'Percent Good (Up)' -Value $reportUpPercent
            $computer | Add-Member -MemberType NoteProperty -Name 'Inconsistant Count' -Value ($report.inconsistant.where({$_ -eq $true})).count
            $computer | Add-Member -MemberType NoteProperty -Name 'Inconsistant Percent' -Value $inconsistantPercent
            $computer | Add-Member -MemberType NoteProperty -Name 'Down Max' -Value $downMax
            $computer | Add-Member -MemberType NoteProperty -Name 'Down Min' -Value $downMin
            $computer | Add-Member -MemberType NoteProperty -Name 'Up Max' -Value $upMax
            $computer | Add-Member -MemberType NoteProperty -Name 'Up Min' -Value $upMin
            $computer | Add-Member -MemberType NoteProperty -Name 'Down Spread' -Value $downSpread
            $computer | Add-Member -MemberType NoteProperty -Name 'Up Spread' -Value $upSpread
            $computer | Add-Member -MemberType NoteProperty -Name 'Report' -Value $report
            $department += $computer

            if ($idealReport -ne $null)
            {
                # calculate overall percent good for down/up
                $idealReportDownPercent = [math]::round(((($idealReport.'down state'.where({$_ -eq 'good'})).count / $idealReport.Count)*100),2)
                $idealReportUpPercent = [math]::round(((($idealReport.'up state'.where({$_ -eq 'good'})).count / $idealReport.Count)*100),2)

                # calculate percent inconsistant
                $idealInconsistantPercent = [math]::round(((($idealReport.'inconsistant'.where({$_ -eq $true})).count / $idealReport.inconsistant.Count)*100),2)

                $downMax = ($idealReport.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Maximum).Maximum
                $downMin = ($idealReport.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Minimum).Minimum
                $upMax = ($idealReport.'Client Send Rate (Mb/s)' | Measure-Object -Maximum).Maximum
                $upMin = ($idealReport.'Client Send Rate (Mb/s)' | Measure-Object -Minimum).Minimum
                $downSpread = $downMax - $downMin
                $upSpread = $upMax - $upMin

                # create an object and add it to the larger object
                $idealComputer = New-Object pscustomobject
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Name' -Value $computerName
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Percent Good (Down)' -Value $idealReportDownPercent
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Percent Good (Up)' -Value $idealReportUpPercent
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Inconsistant Count' -Value ($idealReport.inconsistant.where({$_ -eq $true})).count
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Inconsistant Percent' -Value $idealInconsistantPercent
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Down Max' -Value $downMax
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Down Min' -Value $downMin
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Up Max' -Value $upMax
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Up Min' -Value $upMin
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Down Spread' -Value $downSpread
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Up Spread' -Value $upSpread
                $idealComputer | Add-Member -MemberType NoteProperty -Name 'Report' -Value $idealReport
                $idealDepartment += $idealComputer
            }

            if ($nonIdealReport -ne $null)
            {
                # calculate overall percent good for down/up
                $nonIdealReportDownPercent = [math]::round(((($nonIdealReport.'down state'.where({$_ -eq 'good'})).count / $nonIdealReport.Count)*100),2)
                $nonIdealReportUpPercent = [math]::round(((($nonIdealReport.'up state'.where({$_ -eq 'good'})).count / $nonIdealReport.Count)*100),2)

                # calculate percent inconsistant
                $nonIdealInconsistantPercent = [math]::round(((($nonIdealReport.'inconsistant'.where({$_ -eq $true})).count / $nonIdealReport.inconsistant.Count)*100),2)

                $downMax = ($nonIdealReport.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Maximum).Maximum
                $downMin = ($nonIdealReport.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Minimum).Minimum
                $upMax = ($nonIdealReport.'Client Send Rate (Mb/s)' | Measure-Object -Maximum).Maximum
                $upMin = ($nonIdealReport.'Client Send Rate (Mb/s)' | Measure-Object -Minimum).Minimum
                $downSpread = $downMax - $downMin
                $upSpread = $upMax - $upMin

                # create an object and add it to the larger object
                $nonIdealComputer = New-Object pscustomobject
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Name' -Value $computerName
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Percent Good (Down)' -Value $nonIdealReportDownPercent
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Percent Good (Up)' -Value $nonIdealReportUpPercent
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Inconsistant Count' -Value ($nonIdealReport.inconsistant.where({$_ -eq $true})).count
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Inconsistant Percent' -Value $nonIdealInconsistantPercent
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Down Max' -Value $downMax
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Down Min' -Value $downMin
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Up Max' -Value $upMax
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Up Min' -Value $upMin
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Down Spread' -Value $downSpread
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Up Spread' -Value $upSpread
                $nonIdealComputer | Add-Member -MemberType NoteProperty -Name 'Report' -Value $nonIdealReport
                $nonIdealDepartment += $nonIdealComputer
            }
        }
    }
    # add current department to overall 
    $reports += $department

    $idealReports += $idealDepartment

    $nonIdealReports += $nonIdealDepartment
}

###################
# export data
###################

$createExportFolder = New-Item -Path $exportPath -Name $exportFolderName -ItemType directory

Export-Stats -property "# reports"-nonideal $nonIdealReports.count -ideal $idealReports.count
Export-Stats -property "# '100% good (down/up)' reports" -nonideal ($nonIdealReports.where({($_.'Percent Good (Down)' -eq '100') -and ($_.'Percent Good (Up)' -eq '100')})).count -ideal ($idealReports.where({($_.'Percent Good (Down)' -eq '100') -and ($_.'Percent Good (Up)' -eq '100')})).count

Export-Stats -property "Average % time 'download speed = good'" -nonideal (([math]::Round((($nonIdealReports.'Percent Good (Down)' | Measure-Object -Average).average),2))) -ideal (([math]::Round((($idealReports.'Percent Good (Down)' | Measure-Object -Average).average),2)))
Export-Stats -property "Average % time 'upload speed = good'" -nonideal (([math]::Round((($nonIdealReports.'Percent Good (Up)' | Measure-Object -Average).average),2))) -ideal ([math]::Round((($idealReports.'Percent Good (Up)' | Measure-Object -Average).average),2))
Export-Stats -property "download speed (Mb/s)" -nonideal ([math]::round((($nonIdealReports.report.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Average).Average),2)) -ideal ([math]::round((($idealReports.report.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Average).Average),2))
Export-Stats -property "upload speed (Mb/s)" -nonideal ([math]::round((($nonIdealReports.report.'Client Send Rate (Mb/s)' | Measure-Object -Average).Average),2)) -ideal ([math]::round((($idealReports.report.'Client Send Rate (Mb/s)' | Measure-Object -Average).Average),2))

#Export-Stats -property "download speed (Mb/s)" -measure "max" -nonideal ($nonIdealReports.report.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Maximum).Maximum -ideal ($idealReports.report.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Maximum).Maximum
#Export-Stats -property "download speed (Mb/s)" -measure "min" -nonideal ($nonIdealReports.report.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Minimum).Minimum -ideal ($idealReports.report.'Client Reverse Receive Rate (Mb/s)' | Measure-Object -Minimum).Minimum
#Export-Stats -property "upload speed (Mb/s)" -measure "max" -nonideal ($nonIdealReports.report.'Client Send Rate (Mb/s)' | Measure-Object -Maximum).Maximum -ideal ($idealReports.report.'Client Send Rate (Mb/s)' | Measure-Object -Maximum).Maximum
#Export-Stats -property "upload speed (Mb/s)" -measure "min" -nonideal ($nonIdealReports.report.'Client Send Rate (Mb/s)' | Measure-Object -Minimum).Minimum -ideal ($idealReports.report.'Client Send Rate (Mb/s)' | Measure-Object -Minimum).Minimum

if ((($nonIdealReports.'Down Max').count % 2) -eq '0') 
{
    $sorted = $nonIdealReports.'Down Max' | Sort-Object -Descending
    $nonIdealMedianDownMax = $sorted[($nonIdealReports.count / 2)]
}
else 
{
    $sorted = $nonIdealReports.'Down Max' | Sort-Object -Descending
    [int]$half = ((($nonIdealreports.'Down Max'.count /2).tostring()).split('.')[0])
    $sorted[$half]
}

Export-Stats -property "reports with at least 1 radio mismatch" -nonideal ($nonIdealReports.report.where({$_.'Radio Mismatch' -eq $true}) | Group 'computer name').count -ideal ($idealReports.report.where({$_.'Radio Mismatch' -eq $true}) | Group 'computer name').count