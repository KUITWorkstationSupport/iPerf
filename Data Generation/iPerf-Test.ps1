#v1.0 by Calvin Schulte 06/2018.
	#1.1 Added driver version and date properties to the export
	#1.2 Removed Get-NetAdapter and replaced with a couple WMI queries so we could use it on Win7 without upgrading Powershell first.
	#1.3 Now runs iPerf in parallel mode, reverse mode, and forward/default mode. Added power plan info, chassis type info, SSID/BSSID/Signal info, user info and Windows info. 
	#1.4 Added last boot time and last sleep-wake time

#Configuration
$ExportFileName = $env:COMPUTERNAME + "_iPerf.csv"
$ExportFileNameError = $env:COMPUTERNAME + "_iPerfErrors.csv"
$ExportDirPath = "\\cfs.home.ku.edu\IT_General\Units\SS\ITT2\iPerf\"
$ExportDirPathError = "\\cfs.home.ku.edu\IT_General\Units\SS\ITT2\iPerf\Errors\"
$ExportFilePath = $ExportDirPath + $ExportFileName
$ExportFilePathError = $ExportDirPathError + $ExportFileNameError
$iPerfEXE = "$env:ProgramData\iPerf\iPerf3.exe"
$iPerfServer = "iperf.ku.edu"
$iPerfTime = "5"             #Duration in seconds to run the iPerf tests.

if (Test-Path $iPerfEXE)
{
	$CPUCores = Get-WmiObject -Class Win32_processor | Select-Object -ExpandProperty NumberOfCores
	$CPUCoresTotal = ($CPUCores | Measure-Object -Sum).Sum
	$iPerfThreads = [math]::Ceiling($CPUCoresTotal/2)		#Physical Core Count, divided by 2, round up to integer. iPerf gives results more in line with a generic web speedtest in parallel test mode. 
															#We arbitrarily picked half the physical cores as the marginal speed increases in parallel mode diminish quickly and we want to as low-impact as possible.

	#iPerf runs in 2 different modes. Client to server, the default, and server to client, which is "reverse".
		$iPerfSendResults = Invoke-Expression -Command "$iPerfEXE --client $iPerfServer --time $iPerfTime --parallel $iPerfThreads --json" | Out-String | ConvertFrom-Json
		$iPerfReverseResults = Invoke-Expression -Command "$iPerfEXE --client $iPerfServer --time $iPerfTime --parallel $iPerfThreads --reverse --json" | Out-String | ConvertFrom-Json
		$iPerfSourceIP = $iperfSendResults.start.connected.local_host | Select-Object -First 1
		$TimeStamp = Get-Date -Format g

		#Send specific values.
		$iPerfSendError = $iperfSendResults.error
		$iPerfSendBitrateClient = $iperfSendResults.end.sum_sent.bits_per_second               #The bitrate the iPerf client sent at.
		$iPerfSendBitrateClientMegabits = $iPerfSendBitrateClient/1000000
		$iPerfSendBitrateServer = $iperfSendResults.end.sum_received.bits_per_second           #The bitrate the iPerf server received at.
		$iPerfSendBitrateServerMegabits = $iPerfSendBitrateServer/1000000

		#Reverse specific values. Note since this is reversed, the physical client receives and the iPerf server sends.
		$iPerfReverseError = $iPerfReverseResults.error
		$iPerfReverseBitrateClient = $iPerfReverseResults.end.sum_received.bits_per_second     #The bitrate the iPerf client recieved at.
		$iPerfReverseBitrateClientMegabits = $iPerfReverseBitrateClient/1000000                
		$iPerfReverseBitrateServer = $iPerfReverseResults.end.sum_sent.bits_per_second         #The bitrate the iPerf server sent at.
		$iPerfReverseBitrateServerMegabits = $iPerfReverseBitrateServer/1000000		
}else{
	Write-Error -Message "iPerf not found!"
	exit 1
}

#Adapter/Driver info.
#Yes, I know Get-NetAdapter is a thing, and yes it has all this data. However, we wanted to maintain Win7 OOB compat without going and upgrading WMF on thousands of machines first.
$AdapterIndex = (Get-WmiObject -Class win32_networkadapterconfiguration | Where-Object {$_.IPAddress -contains $iPerfSourceIP}).Index
$Adapter = Get-WmiObject -Class win32_networkadapter | Where-Object {$_.Index -eq $AdapterIndex}
$AdapterName = $Adapter.ProductName
$AdapterMACAddress = $Adapter.MACAddress
$DriverProperties = Get-WmiObject -Class win32_pnpsigneddriver -Filter "DeviceClass = 'NET'" | Where-Object {$_.DeviceName -eq $AdapterName} | Select-Object -First 1
$DriverVersion = $DriverProperties.DriverVersion
$DriverProvider = $DriverProperties.DriverProviderName
$DriverDate = $DriverProperties.ConvertToDateTime($DriverProperties.DriverDate) | Get-Date -Format d

#netsh info
	#Just pipe the netsh WLAN output to a regex that selects everything and sets up capture groups for the fields.
	#Then build an object containing all the netsh output. 
	#Each field in the netsh output is called a "Match" in regex terms. "Groups" are the capture groups defined in the regex itself. Group 1 in each capture would be the name of the field, Group 2 would be the value. Then trim off the whitespace.
$netsh = netsh.exe wlan show interface
$Parsednetsh = $netsh | Select-String -Pattern "^.\s+(.*)\s+:\s(.*)$"
$netshObj = New-Object System.Object
	foreach ($match in $Parsednetsh.Matches) 
	{
		$netshObj | Add-Member -MemberType NoteProperty -Name ($match.Groups[1].Value).Trim() -Value ($match.Groups[2].Value).Trim()
	}
$SSID = $netshObj.SSID
$BSSID = $netshObj.BSSID
$SignalStrength = $netshObj.Signal
$WirelessMAC = $netshobj.'Physical address'
$RadioType = $netshObj.'Radio type'
$Channel = $netshObj.Channel
	if ($Channel -in 1..14)
		{
			$Band = "2.4GHz"
		}
	if ($Channel -in 36..165)
		{
			$Band = "5GHz"
		}

#Link Speed.
	#For wireless adapters, the transmit rate from netsh is more accurate than the speed value from win32_networkadapter.
	#For wired adapters, the win32_networkadapter speed is just fine.
	#Basically, just see if the adapter given by netsh is the same one we found earlier by matching adapters to the iPerf-reported source IP. If it is, that means wireless was used for the test and thus we should use the link speed from netsh.
if ($null -ne $SSID)
{
	if ($WirelessMAC -eq $AdapterMACAddress)
	{
		$LinkSpeed = $netshObj.'Transmit rate (Mbps)'
	}
	else{
		#This is just to account for scenarios where you have active wired and wireless connections, and blank out the wireless-specific fields because they aren't relevant when the test was on wired.
		$LinkSpeed = ($Adapter.Speed)/1000000
		$SSID = $null
		$RadioType = $null
		$Channel = $null
		$Band = $null
		$BSSID = $null
		$SignalStrength = $null
	}
}else{
	$LinkSpeed = ($Adapter.Speed)/1000000
}

#Model/Chassis info. The Chassis Type thing is just getting the value/values (a system can be multiple types), make an array, convert the types to friendly names, the spit it out comma-separated.
$Model = Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model
$ChassisTypes = @(Get-WmiObject -Class Win32_SystemEnclosure | Select-Object -ExpandProperty ChassisTypes)
$ChassisTypeFriendlyNames = @()
ForEach ($ChassisType in $ChassisTypes) 
{
	switch ($ChassisType)   #From https://blogs.technet.microsoft.com/brandonlinton/2017/09/15/updated-win32_systemenclosure-chassis-types/
	{
		"1" {$ChassisTypeFriendlyNames += "Other"}
		"2" {$ChassisTypeFriendlyNames += "Unknown"}
		"3" {$ChassisTypeFriendlyNames += "Desktop"}
		"4" {$ChassisTypeFriendlyNames += "Low Profile Desktop"}
		"5" {$ChassisTypeFriendlyNames += "Pizza Box"}
		"6" {$ChassisTypeFriendlyNames += "Mini Tower"}
		"7" {$ChassisTypeFriendlyNames += "Tower"}
		"8" {$ChassisTypeFriendlyNames += "Portable"}
		"9" {$ChassisTypeFriendlyNames += "Laptop"}
		"10" {$ChassisTypeFriendlyNames += "Notebook"}
		"11" {$ChassisTypeFriendlyNames += "Hand Held"}
		"12" {$ChassisTypeFriendlyNames += "Docking Station"}
		"13" {$ChassisTypeFriendlyNames += "All in One"}
		"14" {$ChassisTypeFriendlyNames += "Sub Notebook"}
		"15" {$ChassisTypeFriendlyNames += "Space-Saving"}
		"16" {$ChassisTypeFriendlyNames += "Lunch Box"}
		"17" {$ChassisTypeFriendlyNames += "Main System Chassis"}
		"18" {$ChassisTypeFriendlyNames += "Expansion Chassis"}
		"19" {$ChassisTypeFriendlyNames += "SubChassis"}
		"20" {$ChassisTypeFriendlyNames += "Bus Expansion Chassis"}
		"21" {$ChassisTypeFriendlyNames += "Peripheral Chassis"}
		"22" {$ChassisTypeFriendlyNames += "RAID Chassis"}
		"23" {$ChassisTypeFriendlyNames += "Rack Mount Chassis"}
		"24" {$ChassisTypeFriendlyNames += "Sealed-case PC"}
		"25" {$ChassisTypeFriendlyNames += "Multi-system chassis"}
		"26" {$ChassisTypeFriendlyNames += "Compact PCI"}
		"27" {$ChassisTypeFriendlyNames += "Advanced TCA"}
		"28" {$ChassisTypeFriendlyNames += "Blade"}
		"29" {$ChassisTypeFriendlyNames += "Blade Enclosure"}
		"30" {$ChassisTypeFriendlyNames += "Tablet"}
		"31" {$ChassisTypeFriendlyNames += "Convertible"}
		"32" {$ChassisTypeFriendlyNames += "Detachable"}
		"33" {$ChassisTypeFriendlyNames += "IoT Gateway"}
		"34" {$ChassisTypeFriendlyNames += "Embedded PC"}
		"35" {$ChassisTypeFriendlyNames += "Mini PC"}
		"36" {$ChassisTypeFriendlyNames += "Stick PC"}
	}
}
$ChassisTypeFriendlyNameOutput = ($ChassisTypeFriendlyNames | Group-Object | Select-Object -ExpandProperty Name) -join ","

#Power plan and AC state info.
Add-Type -Assembly 'System.Windows.Forms' -ErrorAction 'SilentlyContinue'
$PowerInfo = ([Windows.Forms.PowerStatus]$PowerInfo = [Windows.Forms.SystemInformation]::PowerStatus)
$PowerState = $PowerInfo.PowerLineStatus            #"Online" or "Offline" in reference to if AC power is present or not.
$PowerPlanName = Get-WmiObject -Class win32_powerplan -Namespace root\cimv2\power -Filter "isActive=True" | Select-Object -ExpandProperty ElementName
$PowerPlanSettings = Invoke-Expression -Command "powercfg.exe /query"
	$WirelessAdapterPowerSettings = [regex]::Matches($PowerPlanSettings,"\(Wireless Adapter Settings\).*?Current AC Power Setting Index:\s(.{10}).*?Current DC Power Setting Index:\s(.{10})")
	$WirelessAdapterACPowerSetting = $WirelessAdapterPowerSettings.Groups[1].Value
	$WirelessAdapterDCPowerSetting = $WirelessAdapterPowerSettings.Groups[2].Value
		switch($WirelessAdapterACPowerSetting)
			{
				"0x00000000" {$WirelessAdapterACPowerSettingFriendlyName = "Maximum Performance"}
				"0x00000001" {$WirelessAdapterACPowerSettingFriendlyName = "Low Power Saving"}
				"0x00000002" {$WirelessAdapterACPowerSettingFriendlyName = "Medium Power Saving"}
				"0x00000003" {$WirelessAdapterACPowerSettingFriendlyName = "Maximum Power Saving"}
				$null {$WirelessAdapterACPowerSettingFriendlyName = "N/A"}
			}
		switch($WirelessAdapterDCPowerSetting)
		{
			"0x00000000" {$WirelessAdapterDCPowerSettingFriendlyName = "Maximum Performance"}
			"0x00000001" {$WirelessAdapterDCPowerSettingFriendlyName = "Low Power Saving"}
			"0x00000002" {$WirelessAdapterDCPowerSettingFriendlyName = "Medium Power Saving"}
			"0x00000003" {$WirelessAdapterDCPowerSettingFriendlyName = "Maximum Power Saving"}
			$null {$WirelessAdapterDCPowerSettingFriendlyName = "N/A"}
		}
		
#User info.
$Username = Get-WMIObject -class Win32_ComputerSystem | Select-Object -ExpandProperty username
$FriendlyUsername = ($Username.split("\"))[1]

#OS Info
$OS = Get-CimInstance -ClassName Win32_OperatingSystem
$OSName = $OS.Caption
$OSVersion = $OS.Version
$OSLastBoot = $OS.LastBootUpTime

#Sleep-wake info and wake source.
#Get-WinEvent instead of Get-EventLog because Get-WinEvent will translate the Wake Source into something human friendly by default.
#Not all machines will actually have slept before, especially VMs and desktops.
$WinEvent = Get-WinEvent -FilterHashTable @{LogName="System";ProviderName="Microsoft-Windows-Power-TroubleShooter"} -MaxEvents 1 -ErrorAction SilentlyContinue
if ($null -ne $WinEvent)
{
	$WakeTime = [DateTime]$WinEvent.Properties[1].Value
	$ParsedMessage = $WinEvent.Message | Select-String -Pattern "(?m)^Wake Source: (.*)$"
	$WakeSource = $ParsedMessage.Matches[0].Groups[1].Value
}


#Department is a close approximation for building, so split the computer name and re-set the export path to a folder of the same name.
#If computer name has a dash; if the department folder doesn't exist, create it; if the department folder is present 
if ($env:COMPUTERNAME.Contains("-"))
{
	$Department = $env:COMPUTERNAME.Split("-")[0]
	$DepartmentDir = $ExportDirPath+$Department+"\"
	New-Item -Path $DepartmentDir -ItemType Directory -Force
	if (Test-Path -Path $DepartmentDir) 
	{
		$ExportDirPath = $DepartmentDir
		$ExportFilePath = $ExportDirPath + $ExportFileName
	}
}

#Export the results.
if (($null -eq $iPerfSendError) -and ($null -eq $iPerfReverseError))
{
	$Obj = New-Object System.Object
	$Obj | Add-Member -MemberType NoteProperty -Name 'Time Stamp' -Value $TimeStamp
	$Obj | Add-Member -MemberType NoteProperty -Name 'Client Send Rate (Mb/s)' -Value $iPerfSendBitrateClientMegabits
	$Obj | Add-Member -MemberType NoteProperty -Name 'Client Reverse Receive Rate (Mb/s)' -Value $iPerfReverseBitrateClientMegabits
	$Obj | Add-Member -MemberType NoteProperty -Name 'Link Speed (Mb/s)' -Value $LinkSpeed
	$Obj | Add-Member -MemberType NoteProperty -Name 'SSID' -Value $SSID
	$Obj | Add-Member -MemberType NoteProperty -Name 'Radio Type' -Value $RadioType
	$Obj | Add-Member -MemberType NoteProperty -Name 'Channel' -Value $Channel
	$Obj | Add-Member -MemberType NoteProperty -Name 'Band' -Value $Band
	$Obj | Add-Member -MemberType NoteProperty -Name 'Signal Strength' -Value $SignalStrength
	$Obj | Add-Member -MemberType NoteProperty -Name 'BSSID' -Value $BSSID
	$Obj | Add-Member -MemberType NoteProperty -Name 'MAC Address' -Value $AdapterMACAddress
	$Obj | Add-Member -MemberType NoteProperty -Name 'Source IP' -Value $iPerfSourceIP
	$Obj | Add-Member -MemberType NoteProperty -Name 'Adapter Name' -Value $AdapterName
	$Obj | Add-Member -MemberType NoteProperty -Name 'Driver Version' -Value $DriverVersion
	$Obj | Add-Member -MemberType NoteProperty -Name 'Driver Provider' -Value $DriverProvider
	$Obj | Add-Member -MemberType NoteProperty -Name 'Driver Date' -Value $DriverDate
	$Obj | Add-Member -MemberType NoteProperty -Name 'OS Name' -Value $OSName
	$Obj | Add-Member -MemberType NoteProperty -Name 'OS Version' -Value $OSVersion
	$Obj | Add-Member -MemberType NoteProperty -Name 'OS Last Boot' -Value $OSLastBoot
	$Obj | Add-Member -MemberType NoteProperty -Name 'OS Last Wake' -Value $WakeTime
	$Obj | Add-Member -MemberType NoteProperty -Name 'OS Last Wake Source' -Value $WakeSource
	$Obj | Add-Member -MemberType NoteProperty -Name 'Power Plan' -Value $PowerPlanName
	$Obj | Add-Member -MemberType NoteProperty -Name 'AC Power Status' -Value $PowerState
	$Obj | Add-Member -MemberType NoteProperty -Name 'Wireless Adapter AC Power Setting' -Value $WirelessAdapterACPowerSettingFriendlyName
	$Obj | Add-Member -MemberType NoteProperty -Name 'Wireless Adapter DC Power Setting' -Value $WirelessAdapterDCPowerSettingFriendlyName
	$Obj | Add-Member -MemberType NoteProperty -Name 'Username' -Value $FriendlyUsername
	$Obj | Add-Member -MemberType NoteProperty -Name 'iPerf Server' -Value $iPerfServer
	$Obj | Add-Member -MemberType NoteProperty -Name 'Model' -Value $Model
	$Obj | Add-Member -MemberType NoteProperty -Name 'Chassis Type' -Value $ChassisTypeFriendlyNameOutput
	Export-Csv -InputObject $Obj -Path $ExportFilePath -Append -NoTypeInformation
}else{
	Write-Error "$iPerfError"
	Write-Error "$iPerfReverseError"
	$errorObj = New-Object System.Object
	$errorObj | Add-Member -MemberType NoteProperty -Name 'Time Stamp' -Value $TimeStamp
	$errorObj | Add-Member -MemberType NoteProperty -Name 'Adapter Name' -Value $AdapterName
	$errorObj | Add-Member -MemberType NoteProperty -Name 'MAC Address' -Value $AdapterMACAddress
	$errorObj | Add-Member -MemberType NoteProperty -Name 'iPerf Send Error' -Value $iPerfSendError
	$errorObj | Add-Member -MemberType NoteProperty -Name 'iPerf Reverse Error' -Value $iPerfReverseError
	Export-Csv -Path $ExportFilePathError -InputObject $errorObj -Append
	exit 2
}