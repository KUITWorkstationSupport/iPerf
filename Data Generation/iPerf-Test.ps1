<#
.SYNOPSIS
	To run a variety of iPerf tests.
.DESCRIPTION
	To both gather data about the endpoint and run a variety of iPerf tests against a specified iPerf server, finally exporting data as a CSV.
.PARAMETER iPerfServer
	Specify the iPerf server to use for the tests.
	ie: "iperf.ku.edu"
.PARAMETER iPerfTime
	Specify duration in seconds of how long each iPerf test should run. 
.PARAMETER iPerfMinPort
	Specify lower end of the port range iPerf should be randomly pointed at.
.PARAMETER iPerfMaxPort
	Specify upper end of the port range iPerf should be randomly pointed at. 
.PARAMETER ExportDir
	Specify UNC path to where the results should be exported.
	ie: "\\cfs.home.ku.edu\IT_General\Units\SS\ITT2\iPerf"
	PROTIP: Make sure you give the machine permissions to this path if you run it as a scheduled task.
.INPUTS
	None.
.OUTPUTS
	Exports a CSV to path specified in $ExportDir parameter.
.NOTES
	Author: Calvin Schulte, University of Kansas
	Released: 2018-06-14
	Updated: 2018-06-15 -- Added driver version and date properties to the export
	Updated: 2018-06-19 -- Removed Get-NetAdapter and replaced with a couple WMI queries so we could use it on Win7 without upgrading Powershell first.
	Updated: 2018-06-15 -- Now runs iPerf in parallel mode, reverse mode, and forward/default mode. Added power, chassis type, SSID/BSSID/Signal, user, Windows, and other related fields.
	Updated: 2018-06-29 -- Added radio type and channel number fields, updated code for grabbing OS version.
	Updated: 2019-01-30 -- Added last boot timestamp, last sleep timestamp, last wake source fields
	Updated: 2019-05-09 -- Added forward and reverse UDP tests, added forward jitter, forward packet loss, reverse packet loss. 
	Updated: 2019-05-14 -- Added comment based help and mandatory parameters for iPerf server and Export Path.
	Updated: 2019-11-18 -- Added port range parameters.
.LINK
	Download iPerf: https://iperf.fr/iperf-download.php
	iPerf Documentation: https://iperf.fr/iperf-doc.php
#>

[CmdletBinding()]
PARAM
(
	[string]$iPerfServer = "iperf.ku.edu",
	[int]$iPerfTime = "5",
	[int]$iPerfMinPort = 5201,
	[int]$iPerfMaxPort = 5300,
	[string]$ExportDir = "\\cfs.home.ku.edu\it_general\Units\SS\ITT2\iPerf\"
)
Write-Verbose -Message "iPerf Server: $iPerfServer"
Write-Verbose -Message "Test Duration: $iPerfTime seconds"
Write-Verbose -Message "Min port: $iPerfMinPort"
Write-Verbose -Message "Max port: $iPerfMaxPort"
Write-Verbose -Message "Export Path: $ExportDir"


#region configuration
$ExportFileName = $env:COMPUTERNAME + "_iPerf.csv"
$ExportFilePath = $ExportDir + $ExportFileName
$iPerfEXEPath = (Split-Path -Path $script:MyInvocation.MyCommand.Path -Parent) + "\iperf3.exe"
Write-Debug "Export File Name: $ExportFileName"
Write-Debug "Export File Path: $ExportFilePath"
Write-Debug "iPerf Executable Path: $iPerfEXEPath"
#endregion configuration

#region iPerfTests
if (Test-Path $iPerfEXEPath -PathType Leaf)
{

	$CPUCores = Get-WmiObject -Class Win32_processor | Select-Object -ExpandProperty NumberOfCores
    $port = Get-Random -minimum 5201 -maximum 5300 #Randomize the port for testing in this range: 5201-5300
	$CPUCoresTotal = ($CPUCores | Measure-Object -Sum).Sum
	$iPerfThreads = [math]::Ceiling($CPUCoresTotal/2)		#We arbitrarily picked half the physical cores as the marginal speed increases in parallel mode diminish quickly and we want to as low-impact as possible.
		Write-Debug "CPU Cores Detected: $CPUCores"
		Write-Verbose "iPerf threads/test: $iPerfThreads"

	$TimeStamp = Get-Date -Format g
		Write-Verbose "iPerf Testing started at $TimeStamp"

	#We run iPerf 4 times. Forwards and backwards in TCP mode and again in UDP mode. Note in "Reverse" mode, the client initiates a connection, but then the server sends and the client receives.
	#Note that "--get-server-output" merely ADDS the literal console output from the test on the server's side to the json output. It does NOT parse it in JSON, it literally just adds a string containing the output.
	
	#region forwardTCPTest
	#While we don't gather TCP retransmit data, it was thought about. We found that iPerf does only generates that data in reverse mode and so wasn't super helpful.
	Write-Verbose "Running iPerf in fowards TCP mode"
	$iPerfSendTCPResults = Invoke-Expression -Command "$iPerfEXEPath --client $iPerfServer --port $port --time $iPerfTime --parallel $iPerfThreads --get-server-output --json" | Out-String | ConvertFrom-Json
		Write-Debug "iPerf Command: $iPerfEXEPath --client $iPerfServer --port $port --time $iPerfTime --parallel $iPerfThreads --get-server-output --json"
	if ($iPerfSendTCPResults.error)
		{
			Write-Error "iPerf Test failed: TCP Send"
			Write-Debug $iPerfSendTCPResults.error
			exit 2
		}
	Write-Debug $iPerfSendTCPResults.end.sum_received
	Write-Debug $iPerfSendTCPResults.end.sum_sent
	Write-Debug $iPerfSendTCPResults.server_output_text
	$iPerfSourceIP = $iPerfSendTCPResults.start.connected.local_host | Select-Object -First 1		#We grab the source IP actually used by iPerf to compare against the adapters to make sure we are getting info about the right one.
	$iPerfSendTCPBitrateClient = $iPerfSendTCPResults.end.sum_sent.bits_per_second					#The bitrate (in bits) the iPerf client sent at.
	$iPerfSendTCPBitrateClientMegabits = $iPerfSendTCPBitrateClient/1000000
	#endregion forwardTCPTest

	#region reverseTCPTest
	Write-Verbose "Running iPerf in reverse TCP mode"
	$iPerfReverseTCPResults = Invoke-Expression -Command "$iPerfEXEPath --client $iPerfServer --port $port --time $iPerfTime --parallel $iPerfThreads --reverse --get-server-output --json" | Out-String | ConvertFrom-Json
		Write-Debug "iPerf Command: $iPerfEXEPath --client $iPerfServer --port $port --time $iPerfTime --parallel $iPerfThreads --reverse --get-server-output --json"
	if ($iPerfReverseTCPResults.error)
	{
		Write-Error "iPerf Test failed: TCP Reverse"
		Write-Debug $iPerfReverseTCPResults.error
		exit 2
	}
	Write-Debug $iPerfReverseTCPResults.end.sum_received
	Write-Debug $iPerfReverseTCPResults.end.sum_sent
	Write-Debug $iPerfReverseTCPResults.server_output_text
	$iPerfReverseTCPBitrateClient = $iPerfReverseTCPResults.end.sum_received.bits_per_second		#The bitrate the iPerf client recieved at.
	$iPerfReverseTCPBitrateClientMegabits = $iPerfReverseTCPBitrateClient/1000000
	#endregion reverseTCPTest

	#region forwardUDPTest
	Write-Verbose "Running iPerf in forwards UDP mode"
	$iPerfSendUDPResults = Invoke-Expression -Command "$iPerfEXEPath --client $iPerfServer --port $port --time $iPerfTime --parallel $iPerfThreads --udp --get-server-output --json" | Out-String | ConvertFrom-Json
		Write-Debug "iPerf Command: $iPerfEXEPath --client $iPerfServer --port $port --time $iPerfTime --parallel $iPerfThreads --udp --get-server-output --json"
	if ($iPerfSendUDPResults.error)
	{
		Write-Error "iPerf Test failed: UDP Send"
		Write-Debug $iPerfSendUDPResults.error
		exit 2
	}
	#Note we specifically only wanted packet loss and jitter on the UDP Send Test.
	Write-Debug $iPerfSendUDPResults.end.sum
	Write-Debug $iPerfSendUDPResults.server_output_text
	$iPerfSendUDPPacketLoss = $iPerfSendUDPResults.end.sum.lost_percent
	$iPerfSendUDPJitter = $iPerfSendUDPResults.end.sum.jitter_ms
	#endregion forwardUDPTest

	#region reverseUDPTest
	Write-Verbose "Running iPerf in reverse UDP mode"
	$iPerfReverseUDPResults = Invoke-Expression -Command "$iPerfEXEPath --client $iPerfServer --port $port --time $iPerfTime --parallel $iPerfThreads --reverse --udp --get-server-output --json" | Out-String | ConvertFrom-Json
		Write-Debug "iPerf Command: $iPerfEXEPath --client $iPerfServer --port $port --time $iPerfTime --parallel $iPerfThreads --reverse --udp --get-server-output --json"
	if ($iPerfReverseUDPResults.error)
	{
		Write-Error "iPerf Test failed: UDP Reverse"
		exit 2
	}
	#Note we specifically only wanted packet loss and jitter on UDP.
	#Note that in a reverse UDP test, the client does NOT record packet loss or jitter. This must be parsed out of the server's output instead. Probably a bug, https://github.com/esnet/iperf/issues/584
	#Also note that in some (maybe all) circumstances in a reverse UDP test, neither the client nor the server record jitter correctly. Left commented out in case we want to use it in the future.
	#Regex example for the Jitter: https://regex101.com/r/qiDDSz/11
	#Regex example for the Packet Loss: https://regex101.com/r/qiDDSz/10
	Write-Debug $iPerfReverseUDPResults.end.sum
	Write-Debug $iPerfReverseUDPResults.server_output_text
	#$iPerfReverseUDPJitter = ($iPerfReverseUDPResults.server_output_text | Select-String -Pattern "\[SUM\].*\/sec\s*(\d*\.\d*)").Matches.Groups[1].Value
	$iPerfReverseUDPPacketLoss = ($iPerfReverseUDPResults.server_output_text | Select-String -Pattern "\[SUM\].*\((\d.*)%\)").Matches.Groups[1].Value
	#endregion reverseUDPTest
}else{
	Write-Error -Message "iPerf not found!"
	exit 1
}
#endregion iPerfTests

#region AdapterInfo
#Yes, I know Get-NetAdapter is a thing, and yes it has all this data. However, we wanted to maintain Win7 OOB compat without going and upgrading WMF on thousands of machines first.
$AdapterIndex = (Get-WmiObject -Class win32_networkadapterconfiguration | Where-Object {$_.IPAddress -contains $iPerfSourceIP}).Index
$Adapter = Get-WmiObject -Class win32_networkadapter | Where-Object {$_.Index -eq $AdapterIndex}
$AdapterName = $Adapter.ProductName
$AdapterMACAddress = $Adapter.MACAddress
$DriverProperties = Get-WmiObject -Class win32_pnpsigneddriver -Filter "DeviceClass = 'NET'" | Where-Object {$_.DeviceName -eq $AdapterName} | Select-Object -First 1
$DriverVersion = $DriverProperties.DriverVersion
$DriverProvider = $DriverProperties.DriverProviderName
$DriverDate = $DriverProperties.ConvertToDateTime($DriverProperties.DriverDate) | Get-Date -Format d
	Write-Debug "Test Adapter Index: $AdapterIndex"
	Write-Verbose "Test Adapter Name: $AdapterName"
	Write-Verbose "Test Adapter Driver Version: $DriverVersion"
	Write-Verbose "Test Adapter Driver Date: $DriverDate"
	Write-Verbose "Test Adapter MAC Addreess: $AdapterMACAddress"
#endregion AdapterInfo

#region netsh
	#Just pipe the netsh WLAN output to a regex that selects everything and sets up capture groups for the fields.
	#Then build an object containing all the netsh output. 
	#Each field in the netsh output is called a "Match" in regex terms. "Groups" are the capture groups defined in the regex itself. Group 1 in each capture would be the name of the field, Group 2 would be the value. Then trim off the whitespace.
	#Regex example: https://regex101.com/r/uMGBAA/1
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
	Write-Debug "netsh output: $netsh"
#endregion netsh

#region determine_Wireless/Wired
	#Basically, just see if the adapter given by netsh is the same one we found earlier by matching adapters to the iPerf-reported source IP. If it is, that means wireless was used for the test.
	#For wireless adapters, the transmit rate from netsh is more accurate than the speed value from win32_networkadapter.
	#For wired adapters, the win32_networkadapter speed is just fine.
if ($null -ne $SSID)
{
	Write-Debug "Wireless Adapter MAC: $WirelessMAC"
	Write-Debug "MAC of adapter used by iPerf: $AdapterMACAddress"
	if ($WirelessMAC -eq $AdapterMACAddress)
	{
		$LinkSpeed = $netshObj.'Transmit rate (Mbps)'
		Write-Verbose "SSID is $SSID"
		Write-Verbose "Linkspeed: $LinkSpeed"
	}
	else{
		#This is just to account for scenarios where you have active wired and wireless connections, and blank out the wireless-specific fields because they aren't relevant when the test was on wired.
		$LinkSpeed = ($Adapter.Speed)/1000000
		Write-Verbose "Linkspeed: $LinkSpeed"
		$SSID = $null
		$RadioType = $null
		$Channel = $null
		$Band = $null
		$BSSID = $null
		$SignalStrength = $null
	}
}else{
	$LinkSpeed = ($Adapter.Speed)/1000000
	Write-Verbose "Linkspeed: $LinkSpeed"
}
#endregion determine_Wireless/Wired

#region Get-ChassisType
#Model/Chassis info. The Chassis Type thing is just getting the value/values (a system can be multiple types), make an array, convert the types to friendly names, the spit it out comma-separated.
$Model = Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model
$ChassisTypes = @(Get-WmiObject -Class Win32_SystemEnclosure | Select-Object -ExpandProperty ChassisTypes)
$ChassisTypeFriendlyNames = @()
	Write-Verbose "Model: $Model"
	Write-Debug "Chassis Type(s): $ChassisTypes"
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
Write-Verbose "Chassis Type: $ChassisTypeFriendlyNameOutput"
#endregion Get-ChassisType

#region Get-PowerInfo
Add-Type -Assembly 'System.Windows.Forms' -ErrorAction 'SilentlyContinue'
$PowerInfo = ([Windows.Forms.PowerStatus]$PowerInfo = [Windows.Forms.SystemInformation]::PowerStatus)
$PowerState = $PowerInfo.PowerLineStatus            #"Online" or "Offline" in reference to if AC power is present or not.
$PowerPlanName = Get-WmiObject -Class win32_powerplan -Namespace root\cimv2\power -Filter "isActive=True" | Select-Object -ExpandProperty ElementName
$PowerPlanSettings = Invoke-Expression -Command "powercfg.exe /query" | Out-String
	Write-Verbose "AC Power: $PowerState"
	Write-Debug $PowerPlanSettings

#Regex example: https://regex101.com/r/m47B39/6
$WirelessAdapterPowerSettings = $PowerPlanSettings | Select-String -Pattern "(?s)\(Wireless Adapter Settings\).*?Current AC Power Setting Index:\s(.{10}).*?Current DC Power Setting Index:\s(.{10})"
$WirelessAdapterACPowerSetting = $WirelessAdapterPowerSettings.Matches.Groups[1].Value
$WirelessAdapterDCPowerSetting = $WirelessAdapterPowerSettings.Matches.Groups[2].Value
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
Write-Verbose "Wireless Adapter AC Power Level: $WirelessAdapterACPowerSettingFriendlyName"
Write-Verbose "Wireless Adapter DC Power Level: $WirelessAdapterDCPowerSettingFriendlyName"
#endregion Get-PowerInfo

#region Get-UserInfo
#User info.
$Username = Get-WMIObject -class Win32_ComputerSystem | Select-Object -ExpandProperty username
$FriendlyUsername = ($Username.split("\"))[1]
Write-Verbose "Logged on user: $FriendlyUsername"
#endregion Get-UserInfo

#region Get-OSInfo
$OS = Get-CimInstance -ClassName Win32_OperatingSystem
$OSName = $OS.Caption
$OSVersion = $OS.Version
$OSLastBoot = $OS.LastBootUpTime
Write-Verbose "OS Name: $OSName"
Write-Verbose "OS Version: $OSVersion"
#endregion Get-OSInfo

#region Get-WakeSource
#Sleep-wake info and wake source.
#Get-WinEvent instead of Get-EventLog because Get-WinEvent will translate the Wake Source into something human friendly by default.
#Not all machines will actually have slept before, especially VMs and desktops.
$WinEvent = Get-WinEvent -FilterHashTable @{LogName="System";ProviderName="Microsoft-Windows-Power-TroubleShooter"} -MaxEvents 1 -ErrorAction SilentlyContinue
if ($null -ne $WinEvent)
{
	$WakeTime = [DateTime]$WinEvent.Properties[1].Value
	$ParsedMessage = $WinEvent.Message | Select-String -Pattern "(?m)^Wake Source: (.*)$"
	$WakeSource = $ParsedMessage.Matches[0].Groups[1].Value
	Write-Debug $WinEvent
}
#endregion Get-WakeSource

#region departmentexport
#Department is a close approximation for building, so split the computer name and re-set the export path to a folder of the same name.
#If computer name has a dash; if the department folder doesn't exist, create it; if the department folder is present 
if ($env:COMPUTERNAME.Contains("-"))
{
	$Department = $env:COMPUTERNAME.Split("-")[0]
		Write-Verbose "Assuming department is $Department"
	$DepartmentDir = $ExportDir+"\"+$Department+"\"
		Write-Debug "Writing file to $DepartmentDir"
	New-Item -Path $DepartmentDir -ItemType Directory -Force
	if (Test-Path -Path $DepartmentDir)
	{
		$ExportFilePath = $DepartmentDir + $ExportFileName
	}
}
#endregion departmentexport

#region export
Write-Verbose "Exporting at $ExportFilePath"
$Obj = New-Object System.Object
$Obj | Add-Member -MemberType NoteProperty -Name 'Time Stamp' -Value $TimeStamp
$Obj | Add-Member -MemberType NoteProperty -Name 'Client TCP Send Rate (Mb/s)' -Value $iPerfSendTCPBitrateClientMegabits
$Obj | Add-Member -MemberType NoteProperty -Name 'Client TCP Reverse Receive Rate (Mb/s)' -Value $iPerfReverseTCPBitrateClientMegabits
$Obj | Add-Member -MemberType NoteProperty -Name 'Client UDP Send Packet Loss (%)' -Value $iPerfSendUDPPacketLoss
$Obj | Add-Member -MemberType NoteProperty -Name 'Client UDP Reverse Packet Loss (%)' -Value $iPerfReverseUDPPacketLoss
$Obj | Add-Member -MemberType NoteProperty -Name 'Client UDP Send Jitter (ms)' -Value $iPerfSendUDPJitter
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
#endregion export