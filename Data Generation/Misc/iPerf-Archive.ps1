#Just get the iPerf CSVs and move anything that doesn't have "Driver Version" as a column to the Archive folder.

#Paths setup.
$importPath = "\\cfs.home.ku.edu\IT_General\Units\SS\ITT2\iPerf\"
$importFiles = Get-ChildItem -Path $importPath -Name -Depth 0 -File
$archivePath = "\\cfs.home.ku.edu\IT_General\Units\SS\ITT2\iPerf\Archive\"

foreach ($file in $importFiles)
{
    $filePath = $importPath + $file
    $csv = Import-Csv -Path $filePath
    $csvProperties = $csv | Get-Member
    if ($csvProperties.Name -notcontains "Band") 
    {
        write-host "Going to move $file."    
        Move-Item -Path $filePath -Destination $archivePath
    }
}