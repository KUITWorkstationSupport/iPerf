#Just get the iPerf CSVs and move anything that doesn't have "Driver Version" as a column to the Archive folder.

#Paths setup.
$importPath = "\\cfs.home.ku.edu\IT_General\Units\SS\ITT2\iPerf\ITS\"
$importFiles = Get-ChildItem -Path $importPath -Name -Depth 0 -File

foreach ($file in $importFiles)
{
    $filePath = $importPath + $file
    $csv = Import-Csv -Path $filePath
    if ($csv[0].'OS Name' -ne "Microsoft Windows 10 Enterprise") 
    {
        write-host $filePath -ForegroundColor Green
    }
}