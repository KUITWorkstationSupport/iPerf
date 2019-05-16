#Just get the iPerf CSVs and move anything that doesn't have "Driver Version" as a column to the Archive folder.

#Paths setup.
$iPerfReportPath = "\\cfs.home.ku.edu\IT_General\Units\SS\ITT2\iPerf\"
$iPerfReportDirs = Get-ChildItem -Path $iPerfReportPath -Directory -Exclude "_*","Errors"
$importFiles = @()

#Get a list of files in each folder in the iPerf folder.
foreach ($directory in $iPerfReportDirs.FullName) 
{
    $importFiles += Get-ChildItem -Path $directory -Depth 0 -File  
}

#Import each file, see if it has the column
foreach ($file in $importFiles)
{
    $csv = Import-Csv -Path $file.FullName
    $csvProperties = $csv | Get-Member
    if ($csvProperties.Name -notcontains "OS Last Wake") 
    {
        Write-Host $file.FullName "does NOT have the column." -ForegroundColor Red
        Remove-Item $file.FullName
    }
}