$importPath = "\\cfs.home.ku.edu\IT_General\Units\SS\ITT2\iPerf\"
$importFiles = Get-ChildItem -Path $importPath -Depth 0 -File

foreach ($file in $importFiles) 
{
    $Department = $file.Name.Split("-")[0]
    $Destination = $importPath + $Department
    Write-Host "Moving $file to $Destination"
    New-Item -Path $Destination -ItemType Directory -Force
    Move-Item -Path $file.FullName -Destination $Destination -Force
}