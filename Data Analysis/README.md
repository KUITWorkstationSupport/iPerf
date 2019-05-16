# Purpose
Intended to be run against already generated iPerf data to create reports and statistics.

# Required Variables
$Path : Set this to the UNC location for the already generated iPerf data.
$AirWaveScript : Set this to the UNC location for the AirWave.py script that is used to gather AP details.
$BuildingInfoPath : Set this to the UNC path for the Excel spreadsheet containg details on building information.

# Reports Generated
An excel spreadsheet will be generated in the _Stats folder (within the source location of the iPerf data) for each of the buildings in which tests occured. The data will be split out into individual workbooks within the spreadsheet, with a workbook for each month where tests occured, and an overall workbook.

# Report Data Points
1. Scope
2. Entry Count
3. Unique Computer Count
4. % Unique Computers With AC Card
5. % Unique Computers in Ideal State
6. Mean Upload (Mb/s)
7. Mean Download (Mb/s)
8. Percent Of Time Upload Entries Are Good
9. Percent Of Time Download Entries Are Good

