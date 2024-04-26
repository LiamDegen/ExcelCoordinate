task Test {
    $basePath = Get-Location
 
    $scriptPath = [System.IO.Path]::Join($basePath, "cp-21lid.ExcelCoordinate.ps1")
    $testPath = [System.IO.Path]::Join($basePath, "cp-21lid.ExcelCoordinateTests.ps1")
 
    . $scriptPath
    . $testPath
 
    Invoke-Pester -Output Detailed
}
 
task Release {
    ConvertFrom-ExcelCellCoordinate -i 'A1'
}
 
task . Test, Release