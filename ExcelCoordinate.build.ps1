task Test {
    $basePath = Get-Location
 
    $scriptPath = [System.IO.Path]::Join($basePath, "cp-21lid.ExcelCoordinate.ps1")
    $testPath = [System.IO.Path]::Join($basePath, "cp-21lid.ExcelCoordinateTests.ps1")
 
    . $scriptPath
    . $testPath
 
    Import-Module -Name Pester
    $config = New-PesterConfiguration
    $config.Run.Path = $testPath
    $config.CodeCoverage.Enabled = $true
    $config.CodeCoverage.Path = $scriptPath
    $config.CodeCoverage.CoveragePercentTarget = 80
    $config.CodeCoverage.OutputFormat = 'CoverageGutters'
    $config.CodeCoverage.OutputPath = 'cov.xml'
    $config.CodeCoverage.OutputEncoding = 'UTF8'

    Invoke-Pester -Configuration $config

}
 
task Release {

}
 
task . Test, Release