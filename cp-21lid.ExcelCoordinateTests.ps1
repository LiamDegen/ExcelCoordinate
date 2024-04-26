Describe 'SeparateChar-FromDigit'{
    It 'Given no parameters, throws an error' {
        $currentCell = ConvertFrom-ExcelCellCoordinate -i "A1"
        $currentCell['Column'] | Should -Be 1
    }
}