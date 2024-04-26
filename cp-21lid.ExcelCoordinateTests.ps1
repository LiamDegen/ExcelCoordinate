Describe "Resolve-ExcelCoordinate Tests" {
    It "Correctly resolves a valid input 'A1'" {
        $result = Resolve-ExcelCoordinate -i "A1"
        $result['Column'] | Should -Be 'A'
        $result['Row'] | Should -Be '1'
    }

    It "Throws an exception for an invalid input '1A'" {
        { Resolve-ExcelCoordinate -i "1A" } | Should -Throw "Invalid Input format"
    }
}

Describe "Convert-CharToIndex Tests" {
    It "Correctly converts column 'A' to index 1" {
        $result = Convert-CharToIndex -i "A"
        $result[0] | Should -Be 1
    }

    It "Correctly converts column 'AB' to indices 1, 2" {
        $result = Convert-CharToIndex -i "AB"
        $result[0] | Should -Be 2
        $result[1] | Should -Be 1
    }
}

Describe "Convert-IndexToInt Tests" {
    It "Correctly converts single index 1 to column number 1" {
        $result = Convert-IndexToInt -i @(1)
        $result | Should -Be 1
    }

    It "Correctly converts indices 2, 1 (representing 'AB') to column number 28" {
        $result = Convert-IndexToInt -i @(2,1)
        $result | Should -Be 28
    }
}

Describe "ConvertFrom-ExcelCellCoordinate Tests" {
    It "Correctly processes valid input 'A1' and returns column 1 and row 1" {
        $result = ConvertFrom-ExcelCellCoordinate -i "A1"
        $result['Column'] | Should -Be 1
        $result['Row'] | Should -Be 1
    }

    It "Throws an exception for invalid input '1A'" {
        { ConvertFrom-ExcelCellCoordinate -i "1A" } | Should -Throw "Invalid Input format"
    }
}