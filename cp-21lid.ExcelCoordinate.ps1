function ConvertFrom-ExcelCellCoordinate {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [Alias("i")]
        [string]$InputObject
    )
    process{
        if ($InputObject -notmatch "^(?<Column>[a-z]{1,3})(?<Row>[0-9]+)$") {
            throw [Exception] "Invalid Input format"
        }

        $cell = [hashtable]@{
            'Column' = $Matches['Column']
            'Row' = $Matches['Row']
        }
        Convert-CharToIndex -i $cell.Column
    }
}

function Convert-CharToIndex {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [Alias("i")]
        [string]$Column
    )
    process{
        [int[]] $column_chars =  $Column.ToUpper().ToCharArray()
        [Array]::Reverse($column_chars)
        for ($i = 0; $i -lt $Column.Length; $i++) {
            $column_chars[$i] = (($column_chars[$i]) - 64)
        }
        Convert-IndexToInt -i $column_chars
    }
}

function Convert-IndexToInt{
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [Alias("i")]
        [int[]]$column_chars
    )
    process{
        $column_number = 0
        for ($i = 0; $i -lt $column_chars.Count; $i++) {
            $column_number += $column_chars[$i] * [Math]::Pow(26, $i)
        }
    }
}
