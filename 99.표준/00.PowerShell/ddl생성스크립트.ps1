# ============================================================
# Excel COM 기반 Oracle DDL 생성
# PowerShell 5.1 호환
# ============================================================

$excelPath = "C:\Users\윤혁준\Documents\수신(독자)시스템구축\99.CreatePath\테이블정의서.xlsx"
$outputDir = "C:\Users\윤혁준\Documents\수신(독자)시스템구축\99.CreatePath\DDL"
$sheetName = "Sheet1"
$schema = "ODS"

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$tableMap = @{}

if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

$excel = $null
$workbook = $null
$worksheet = $null

try {
    # --------------------------------------------------------
    # 1. Excel 열기
    # --------------------------------------------------------

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($excelPath)
    $worksheet = $workbook.Worksheets.Item($sheetName)

    $usedRange = $worksheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    # 첫 번째 행을 헤더로 사용
    $headers = @{}

    for ($c = 1; $c -le $colCount; $c++) {
        $header = [string]$worksheet.Cells.Item(1, $c).Text

        if (-not [string]::IsNullOrWhiteSpace($header)) {
            $headers[$header.Trim().ToUpper()] = $c
        }
    }

    # 필요한 헤더 확인
    $requiredHeaders = @(
        "테이블명",
        "COLUMNNAME",
        "COLUMNTYPE"
    )

    foreach ($header in $requiredHeaders) {
        if (-not $headers.ContainsKey($header)) {
            throw "엑셀에 필수 컬럼이 없습니다: $header"
        }
    }

    # Excel 셀 값 가져오기
    function Get-ExcelValue {
        param(
            $Sheet,
            [hashtable]$HeaderMap,
            [int]$Row,
            [string]$Column
        )

        $key = $Column.ToUpper()

        if (-not $HeaderMap.ContainsKey($key)) {
            return $null
        }

        $value = $Sheet.Cells.Item(
            $Row,
            $HeaderMap[$key]
        ).Value2

        if ($null -eq $value) {
            return $null
        }

        return ([string]$value).Trim()
    }

    # --------------------------------------------------------
    # 2. Excel 행을 테이블별로 그룹화
    # --------------------------------------------------------

    for ($r = 2; $r -le $rowCount; $r++) {

        $table = Get-ExcelValue $worksheet $headers $r "테이블명"
        $column = Get-ExcelValue $worksheet $headers $r "COLUMNNAME"
        $dataType = Get-ExcelValue $worksheet $headers $r "COLUMNTYPE"

        if ([string]::IsNullOrWhiteSpace($table) -or
            [string]::IsNullOrWhiteSpace($column)) {
            continue
        }

        if (-not $tableMap.ContainsKey($table)) {
            $tableMap[$table] = New-Object 'System.Collections.Generic.List[object]'
        }

        $length = Get-ExcelValue $worksheet $headers $r "COLUMNLENGTH"
        $precision = Get-ExcelValue $worksheet $headers $r "PRE"
        $scale = Get-ExcelValue $worksheet $headers $r "SCA"
        $nullable = Get-ExcelValue $worksheet $headers $r "NULLABLE"
        $comment = Get-ExcelValue $worksheet $headers $r "COLCOMMENTSTRING"

        $tableMap[$table].Add([PSCustomObject]@{
            Name      = $column
            DataType  = $dataType
            Length    = $length
            Precision = $precision
            Scale     = $scale
            Nullable  = $nullable
            Comment   = $comment
        })
    }

    # --------------------------------------------------------
    # 3. Oracle 데이터 타입 변환
    # --------------------------------------------------------

    function Convert-ToOracleType {
        param($Column)

        $type = ([string]$Column.DataType).Trim().ToUpper()

        switch -Regex ($type) {
            '^(VARCHAR|VARCHAR2|CV|VARBYTE)$' {
                if ($type -eq "VARBYTE") {
                    throw "VARBYTE는 Oracle 타입 매핑을 별도로 지정해야 합니다: $($Column.Name)"
                }

                if (-not $Column.Length) {
                    throw "길이 누락: $($Column.Name)"
                }

                return "VARCHAR2($($Column.Length))"
            }

            '^(CHAR|CF)$' {
                if (-not $Column.Length) {
                    throw "길이 누락: $($Column.Name)"
                }

                return "CHAR($($Column.Length))"
            }

            '^(NUMBER|NUMERIC|DECIMAL|D)$' {
                if ($Column.Precision) {
                    if ($Column.Scale -and [int]$Column.Scale -gt 0) {
                        return "NUMBER($($Column.Precision),$($Column.Scale))"
                    }

                    return "NUMBER($($Column.Precision))"
                }

                return "NUMBER"
            }

            '^(INTEGER|INT|I)$' {
                return "NUMBER(10)"
            }

            '^(SMALLINT|I2)$' {
                return "NUMBER(5)"
            }

            '^(BIGINT|I8)$' {
                return "NUMBER(19)"
            }

            '^(BYTEINT|I1)$' {
                return "NUMBER(3)"
            }

            '^(DATE|DA)$' {
                return "DATE"
            }

            '^(FLOAT|REAL|DOUBLE|F)$' {
                return "FLOAT"
            }

            default {
                throw "미지원 데이터 타입: $type (컬럼: $($Column.Name))"
            }
        }
    }

    # --------------------------------------------------------
    # 4. 테이블별 CREATE TABLE 생성
    # --------------------------------------------------------

    foreach ($table in $tableMap.Keys) {

        $columns = $tableMap[$table]
        $columnDefinitions = @()

        foreach ($col in $columns) {

            $oracleType = Convert-ToOracleType $col

            $nullClause = ""

            if ([string]$col.Nullable -match '^(N|NO|0)$') {
                $nullClause = " NOT NULL"
            }

            $columnDefinitions += (
                $col.Name + " " + $oracleType + $nullClause
            )
        }

        $createSql = @"
CREATE TABLE $schema.$table
(
   $($columnDefinitions -join "`r`n , ")
);
"@

        $filePath = Join-Path $outputDir "${table}.sql"

        [System.IO.File]::WriteAllText(
            $filePath,
            $createSql,
            $utf8NoBom
        )

        Write-Host "생성 완료: $filePath"
    }

}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
    }

    if ($excel -ne $null) {
        $excel.Quit()
    }

    if ($worksheet -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
    }

    if ($workbook -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($excel -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}