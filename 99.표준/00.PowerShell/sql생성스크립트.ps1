Add-Type -AssemblyName System.Data.OracleClient

$PSDefaultParameterValues['*:Encoding'] = 'Default'

$connectionString = "DSN=MY_DSN;UID=MY_USER;PWD=MY_PASSWORD"
$conn = New-Object System.Data.Odbc.OdbcConnection($connectionString)
$conn.Open()

# ============================================================
# SQL 생성 스크립트
# 전제: $conn 은 Oracle 연결 객체
# ============================================================

$outputDir = "C:\temp\SQL"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$map = @{}

if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force |
        Out-Null
}

$sql = @"
SELECT
    C.DatabaseName AS DATABASENAME,
    C.TableName AS TABLENAME,
    C.ColumnName AS COLUMNNAME,
    C.ColumnId AS COLUMNID,
    C.ColumnLength AS COLUMNLENGTH,
    C.ColumnType AS COLUMNTYPE,
    C.CommentString AS COLCOMMENTSTRING,
    T.CommentString AS TABCOMMENTSTRING
FROM DBC.ColumnsV C
LEFT JOIN DBC.TablesV T
    ON C.DatabaseName = T.DatabaseName
   AND C.TableName = T.TableName
WHERE C.DatabaseName = 'ODS'
ORDER BY C.TableName, C.ColumnId
"@

$cmd = $conn.CreateCommand()
$cmd.CommandText = $sql
$reader = $cmd.ExecuteReader()

try {
    while ($reader.Read()) {

        $table = [string]$reader["TABLENAME"]
        $column = [string]$reader["COLUMNNAME"]
        $columnId = [int]$reader["COLUMNID"]

        if (-not $map.ContainsKey($table)) {
            $map[$table] = @{
                DatabaseName = [string]$reader["DATABASENAME"]
                TableComment = [string]$reader["TABCOMMENTSTRING"]
                Columns = New-Object 'System.Collections.Generic.List[object]'
                BaseDtType = $null
                BaseDtCol = $null
                BaseDtParam = $null
                StgTable = "S" + $table.Substring(1)
            }
        }

        $meta = $map[$table]

        # 기준일자 / 기준년월 판별
        if ($columnId -eq 1025 -and $column -eq "CRTR_YMD") {
            $meta.BaseDtType = "D"
            $meta.BaseDtCol = "CRTR_YMD"
            $meta.BaseDtParam = "{base_dt}"
        }
        elseif ($columnId -eq 1025 -and $column -eq "CRTR_YM") {
            $meta.BaseDtType = "M"
            $meta.BaseDtCol = "CRTR_YM"
            $meta.BaseDtParam = "{base_ym}"
        }

        $meta.Columns.Add([PSCustomObject]@{
            Name = $column
            ColumnId = $columnId
        })
    }
}
finally {
    $reader.Close()
    $cmd.Dispose()
}

foreach ($table in $map.Keys) {

    $meta = $map[$table]

    $targetTable = "$($meta.DatabaseName).$table"
    $stgTable = "STG.$($meta.StgTable)"

    $colNames = @(
        $meta.Columns | ForEach-Object { $_.Name }
    )

    $colList = $colNames -join ",`r`n    "
    $selList = $colNames -join ",`r`n    "

    # DELETE
    $deleteSql = ""

    if ($meta.BaseDtCol) {
        $deleteSql = @"
DELETE
FROM $targetTable
WHERE $($meta.BaseDtCol) = '$($meta.BaseDtParam)';
"@
    }

    # INSERT / SELECT
    $insertSql = @"
INSERT INTO $targetTable
(
    $colList
)
SELECT
    $selList
FROM $stgTable;
"@

    # 프로그램 정보 및 SQL 구분자는 필요에 맞게 추가
    $scriptContent = @"
/********************************************************
$("/* TABLE : $table".PadRight(100))
-- COMMENT: $($meta.TableComment)

$deleteSql
####SQL
$insertSql
"@

    $dateType = $meta.BaseDtType
    if (-not $dateType) {
        $dateType = "N"
    }

    $fileName = "${dateType}_${table}_00.sql"
    $savePath = Join-Path $outputDir $fileName

    [System.IO.File]::WriteAllText($savePath, $scriptContent, $utf8NoBom)

    Write-Host "$table : $savePath"
}

Write-Host "SQL 생성 완료"