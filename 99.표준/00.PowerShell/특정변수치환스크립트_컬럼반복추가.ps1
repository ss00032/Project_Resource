# ==========================================
# 설정
# ==========================================
$sampleFile = "C:\Batch\sample.sql"
$outputDir  = "C:\Batch\output"

$connectionString = "DSN=MY_DSN;UID=MY_USER;PWD=MY_PASSWORD"

# 테이블 및 컬럼 정보 조회 SQL
$query = @"
SELECT
    TABLE_NAME,
    COLUMN_NAME,
    DATA_LENGTH,
    COLUMN_ID
FROM USER_TAB_COLUMNS
WHERE TABLE_NAME = 'MY_TABLE'
ORDER BY COLUMN_ID
"@

# ==========================================
# 샘플 SQL 읽기
# ==========================================
$sampleSql = [System.IO.File]::ReadAllText(
    $sampleFile,
    [System.Text.Encoding]::UTF8
)

New-Item -ItemType Directory -Path $outputDir -Force |
    Out-Null

# ==========================================
# DB 조회
# ==========================================
$conn = New-Object System.Data.Odbc.OdbcConnection(
    $connectionString
)

try {
    $conn.Open()

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $query

    $adapter = New-Object System.Data.Odbc.OdbcDataAdapter($cmd)
    $data = New-Object System.Data.DataTable

    [void]$adapter.Fill($data)

    if ($data.Rows.Count -eq 0) {
        throw "DB 조회 결과가 없습니다."
    }

    # 테이블별로 그룹화
    $tableGroups = $data.Rows |
        Group-Object { $_["TABLE_NAME"].ToString() }

    foreach ($group in $tableGroups) {

        $table = $group.Name
        $columns = @($group.Group)

        # ----------------------------------
        # 1. 반복 구문 생성
        # ----------------------------------
        $caseLines = for ($i = 0; $i -lt $columns.Count; $i++) {
            $column = $columns[$i]["COLUMN_NAME"].ToString()
            $number = $i + 1

            "        WHEN $column IS NOT NULL THEN $number"
        }

        $caseList = $caseLines -join "`r`n"

        # ----------------------------------
        # 2. 일반 변수 치환
        # ----------------------------------
        $values = @{
            table     = $table
            column    = $columns[0]["COLUMN_NAME"].ToString()
            length    = $columns[0]["DATA_LENGTH"].ToString()
            case_list = $caseList
        }

        $sql = $sampleSql

        foreach ($key in $values.Keys) {
            $sql = $sql.Replace(
                "{$key}",
                [string]$values[$key]
            )
        }

        # ----------------------------------
        # 3. 파일 저장
        # ----------------------------------
        $outputFile = Join-Path $outputDir "$table.sql"

        [System.IO.File]::WriteAllText(
            $outputFile,
            $sql,
            (New-Object System.Text.UTF8Encoding($false))
        )

        Write-Host "생성 완료: $outputFile"
    }
}
finally {
    if ($conn.State -ne "Closed") {
        $conn.Close()
    }

    $conn.Dispose()
}