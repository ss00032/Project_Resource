Add-Type -AssemblyName System.Data.OracleClient

$PSDefaultParameterValues['*:Encoding'] = 'Default'

# 운영계 통합운영정보
$driver = "MySQL ODBC 8.0 ANSI Driver"
$server = "dept-rds-an2-op-dev-opmdsdbd.cluster-c92dc3uwdt0o.ap-northeast-2.rds.amazonaws.com"
$username = "FEDSEL"
$password = "d55!43EL"
$data_source = "MDSDEV"
#$connection_string = "User Id=$username;Password=$password;Data Source=$data_source"
$connectionString = "DRIVER={$driver};Server=$server;Port=3306;Database=$data_source;Uid=$username;Pwd=$password"

# 분석계 개발DB(Redshift)
$redhost = "dept-rsc-an2-cm-dwdev-ansla.ctdnukmkv08j.ap-northeast-2.redshift.amazonaws.com"
$redport = 5439
$reddbName = "bludbev"
$redusername = "etluser08"
$redpassword = "!Etluser08"
$redconnectionString = "Driver={Amazon Redshift (x64)}; Server=$redhost; Port=$redport; Database=$reddbName; UID=$redusername; PWD=$redpassword;"

# 파일 경로
$sampleFile = "C:\Batch\sample.sql"
$outputDir  = "C:\Batch\output"

# 출력 폴더 생성
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

# 샘플 SQL 읽기
$sampleSql = [System.IO.File]::ReadAllText(
    $sampleFile,
    [System.Text.Encoding]::UTF8
)

# DB 연결 설정 예시 - 실제 환경에 맞게 수정
$connectionString = "DSN=MY_DSN;UID=MY_USER;PWD=MY_PASSWORD"

# 실제 DB 조회 SQL로 변경
$query = @"
SELECT
    TABLE_NAME,
    COLUMN_NAME,
    DATA_LENGTH
FROM USER_TAB_COLUMNS
WHERE TABLE_NAME = 'MY_TABLE'
ORDER BY COLUMN_ID
"@

$conn = New-Object System.Data.Odbc.OdbcConnection($connectionString)

try {
    $conn.Open()

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $query

    $adapter = New-Object System.Data.Odbc.OdbcDataAdapter($cmd)
    $data = New-Object System.Data.DataTable
    [void]$adapter.Fill($data)

    foreach ($row in $data.Rows) {
        # DB 조회 결과를 치환 변수에 매핑
        $values = @{
            table  = [string]$row["TABLE_NAME"]
            column = [string]$row["COLUMN_NAME"]
            length = [string]$row["DATA_LENGTH"]
        }

        # 샘플 SQL 복사본에서 지정된 변수만 치환
        $sql = $sampleSql

        foreach ($key in $values.Keys) {
            $sql = $sql.Replace(
                "{$key}",
                $values[$key]
            )
        }

        # 컬럼별 파일명 생성
        $fileName = "{0}_{1}.sql" -f `
            $values["table"], $values["column"]

        $outputFile = Join-Path $outputDir $fileName

        # UTF-8 인코딩으로 저장
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