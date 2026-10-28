try {
    $sql = @"
    SELECT T10.TABLENAME
         , T10.COLUMNNAME
      FROM DBC.ColumnsV T10
"@

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $sql
    $reader = $cmd.ExecuteReader()
    $rowcount = $reader.RecordsAffected
    $map = @{}

    while ($reader.Read()) {
        $TABLENAME = $reader["TABLENAME"]
        $COLUMNNAME = $reader["COLUMNNAME"]

        if (-not $map.ContainsKey($TABLENAME)) {
            $map[$TABLENAME] = @{
                Columns     = @()
                BaseDtType  = $null
                BaseDtCol   = $null
                BaseDtParam = $null
                tabComment  = $TABCOMMENTSTRING
                colLength   = $COLUMNLENGTH
            }
        }

        <#if ($COLUMNNAME -eq "CRTR_YMD") {
            $map[$TABLENAME].BaseDtType  = "D"
            $map[$TABLENAME].BaseDtCol   = "CRTR_YMD"
            $map[$TABLENAME].BaseDtParam = "'{BASE_YMD}'"
        }#>

        $colObj = [PSCustomObject]@{
            columnName = $reader["COLUMNNAME"]
            #caseName = if ($COLUMNNAME.Contains("ERR_YN")) {("WHEN ERR_COL = '" + $COLUMNNAME + "'").PadRight(70) + "THEN " + $ERR_COLUMN_ID} else { "" }
            #pivotName = if ($COLUMNNAME.Contains("ERR_YN")) {$COLUMNNAME} else {}
            #Length = [int]$COLUMNLENGTH
        }
        $map[$TABLENAME].Columns += $colObj
    }

    foreach ($TABLENAME in $map.Keys) {
        $meta = $map[$TABLENAME]
        $colNames = $meta.Columns | ForEach-Object { $_.columnName }
        $dttype = $meta.BaseDtType

        #$casecolNames = $meta.Columns | ForEach-Object { $_.caseName }
        #$pivotcolNames = $meta.Columns | ForEach-Object { $_.pivotName }

        $columnList = $colNames -join "`n ,"
        #$caseList = $casecolNames -join "`n"
        #$pivotList = $pivotcolNames -join "`n"

        $A = @(
            "내용"
        )

        $B = @(
            "내용"
        )

        $odsscript = ($A + $selectSql + $B) -join "`n"

        $savePath = "C:\Users\FSS_NT551XDA\Documents\금융감독원DX01_작업용\00.SQL\00.TPT_EXMUC" + $dttype + "_" + $TABLENAME + "_00.sql"

        #$odsscript > $savePath
        [System.IO.File]::WriteAllText($savePath,$odsscript,$utf8NoBom)

        $odsscript = ""
    }
}
catch {
    Write-Host $_
    break
}
finally {
    $reader.Close()
    $conn.Close()
}