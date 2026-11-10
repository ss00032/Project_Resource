#=====================================================================
# 엑셀 기준 재귀 계보 분석 - 통합 백업본
#  - 정리해둔 엑셀의 [파일명 / 원천스키마 / 원천테이블 / 타겟스키마·테이블 (최대 3쌍)] 을 읽어서
#  - 타겟테이블 -> 원천 -> 그 원천을 타겟으로 만드는 파일의 원천 -> ... 원천이 없을 때까지 추적
#  - 결과는 새 엑셀 파일로 저장 (원본 엑셀은 수정하지 않음)
#
#  [이 버전에 반영된 수정]
#   1) 데이터는 Range(...).Value2 가 null 로 나오는 환경이라 Cells.Item(행,열).Value2 로 한 칸씩 읽음
#   2) 마지막 행은 UsedRange 기준
#   3) 타겟스키마/타겟테이블 3쌍(타겟, 타겟2, 타겟3) 처리
#   4) 타겟스키마 제외 목록($excludeTgtSchemas)
#   5) 결과 엑셀 쓰기는 Range 로 한 번에 쓰고, 실패하면 한 칸씩 쓰는 방식으로 자동 전환
#=====================================================================
$ErrorActionPreference = "Stop"

#---------------------------------------------------------------
# 설정  (경로는 본인 환경에 맞게 수정)
#---------------------------------------------------------------
$inputExcel = "C:\경로\정리된엑셀.xlsx"        # 정리해둔 엑셀 파일 (디스크에 저장된 상태로 읽음)
$sheetName  = "재귀분석_모수"                   # 읽을 시트 이름 (비우면 헤더가 있는 첫 시트 자동 탐색)
$outputFile = "C:\경로\재귀분석_결과.xlsx"      # 결과 저장 파일
$maxLevel   = 30                               # 재귀 최대 깊이 (안전장치)
$showExcel  = $false                           # $true 로 두면 작업 중 엑셀 창이 보임 (멈췄을 때 대화상자 확인용)

# 제외할 타겟스키마 (타겟스키마 1~3 모두에 적용, 대소문자 무시)
#  - "S_RSK"   : 정확히 S_RSK 인 스키마만 제외
#  - "S_RSK*"  : S_RSK 로 시작하는 스키마 제외  /  "*RSK*" : RSK 가 포함된 스키마 제외
$excludeTgtSchemas = @("S_RSK")

#---------------------------------------------------------------
# 함수
#---------------------------------------------------------------
function Normalize-Header {
    param($v)
    if ($null -eq $v) { return "" }
    return ([string]$v -replace '\s', '')
}

function Get-Cell {
    param($arr, $r, $c)
    if (-not $c) { return "" }
    $v = $arr[$r, $c]
    if ($null -eq $v) { return "" }
    return ([string]$v).Trim()
}

# 스키마 제거 + 따옴표 제거 + 대문자
function Get-BaseName {
    param([string]$name)
    if ([string]::IsNullOrWhiteSpace($name)) { return "" }
    return (($name -split '\.')[-1]).Trim().Trim('"').ToUpper()
}

# 스키마.테이블 키 (스키마 없으면 테이블명만)
function Get-Key {
    param([string]$schema, [string]$table)
    $t = Get-BaseName $table
    if (-not $t) { return "" }
    $s = ([string]$schema).Trim().Trim('"').ToUpper()
    if ($s) { return "$s.$t" }
    return $t
}

# 시트 상단 15행에서 헤더 행 / 컬럼 위치 찾기
function Find-HeaderMap {
    param($ws)
    $top = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(15, 80)).Value2
    for ($r = 1; $r -le 15; $r++) {
        $map = @{}
        for ($c = 1; $c -le 80; $c++) {
            $h = Normalize-Header $top[$r, $c]
            if ($h -and -not $map.ContainsKey($h)) { $map[$h] = $c }
        }
        if ($map.ContainsKey("파일명") -and $map.ContainsKey("원천테이블") -and $map.ContainsKey("타겟테이블")) {
            return [PSCustomObject]@{ Row = $r; Map = $map }
        }
    }
    return $null
}

# 이 테이블을 타겟으로 만드는 행들 찾기
# - 스키마가 있으면 스키마.테이블 로 정확히 매칭 (S_CBK.X 와 S_IDA.X 는 다른 테이블)
# - 스키마가 비어 있으면 테이블명만으로 매칭
function Find-Producers {
    param([string]$schema, [string]$table)
    $t = Get-BaseName $table
    $s = ([string]$schema).Trim().Trim('"').ToUpper()
    if ($s) {
        $k = "$s.$t"
        if ($idxFull.ContainsKey($k)) { return ,$idxFull[$k] }
        return $null
    }
    if ($idxName.ContainsKey($t)) { return ,$idxName[$t] }
    return $null
}

function Trace-Source {
    param(
        [string]$Root,
        [string]$CurSchema,
        [string]$CurTable,
        [int]$Level,
        [string[]]$Path
    )

    if ($Level -gt $maxLevel) { return }

    $prod = Find-Producers $CurSchema $CurTable
    if (-not $prod) { return }

    $curKey = Get-Key $CurSchema $CurTable
    $seen = @{}

    foreach ($r in $prod) {
        $srcKey = Get-Key $r.SrcSchema $r.SrcTable
        if (-not $srcKey) { continue }

        $dup = "$($r.FileName)|$srcKey"
        if ($seen.ContainsKey($dup)) { continue }
        $seen[$dup] = $true

        # 순환 (자기 자신 / 이미 지나온 테이블)
        $isCycle = ($srcKey -eq $curKey) -or ($Path -contains $srcKey)

        # 이 원천을 타겟으로 만드는 파일이 있고, 거기에 원천이 있으면 더 내려감
        $hasMore = $false
        if (-not $isCycle) {
            $next = Find-Producers $r.SrcSchema $r.SrcTable
            if ($next) {
                foreach ($x in $next) {
                    if ($x.SrcTable) { $hasMore = $true; break }
                }
            }
        }

        $newPath = @($Path) + $srcKey

        [void]$lineage.Add(
            [PSCustomObject]@{
                Root         = $Root
                Level        = $Level
                Parent       = $curKey
                FileName     = $r.FileName
                SourceSchema = $r.SrcSchema
                SourceTable  = $r.SrcTable
                SourceKor    = $r.SrcKor
                IsFinal      = $(if ($hasMore) { "" } else { "Y" })
                IsCycle      = $(if ($isCycle) { "Y" } else { "" })
                PathText     = ($newPath -join " <- ")
            }
        )

        if ($hasMore) {
            Trace-Source -Root $Root -CurSchema $r.SrcSchema -CurTable $r.SrcTable -Level ($Level + 1) -Path $newPath
        }
    }
}

# 시트에 표 쓰기: Range 로 한 번에 쓰고, 실패/불일치하면 한 칸씩 쓴다
function Write-Table {
    param($sheet, [string[]]$headers, $rows, [int[]]$widths)

    $n = $rows.Count + 1
    $arr = New-Object 'object[,]' $n, $headers.Count

    for ($c = 0; $c -lt $headers.Count; $c++) { $arr[0, $c] = $headers[$c] }
    for ($i = 0; $i -lt $rows.Count; $i++) {
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $arr[($i + 1), $c] = $rows[$i][$c]
        }
    }

    # 1) 한 번에 쓰기
    $ok = $false
    try {
        $rng = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($n, $headers.Count))
        $rng.Value2 = $arr
        # 쓴 값 검증 (헤더 + 마지막 행 첫 칸)
        if (([string]$sheet.Cells.Item(1, 1).Value2 -eq [string]$arr[0, 0]) -and
            ([string]$sheet.Cells.Item($n, 1).Value2 -eq [string]$arr[($n - 1), 0])) { $ok = $true }
    } catch { $ok = $false }

    # 2) 실패하면 한 칸씩 쓰기
    if (-not $ok) {
        Write-Host "  (Range 로 한 번에 쓰기 실패 -> 한 칸씩 씁니다. 시간이 걸릴 수 있습니다)"
        for ($i = 0; $i -lt $n; $i++) {
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $v = $arr[$i, $c]
                if ($null -ne $v -and "$v" -ne "") {
                    $sheet.Cells.Item(($i + 1), ($c + 1)).Value2 = $v
                }
            }
            if ($i -gt 0 -and $i % 1000 -eq 0) { Write-Host "  쓰는 중 ... $i / $n" }
        }
    }

    # 서식 (실패해도 결과에는 영향 없음)
    try {
        $hr = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item(1, $headers.Count))
        $hr.Font.Bold = $true
        $hr.HorizontalAlignment = -4108

        for ($c = 0; $c -lt $widths.Count; $c++) { $sheet.Columns.Item($c + 1).ColumnWidth = $widths[$c] }

        [void]$sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($n, $headers.Count)).AutoFilter()

        $sheet.Activate()
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.FreezePanes = $true
    } catch {
        Write-Host "  (서식 적용 실패 - 무시하고 계속 진행) $($_.Exception.Message)"
    }
}

#---------------------------------------------------------------
# 실행
#---------------------------------------------------------------
$idxFull = @{}
$idxName = @{}
$rootMap = @{}
$lineage = New-Object System.Collections.ArrayList

$tmp = Join-Path $env:TEMP ("lineage_src_" + [guid]::NewGuid().ToString("N") + ".xlsx")
Copy-Item $inputExcel $tmp -Force

$excel = $null; $srcWb = $null; $outWb = $null; $ws = $null; $sh1 = $null; $sh2 = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $showExcel
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false

    # 1. 입력 엑셀 읽기 -------------------------------------------------
    Write-Host "엑셀 여는 중 ... (파일이 크면 시간이 걸립니다)"
    $srcWb = $excel.Workbooks.Open($tmp, 0, $true)

    $hdr = $null
    foreach ($s in $srcWb.Worksheets) {
        if ($sheetName -and $s.Name -ne $sheetName) { continue }
        $h = Find-HeaderMap $s
        if ($h) { $ws = $s; $hdr = $h; break }
    }
    if (-not $ws) { throw "헤더(파일명 / 원천테이블 / 타겟테이블)를 가진 시트를 찾지 못했습니다. `$sheetName 을 지정하거나 헤더명을 확인하세요." }

    Write-Host "읽는 시트 : $($ws.Name) (헤더 $($hdr.Row)행)"

    $firstData = $hdr.Row + 1

    $cFile   = $hdr.Map["파일명"]
    $cSrc    = $hdr.Map["원천테이블"]
    $cSrcSch = $hdr.Map["원천스키마"]
    $cKor    = $hdr.Map["원천테이블한글명"]

    # 타겟은 최대 3쌍 (스키마 / 테이블) : 타겟, 타겟2, 타겟3
    $tgtCols = @(
        @{ Sch = $hdr.Map["타겟스키마"];  Tbl = $hdr.Map["타겟테이블"]  },
        @{ Sch = $hdr.Map["타겟스키마2"]; Tbl = $hdr.Map["타겟테이블2"] },
        @{ Sch = $hdr.Map["타겟스키마3"]; Tbl = $hdr.Map["타겟테이블3"] }
    )

    if (-not $cSrcSch -or -not $tgtCols[0].Sch) {
        Write-Host "[주의] 원천스키마/타겟스키마 헤더를 못 찾았습니다. 테이블명만으로 매칭합니다."
    }

    # 실제로 읽을 컬럼 (필요한 컬럼만 읽어서 시간 단축)
    $needCols = @($cFile, $cSrc, $cSrcSch, $cKor,
                  $hdr.Map["타겟스키마"],  $hdr.Map["타겟테이블"],
                  $hdr.Map["타겟스키마2"], $hdr.Map["타겟테이블2"],
                  $hdr.Map["타겟스키마3"], $hdr.Map["타겟테이블3"]) | Where-Object { $_ } | Sort-Object -Unique
    $maxCol = ($needCols | Measure-Object -Maximum).Maximum

    # 마지막 행 (UsedRange 기준)
    $used    = $ws.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    Write-Host "데이터 범위 : $firstData ~ $lastRow 행 / 1 ~ $maxCol 열 (읽는 컬럼 $($needCols.Count)개)"
    if ($lastRow -lt $firstData) { throw "데이터 행이 없습니다." }

    # Range(...).Value2 가 null 로 나오는 환경이라 Cells.Item(행,열).Value2 로 한 칸씩 읽는다
    # $data 는 엑셀 배열과 같이 1 부터 시작하는 2차원 배열 ($data[행번호, 시트열번호])
    $nRows = $lastRow - $firstData + 1
    $data  = [Array]::CreateInstance([object], [int[]]@($nRows, $maxCol), [int[]]@(1, 1))

    for ($i = 1; $i -le $nRows; $i++) {
        $rr = $firstData + $i - 1
        foreach ($c in $needCols) {
            $data.SetValue($ws.Cells.Item($rr, $c).Value2, $i, $c)
        }
        if ($i % 500 -eq 0) { Write-Host "  읽는 중 ... $i / $nRows" }
    }
    Write-Host "읽기 완료 : $nRows 행 x $maxCol 열"

    # 2. 타겟테이블 인덱스 --------------------------------------------
    $lo = 1
    $hi = $nRows
    $rowCount = 0
    $excludedCount = 0

    for ($r = $lo; $r -le $hi; $r++) {

        $file   = Get-Cell $data $r $cFile
        $srcSch = (Get-Cell $data $r $cSrcSch).Trim('"').ToUpper()
        $srcTbl = Get-BaseName (Get-Cell $data $r $cSrc)
        $kor    = Get-Cell $data $r $cKor

        foreach ($tc in $tgtCols) {
            if (-not $tc.Tbl) { continue }

            $tgtRaw = Get-Cell $data $r $tc.Tbl
            if (-not $tgtRaw) { continue }

            $tgtSch = (Get-Cell $data $r $tc.Sch).Trim('"').ToUpper()

            # 제외 스키마 판별
            $skip = $false
            foreach ($ex in $excludeTgtSchemas) {
                if ($ex -and $tgtSch -like $ex) { $skip = $true; break }
            }
            if ($skip) { $excludedCount++; continue }

            foreach ($t in ($tgtRaw -split ',')) {
                $tb = Get-BaseName $t
                if (-not $tb) { continue }

                $row = [PSCustomObject]@{
                    FileName  = $file
                    SrcSchema = $srcSch
                    SrcTable  = $srcTbl
                    SrcKor    = $kor
                    TgtSchema = $tgtSch
                    TgtTable  = $tb
                }

                $kf = Get-Key $tgtSch $tb
                if (-not $idxFull.ContainsKey($kf)) { $idxFull[$kf] = New-Object System.Collections.ArrayList }
                [void]$idxFull[$kf].Add($row)

                if (-not $idxName.ContainsKey($tb)) { $idxName[$tb] = New-Object System.Collections.ArrayList }
                [void]$idxName[$tb].Add($row)

                if (-not $rootMap.ContainsKey($kf)) { $rootMap[$kf] = @($tgtSch, $tb) }
                $rowCount++
            }
        }
    }
    Write-Host "읽은 타겟 항목 : $rowCount 개 / 타겟테이블 : $($rootMap.Count) 개 / 스키마 제외 : $excludedCount 개"

    # 3. 재귀 분석 ----------------------------------------------------
    foreach ($k in ($rootMap.Keys | Sort-Object)) {
        Trace-Source -Root $k -CurSchema $rootMap[$k][0] -CurTable $rootMap[$k][1] -Level 1 -Path @($k)
    }
    Write-Host "재귀 분석 결과 : $($lineage.Count) 행"

    # 4. 결과 행 만들기 -----------------------------------------------
    $rows1 = New-Object System.Collections.ArrayList
    foreach ($x in $lineage) {
        [void]$rows1.Add(@($x.Root, $x.Level, $x.Parent, $x.FileName, $x.SourceSchema, $x.SourceTable, $x.SourceKor, $x.IsFinal, $x.IsCycle, $x.PathText))
    }

    # 최종 원천 요약 (최초타겟별로 끝까지 내려간 원천)
    $sumMap = @{}
    foreach ($x in $lineage) {
        if ($x.IsFinal -eq "Y" -and $x.IsCycle -ne "Y") {
            $k = "$($x.Root)|$($x.SourceSchema)|$($x.SourceTable)"
            if (-not $sumMap.ContainsKey($k) -or $sumMap[$k][4] -lt $x.Level) {
                $sumMap[$k] = @($x.Root, $x.SourceSchema, $x.SourceTable, $x.SourceKor, $x.Level)
            }
        }
    }
    $rows2 = New-Object System.Collections.ArrayList
    foreach ($k in ($sumMap.Keys | Sort-Object)) { [void]$rows2.Add($sumMap[$k]) }

    # 5. 결과 엑셀 저장 -----------------------------------------------
    $excel.SheetsInNewWorkbook = 1
    $outWb = $excel.Workbooks.Add()

    $sh1 = $outWb.Worksheets.Item(1)
    $sh1.Name = "재귀분석"
    Write-Table $sh1 @("최초타겟(스키마.테이블)", "레벨", "상위테이블(타겟)", "파일명", "원천스키마", "원천테이블", "원천테이블한글명", "최종원천", "순환", "경로") $rows1 @(32, 8, 32, 45, 14, 32, 30, 10, 8, 90)

    $sh2 = $outWb.Worksheets.Add([System.Reflection.Missing]::Value, $outWb.Worksheets.Item($outWb.Worksheets.Count))
    $sh2.Name = "최종원천요약"
    Write-Table $sh2 @("최초타겟(스키마.테이블)", "최종원천스키마", "최종원천테이블", "원천테이블한글명", "도달레벨") $rows2 @(32, 16, 32, 30, 10)

    $sh1.Activate()

    if (Test-Path $outputFile) { Remove-Item $outputFile -Force }
    $outWb.SaveAs($outputFile, 51)

    Write-Host ""
    Write-Host "================================="
    Write-Host "재귀분석 엑셀 생성 완료"
    Write-Host "파일 : $outputFile"
    Write-Host "재귀분석 행 : $($rows1.Count)개 / 최종원천요약 행 : $($rows2.Count)개"
}
finally {
    if ($srcWb) { $srcWb.Close($false) }
    if ($outWb) { $outWb.Close($true) }

    foreach ($o in @($sh1, $sh2, $ws, $srcWb, $outWb)) {
        if ($o) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
    }
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
