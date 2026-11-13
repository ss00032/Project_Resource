#=====================================================================
# 엑셀 기준 재귀 계보 분석 - LEVEL 세로형 (최초타겟 / 원천 / 레벨, CSV 출력)
#
#  [동작]
#   - 모수 시트의 [파일명 / 원천스키마 / 원천테이블 / 타겟스키마·테이블(최대 3쌍)] 을 읽는다
#   - 파일명에 들어 있는 테이블명으로 "그 테이블을 만드는 파일"을 찾는다
#       예) TSP-PIMC_IRA_PMNCUSTINFOM_01_1692.sql  ->  IRA_PMNCUSTINFOM 을 만드는 파일
#   - 최초타겟 -> LEVEL1(그 파일의 원천 전부) -> LEVEL2(LEVEL1 각 테이블을 만드는 파일의 원천 전부) -> ...
#     원천이 더 없을 때까지 내려가고, 끝까지 간 경로 1개를 1행으로 출력한다
#   - 엑셀 쓰기를 쓰지 않고 CSV 로 바로 저장한다 (빠름)
#=====================================================================
$ErrorActionPreference = "Stop"

#---------------------------------------------------------------
# 설정  (경로는 본인 환경에 맞게 수정)
#---------------------------------------------------------------
$inputExcel = "C:\경로\정리된엑셀.xlsx"        # 정리해둔 엑셀 (디스크에 저장된 상태로 읽음)
$sheetName  = "재귀분석_모수"                   # 읽을 시트
$outputXlsx = "C:\경로\재귀분석_세로결과.xlsx"  # 결과 엑셀 (엑셀에서 열어 둔 파일이면 저장이 안 되니 닫고 실행)

$maxLevel        = 30        # 최대 LEVEL 깊이 (안전장치)
$maxPathsPerRoot = 3000      # 최초타겟 1개당 최대 행 수 (경로가 폭발하는 경우 방지)
$maxTotalRows    = 1000000   # 전체 최대 행 수 (엑셀 한 시트 한계 1,048,576)
$showExcel       = $false    # $true 로 두면 읽는 동안 엑셀 창이 보임

# 최초타겟(모수)에서 제외할 타겟스키마. 비워 두면 제외 없음.
#  - 모든 타겟(타겟1~3)의 스키마가 아래 패턴에 해당하는 파일은 최초타겟에서 뺀다 (원천을 찾을 때는 그대로 사용)
#  - 예) @("S_RSK")  /  @("S_RSK*")
$excludeRootSchemas = @("S_RSK")

# 검색(펼치기) 제외 스키마. 이 스키마의 테이블은 LEVEL 에는 나오지만, 그 테이블을 만드는 파일은 더 찾지 않는다 (최종 원천 취급)
#  - 예) @("S_CBK")  /  @("S_CBK","S_ABC")  /  @("S_CBK*")
$excludeSearchSchemas = @("S_CBK")

# 1:1 복사 파일 제외 여부
#  - 원천이 1개뿐이고 그 원천이 타겟과 이름이 같은 테이블(스키마만 다름)인 파일은 분석대상이 아니므로 검색/최초타겟에서 뺀다
#  - 같은 이름의 원천 외에 다른 원천이 하나라도 있으면 분석대상이다
$skipOneToOne = $true

# 최초타겟에서 제외할 문자 (테이블명에 아래 문자가 들어 있으면 최초타겟에서 뺀다. 원천으로 나올 때는 그대로 표시)
$excludeRootChars = @("#")


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

# 스키마.테이블 (스키마 없으면 테이블명만)
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

# 파일명 -> 그 파일이 만드는 테이블명
#  TSP-PIMC_IRA_PMNCUSTINFOM_01_1692.sql -> IRA_PMNCUSTINFOM
#  파일명 뒤쪽 번호 형식(_01_986 , _01_1692 , 가운데 숫자 자리수가 달라도)에 의존하지 않는다.
#  파일명을 "_" 또는 "-" 단위로 나눈 뒤, 연속된 조각의 조합이 엑셀에 나오는 테이블명과 일치하는 것 중
#  가장 긴 것을 선택한다 (길이가 같으면 파일명에서 더 뒤쪽 것). 숫자만으로 된 조합은 제외.
function Get-FileTable {
    param([string]$fileName)
    $stem   = [System.IO.Path]::GetFileNameWithoutExtension($fileName).ToUpper()
    $tokens = @($stem -split '[_\-]' | Where-Object { $_ })

    $best = ""
    $bestLen = 0
    for ($i = 0; $i -lt $tokens.Count; $i++) {
        $cand = ""
        for ($j = $i; $j -lt $tokens.Count; $j++) {
            if ($cand) { $cand = $cand + "_" + $tokens[$j] } else { $cand = $tokens[$j] }
            if ($knownTables.ContainsKey($cand) -and $cand -notmatch '^\d+$' -and $cand.Length -ge $bestLen) {
                $best    = $cand
                $bestLen = $cand.Length
            }
        }
    }
    return $best
}

function Test-ExcludedFile {
    param([string]$f)
    if ($excludeRootSchemas.Count -eq 0) { return $false }
    $schs = @($fileTargetSch[$f] | Where-Object { $_ })
    if ($schs.Count -eq 0) { return $false }
    foreach ($s in $schs) {
        $hit = $false
        foreach ($ex in $excludeRootSchemas) {
            if ($ex -and $s -like $ex) { $hit = $true; break }
        }
        if (-not $hit) { return $false }
    }
    return $true
}

# 노드(스키마.테이블)의 스키마가 검색 제외 스키마인지
function Test-SearchExcluded {
    param([string]$nodeKey)
    if ($excludeSearchSchemas.Count -eq 0) { return $false }
    $i = $nodeKey.LastIndexOf('.')
    if ($i -le 0) { return $false }
    $sch = $nodeKey.Substring(0, $i)
    foreach ($ex in $excludeSearchSchemas) {
        if ($ex -and $sch -like $ex) { return $true }
    }
    return $false
}

# 경로 1개를 결과에 추가 (같은 경로는 한 번만)
function Emit-Path {
    param([string]$Root, [string[]]$Nodes)

    $k = $Root + "|" + ($Nodes -join "|")
    if ($seenPath.ContainsKey($k)) { return }
    $seenPath[$k] = $true

    [void]$results.Add([string[]](@($Root) + $Nodes))

    $state.RootRows++
    if ($Nodes.Count -gt $state.MaxDepth) { $state.MaxDepth = $Nodes.Count }

    if ($state.RootRows -ge $maxPathsPerRoot) { $state.RootStop = $true }
    if ($results.Count -ge $maxTotalRows)     { $state.AllStop  = $true; $state.RootStop = $true }
}

# 재귀: 지금 테이블($Table)을 만드는 파일들의 원천을 LEVEL 하나 더 내려가며 펼친다
function Walk {
    param(
        [string]$Root,
        [string]$Table,
        [string[]]$Nodes,        # 지금까지의 경로 (LEVEL1, LEVEL2 ...)
        [string[]]$UsedFiles,    # 이 경로에서 이미 사용한 파일 (순환 방지)
        [string]$CurKey          # 지금 노드 (스키마.테이블)
    )

    if ($state.RootStop) { return }

    $expanded = $false

    if ($Nodes.Count -lt $maxLevel -and $byTable.ContainsKey($Table) -and -not (Test-SearchExcluded $CurKey)) {
        foreach ($f in $byTable[$Table]) {
            if ($UsedFiles -contains $f) { continue }
            if ($passFiles.ContainsKey($f)) { continue }      # 1:1 복사 파일은 검색 제외

            foreach ($s in $fileSources[$f]) {
                if ($s.Key -eq $CurKey) { continue }      # 자기 자신을 원천으로 쓰는 경우는 건너뜀
                $expanded = $true
                Walk -Root $Root -Table $s.Table -Nodes (@($Nodes) + $s.Key) -UsedFiles (@($UsedFiles) + $f) -CurKey $s.Key
                if ($state.RootStop) { return }
            }
        }
    }

    if (-not $expanded) { Emit-Path -Root $Root -Nodes $Nodes }
}

function Csv-Field {
    param($v)
    $s = [string]$v
    if ($s -match '[",\r\n]') { return '"' + $s.Replace('"', '""') + '"' }
    return $s
}

#---------------------------------------------------------------
# 1. 엑셀 읽기
#---------------------------------------------------------------
$tmp = Join-Path $env:TEMP ("lineage_src_" + [guid]::NewGuid().ToString("N") + ".xlsx")
Copy-Item $inputExcel $tmp -Force

$excel = $null; $srcWb = $null; $ws = $null
$data = $null; $nRows = 0
$cFile = $null; $cSrc = $null; $cSrcSch = $null; $tgtCols = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $showExcel
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false

    Write-Host "엑셀 여는 중 ..."
    $srcWb = $excel.Workbooks.Open($tmp, 0, $true)

    $hdr = $null
    foreach ($s in $srcWb.Worksheets) {
        if ($sheetName -and $s.Name -ne $sheetName) { continue }
        $h = Find-HeaderMap $s
        if ($h) { $ws = $s; $hdr = $h; break }
    }
    if (-not $ws) { throw "헤더(파일명 / 원천테이블 / 타겟테이블)를 가진 시트를 찾지 못했습니다. `$sheetName 을 확인하세요." }
    Write-Host "읽는 시트 : $($ws.Name) (헤더 $($hdr.Row)행)"

    $firstData = $hdr.Row + 1

    $cFile   = $hdr.Map["파일명"]
    $cSrc    = $hdr.Map["원천테이블"]
    $cSrcSch = $hdr.Map["원천스키마"]

    # 타겟은 최대 3쌍 (스키마 / 테이블)
    $tgtCols = @(
        @{ Sch = $hdr.Map["타겟스키마"];  Tbl = $hdr.Map["타겟테이블"]  },
        @{ Sch = $hdr.Map["타겟스키마2"]; Tbl = $hdr.Map["타겟테이블2"] },
        @{ Sch = $hdr.Map["타겟스키마3"]; Tbl = $hdr.Map["타겟테이블3"] }
    )

    $needCols = @($cFile, $cSrc, $cSrcSch) + @($tgtCols | ForEach-Object { $_.Sch; $_.Tbl }) |
                Where-Object { $_ } | Sort-Object -Unique
    $maxCol = ($needCols | Measure-Object -Maximum).Maximum

    # 마지막 행 (UsedRange 기준)
    $used    = $ws.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    Write-Host "데이터 범위 : $firstData ~ $lastRow 행 / 읽는 컬럼 $($needCols.Count)개"
    if ($lastRow -lt $firstData) { throw "데이터 행이 없습니다." }

    # Range(...).Value2 가 null 로 나오는 환경이라 Cells.Item(행,열).Value2 로 한 칸씩 읽는다
    $nRows = $lastRow - $firstData + 1
    $data  = [Array]::CreateInstance([object], [int[]]@($nRows, $maxCol), [int[]]@(1, 1))

    for ($i = 1; $i -le $nRows; $i++) {
        $rr = $firstData + $i - 1
        foreach ($c in $needCols) {
            $data.SetValue($ws.Cells.Item($rr, $c).Value2, $i, $c)
        }
        if ($i % 500 -eq 0) { Write-Host "  읽는 중 ... $i / $nRows" }
    }
    Write-Host "읽기 완료 : $nRows 행"
}
finally {
    if ($srcWb) { $srcWb.Close($false) }
    foreach ($o in @($ws, $srcWb)) {
        if ($o) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
    }
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

#---------------------------------------------------------------
# 2. 파일별 원천 / 알려진 테이블명 정리 (여기부터는 엑셀을 쓰지 않음)
#---------------------------------------------------------------
$fileSources   = @{}     # 파일명 -> 원천 목록 (Key = 스키마.테이블, Table = 테이블명)
$fileTargetSch = @{}     # 파일명 -> 타겟스키마 목록 (모수 제외 판단용)
$fileFirstTgt  = @{}     # 파일명 -> 첫 타겟테이블명 (파일명에서 테이블을 못 찾을 때 대비)
$fileTgts      = @{}     # 파일명 -> 타겟테이블명 목록 (타겟1~3, 중복 제거)
$knownTables   = @{}     # 엑셀에 나오는 모든 테이블명
$seenSrc       = @{}
$fileList      = New-Object System.Collections.ArrayList

for ($r = 1; $r -le $nRows; $r++) {

    $file = Get-Cell $data $r $cFile
    if (-not $file) { continue }

    if (-not $fileSources.ContainsKey($file)) {
        $fileSources[$file]   = New-Object System.Collections.ArrayList
        $fileTargetSch[$file] = New-Object System.Collections.ArrayList
        $fileTgts[$file]      = New-Object System.Collections.ArrayList
        [void]$fileList.Add($file)
    }

    $srcTbl = Get-BaseName (Get-Cell $data $r $cSrc)
    $srcSch = (Get-Cell $data $r $cSrcSch).Trim('"').ToUpper()

    if ($srcTbl) {
        $knownTables[$srcTbl] = $true
        $key = Get-Key $srcSch $srcTbl
        $dk  = "$file|$key"
        if (-not $seenSrc.ContainsKey($dk)) {
            $seenSrc[$dk] = $true
            [void]$fileSources[$file].Add([PSCustomObject]@{ Key = $key; Table = $srcTbl })
        }
    }

    foreach ($tc in $tgtCols) {
        if (-not $tc.Tbl) { continue }
        $tgtRaw = Get-Cell $data $r $tc.Tbl
        if (-not $tgtRaw) { continue }
        $tgtSch = (Get-Cell $data $r $tc.Sch).Trim('"').ToUpper()

        foreach ($t in ($tgtRaw -split ',')) {
            $tb = Get-BaseName $t
            if (-not $tb) { continue }
            $knownTables[$tb] = $true
            [void]$fileTargetSch[$file].Add($tgtSch)
            if (-not $fileFirstTgt.ContainsKey($file)) { $fileFirstTgt[$file] = $tb }
            if (-not $fileTgts[$file].Contains($tb)) { [void]$fileTgts[$file].Add($tb) }
        }
    }
}

# 파일 -> 테이블 / 테이블 -> 파일들
$byTable      = @{}
$fileTableMap = @{}
$passFiles    = @{}      # 1:1 복사 파일 (분석대상 아님)
$noTable      = 0
foreach ($f in $fileList) {
    # 이 파일이 만드는 테이블 = 타겟테이블명(타겟1~3 전부)
    $tlist = @($fileTgts[$f])
    if ($tlist.Count -eq 0) { $noTable++; continue }

    $fileTableMap[$f] = $tlist[0]
    foreach ($t in $tlist) {
        if (-not $byTable.ContainsKey($t)) { $byTable[$t] = New-Object System.Collections.ArrayList }
        if (-not $byTable[$t].Contains($f)) { [void]$byTable[$t].Add($f) }
    }

    # 원천이 1개뿐이고 그 원천이 타겟 중 하나와 같은 테이블명이면 1:1
    if ($skipOneToOne -and $fileSources[$f].Count -eq 1 -and ($tlist -contains $fileSources[$f][0].Table)) {
        $passFiles[$f] = $true
    }
}

# 최초타겟 목록
$roots = New-Object System.Collections.ArrayList
foreach ($t in ($byTable.Keys | Sort-Object)) {
    $skipRoot = $false
    foreach ($ch in $excludeRootChars) { if ($ch -and $t.Contains($ch)) { $skipRoot = $true; break } }
    if ($skipRoot) { continue }          # 테이블명에 # 등이 있으면 최초타겟 제외
    $keep = $false
    foreach ($f in $byTable[$t]) {
        if ($passFiles.ContainsKey($f)) { continue }          # 1:1 복사 파일은 최초타겟에서도 제외
        if (-not (Test-ExcludedFile $f)) { $keep = $true; break }
    }
    if ($keep) { [void]$roots.Add($t) }
}

Write-Host "파일 $($fileList.Count)개 / 테이블 $($byTable.Count)개 / 최초타겟 $($roots.Count)개 / 1:1 제외 파일 $($passFiles.Count)개 / 테이블명을 못 찾은 파일 $noTable 개"

#---------------------------------------------------------------
# 3. 재귀 분석 (세로형)  최초타겟 | 원천테이블 | 레벨 | 상위테이블
#    - 최초타겟마다 위에서부터 한 레벨씩 내려가며, 테이블을 처음 만난 레벨에서 1행만 기록한다
#---------------------------------------------------------------
$results = New-Object System.Collections.ArrayList
$rootCnt = 0

foreach ($t in $roots) {
    $visited = @{}
    $visited[$t] = $true

    $queue = New-Object System.Collections.Queue
    $queue.Enqueue(@($t, $t, 0))          # 노드키, 테이블명, 레벨

    while ($queue.Count -gt 0) {
        $cur    = $queue.Dequeue()
        $curKey = [string]$cur[0]
        $curTbl = [string]$cur[1]
        $curLv  = [int]$cur[2]

        if ($curLv -ge $maxLevel) { continue }
        if (-not $byTable.ContainsKey($curTbl)) { continue }
        if (Test-SearchExcluded $curKey) { continue }

        foreach ($f in $byTable[$curTbl]) {
            if ($passFiles.ContainsKey($f)) { continue }
            foreach ($s in $fileSources[$f]) {
                if ($s.Key -eq $curKey) { continue }
                if ($visited.ContainsKey($s.Key)) { continue }
                $visited[$s.Key] = $true

                $lv = $curLv + 1
                [void]$results.Add([string[]]@($t, $s.Key, "LEVEL$lv", $curKey))
                $queue.Enqueue(@($s.Key, $s.Table, $lv))
            }
        }
    }

    $rootCnt++
    Write-Host "  [$rootCnt / $($roots.Count)] $t (누적 $($results.Count)) $(Get-Date -Format HH:mm:ss)"
}

Write-Host "재귀 분석 결과 : $($results.Count) 행"

#---------------------------------------------------------------
# 4. 엑셀(xlsx) 저장  - 엑셀이 직접 저장하므로 DRM 적용
#---------------------------------------------------------------
$header = @("최초타겟테이블", "원천테이블", "레벨", "상위테이블")
$colCnt = $header.Count

$sheetMax = 1000000      # 시트 1개당 최대 데이터 행 (넘으면 시트를 나눔)
$chunk    = 5000         # 한 번에 쓰는 행 수

if ([string]::IsNullOrWhiteSpace($outputXlsx)) { throw "`$outputXlsx 경로가 비어 있습니다." }
$outDir = [System.IO.Path]::GetDirectoryName($outputXlsx)
if ($outDir -and -not (Test-Path $outDir)) { [void](New-Item -ItemType Directory -Path $outDir -Force) }

$xl = $null; $wb = $null; $sh = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $showExcel
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false

    $wb = $xl.Workbooks.Add()

    $total   = $results.Count
    $sheetNo = 0
    $pos     = 0

    do {
        $sheetNo++
        if ($sheetNo -eq 1) { $sh = $wb.Worksheets.Item(1) }
        else {
            $prev = $sh
            $sh = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $prev)
        }
        $sh.Name = "재귀분석_$sheetNo"

        # 헤더
        for ($c = 1; $c -le $colCnt; $c++) { $sh.Cells.Item(1, $c).Value2 = $header[$c - 1] }

        $sheetEnd = [Math]::Min($pos + $sheetMax, $total)
        $outRow   = 2

        while ($pos -lt $sheetEnd) {
            $n   = [Math]::Min($chunk, $sheetEnd - $pos)
            $arr = New-Object 'object[,]' $n, $colCnt
            for ($i = 0; $i -lt $n; $i++) {
                $row = $results[$pos + $i]
                for ($c = 0; $c -lt $colCnt; $c++) {
                    if ($c -lt $row.Length) { $arr[$i, $c] = [string]$row[$c] } else { $arr[$i, $c] = "" }
                }
            }

            $rng = $sh.Range($sh.Cells.Item($outRow, 1), $sh.Cells.Item($outRow + $n - 1, $colCnt))
            # 1차: COM 직접 호출로 배열 쓰기 / 2차: 일반 대입 / 3차: 한 칸씩 쓰기
            $ok = $false
            try {
                [void]$rng.GetType().InvokeMember("Value2", [System.Reflection.BindingFlags]::SetProperty, $null, $rng, @(,$arr))
                $ok = $true
            } catch { }

            if (-not $ok) {
                try { $rng.Value2 = $arr; $ok = $true } catch { }
            }

            if ($ok) {
                $chk = $sh.Cells.Item($outRow, 1).Value2
                if ($null -eq $chk -or "$chk" -eq "") { $ok = $false }
            }

            if (-not $ok) {
                Write-Host "  배열 쓰기 실패 -> 한 칸씩 쓰기 ($outRow 행부터)"
                for ($i = 0; $i -lt $n; $i++) {
                    for ($c = 0; $c -lt $colCnt; $c++) {
                        $v = $arr[$i, $c]
                        if ($v -ne "") { $sh.Cells.Item($outRow + $i, $c + 1).Value2 = $v }
                    }
                }
            }

            $pos    += $n
            $outRow += $n
            Write-Host "  엑셀 쓰기 ... $pos / $total"
        }

        $hdrRng = $sh.Range($sh.Cells.Item(1, 1), $sh.Cells.Item(1, $colCnt))
        $hdrRng.Font.Bold = $true
        [void]$hdrRng.AutoFilter()
        for ($c = 1; $c -le $colCnt; $c++) { $sh.Columns.Item($c).ColumnWidth = 30 }
    } while ($pos -lt $total)

    if (Test-Path $outputXlsx) { Remove-Item $outputXlsx -Force }
    $wb.SaveAs($outputXlsx, 51)      # 51 = xlsx
    $wb.Close($false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} ; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {} }
    if ($sh) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sh) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {} ; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "================================="
Write-Host "엑셀 저장 완료 : $outputXlsx"
Write-Host "행 수 : $($results.Count) / 시트 $sheetNo 개"

