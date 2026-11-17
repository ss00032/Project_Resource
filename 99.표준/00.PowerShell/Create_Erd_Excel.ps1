# =====================================================================
#  Build-ERD.ps1
#  - '테라스트림 테이블분석' 시트의 타겟테이블(타겟/타겟2/타겟3...)을 기준으로
#    '테이블정의서' 시트의 컬럼 정보를 읽어 ERD 시트를 그립니다.
#    (박스 제목 = 테이블한글명, 박스 안 = 컬럼한글명)
#  - 분석 시트의 원천테이블한글명 / 타겟테이블한글명(1,2,3...) 셀을 누르면
#    ERD 시트의 해당 테이블 박스로 이동합니다.
#  - 엑셀을 화면 없이(백그라운드) 열어 작업하고, 원본은 그대로 둔 채
#    '원본이름_ERD.xlsx' 로 저장(같은 이름이 있으면 덮어씀)한 뒤 엑셀을 종료합니다.
#
#  사용: 아래 설정의 $Path 를 수정한 뒤
#        powershell -ExecutionPolicy Bypass -File .\Build-ERD.ps1
#        (콘솔에 붙여넣어 실행해도 됩니다. 붙여넣기 전에 새 PowerShell 창을 여세요)
# =====================================================================

# ===================== 설정 =====================
$Path          = ""     # 예: "C:\작업\EDW_분석.xlsx"  (비우면 지금 열려 있는 엑셀에서 작업하고 저장은 직접)
$OutPath       = ""     # 비우면 원본폴더\원본이름_ERD.xlsx 로 저장
$DefSheet      = "테이블정의서"
$AnaSheet      = "테라스트림 테이블분석"
$ErdSheet      = "ERD"
$BoxesPerRow   = 6      # 한 줄에 배치할 테이블 박스 수
$BandGap       = 2      # 박스 줄 사이 빈 행 수
$MaxTargetSets = 5      # 타겟테이블 / 타겟테이블2 / ... 최대 몇 세트까지 볼지
# ================================================
$ErrorActionPreference = "Stop"

function Norm($v) {
    if ($null -eq $v) { return "" }
    return (([string]$v) -replace "[\s\u00A0]", "").ToUpper()
}
function Cell($d, [int]$i, [int]$c) {
    if ($c -le 0) { return "" }
    $x = $d[$i, $c]
    if ($null -eq $x) { return "" }
    return ([string]$x).Trim()
}
function ColLetter([int]$n) {
    $s = ""
    while ($n -gt 0) {
        $m = ($n - 1) % 26
        $s = [string][char](65 + $m) + $s
        $n = [int][math]::Floor(($n - 1) / 26)
    }
    return $s
}
function RGBc([int]$r, [int]$g, [int]$b) { return $r + 256 * $g + 65536 * $b }

# 헤더 문구가 있는 행(1~30행)을 찾아 행 번호 반환 (없으면 0)
function Find-HeaderRow($ws, [string]$anchor) {
    $ur = $ws.UsedRange
    $lastCol = $ur.Column + $ur.Columns.Count - 1
    $blk = $ws.Range($ws.Cells(1, 1), $ws.Cells(30, $lastCol)).Value2
    $k = Norm $anchor
    for ($r = 1; $r -le 30; $r++) {
        for ($c = 1; $c -le $lastCol; $c++) {
            if ((Norm ($blk[$r, $c])) -eq $k) { return $r }
        }
    }
    return 0
}
# 헤더 행에서 정확히 일치하는 열 번호 (없으면 0)
function ColOf($ws, [int]$row, [string]$title) {
    $lastC = $ws.Cells($row, $ws.Columns.Count).End(-4159).Column
    $k = Norm $title
    for ($c = 1; $c -le $lastC; $c++) {
        if ((Norm ($ws.Cells($row, $c).Value2)) -eq $k) { return $c }
    }
    return 0
}
# '<prefix>한글명' 정확히 일치가 없으면 prefix 로 시작하고 '한글'이 들어간 헤더를 사용
function Find-KoColRow($ws, [int]$row, [string]$prefix) {
    $exact = ColOf $ws $row ($prefix + "한글명")
    if ($exact -gt 0) { return $exact }
    $p = Norm $prefix
    $lastC = $ws.Cells($row, $ws.Columns.Count).End(-4159).Column
    for ($c = 1; $c -le $lastC; $c++) {
        $n = Norm ($ws.Cells($row, $c).Value2)
        if ($n.StartsWith($p) -and $n -ne $p -and $n.Contains("한글")) { return $c }
    }
    return 0
}

# ===================== 실행 =====================
$excel = $null
$wb = $null
$own = $false
$prevCalc = $null
try {
    # ---------- 엑셀 연결 ----------
    if ($Path) {
        $full = (Resolve-Path -LiteralPath $Path).Path
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($full, 0, $true)    # 읽기 전용으로 열기 (원본 보호)
        $own = $true
    } else {
        try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") }
        catch { throw "실행 중인 Excel을 찾지 못했습니다. 설정의 `$Path 를 지정하거나 파일을 먼저 여세요. (PowerShell과 Excel은 같은 권한으로 실행해야 합니다)" }
        $wb = $excel.ActiveWorkbook
    }
    $prevCalc = $excel.Calculation

    $wsD = $null
    $wsA = $null
    try { $wsD = $wb.Worksheets.Item($DefSheet) } catch {}
    try { $wsA = $wb.Worksheets.Item($AnaSheet) } catch {}
    if (-not $wsD -or -not $wsA) {
        $names = @()
        foreach ($sh in $wb.Worksheets) { $names += $sh.Name }
        throw "시트를 찾지 못했습니다 ($DefSheet / $AnaSheet). 현재 시트: $($names -join ', ')"
    }

    # ---------- 1. 테이블정의서 읽기 ----------
    $hdrD = Find-HeaderRow $wsD "테이블명"
    if ($hdrD -eq 0) { throw "$DefSheet 에서 '테이블명' 헤더를 찾지 못했습니다." }
    $cT  = ColOf $wsD $hdrD "테이블명"
    $cTK = ColOf $wsD $hdrD "테이블한글명"
    $cC  = ColOf $wsD $hdrD "컬럼명"
    $cCK = ColOf $wsD $hdrD "컬럼한글명"
    if ($cC -eq 0) { throw "$DefSheet 에서 '컬럼명' 헤더를 찾지 못했습니다." }
    $maxD = (@($cT, $cTK, $cC, $cCK) | Measure-Object -Maximum).Maximum
    $lastD = $wsD.Cells($wsD.Rows.Count, $cT).End(-4162).Row
    if ($lastD -le $hdrD) { throw "$DefSheet 에 데이터가 없습니다." }
    $dataD = $wsD.Range($wsD.Cells(1, 1), $wsD.Cells($lastD, $maxD)).Value2

    $defCols = @{}
    $defKo = @{}
    for ($i = $hdrD + 1; $i -le $lastD; $i++) {
        $t = Norm ($dataD[$i, $cT])
        if (-not $t) { continue }
        $cn = Cell $dataD $i $cC
        $ck = Cell $dataD $i $cCK
        $lab = $cn
        if ($ck) { $lab = $ck }
        if (-not $lab) { continue }
        if (-not $defCols.ContainsKey($t)) {
            $defCols[$t] = New-Object System.Collections.Generic.List[string]
            $defKo[$t] = ""
        }
        if (-not $defKo[$t]) { $defKo[$t] = Cell $dataD $i $cTK }
        $defCols[$t].Add($lab)
    }

    # ---------- 2. 테라스트림 테이블분석: 열 찾기 ----------
    $hdrA = Find-HeaderRow $wsA "타겟테이블"
    if ($hdrA -eq 0) { throw "$AnaSheet 에서 '타겟테이블' 헤더를 찾지 못했습니다." }
    $aTgt = ColOf $wsA $hdrA "타겟테이블"
    $aSrc = ColOf $wsA $hdrA "원천테이블"
    $aSrcKo = ColOf $wsA $hdrA "원천테이블한글명"
    if ($aSrcKo -eq 0) { $aSrcKo = Find-KoColRow $wsA $hdrA "원천테이블" }

    # 타겟 세트: (타겟테이블한글명, 타겟테이블) / (타겟테이블한글명2, 타겟테이블2) / ...
    $tgtPairs = @()
    foreach ($suf in (@("") + @(2..$MaxTargetSets))) {
        $kc   = ColOf $wsA $hdrA ("타겟테이블한글명" + $suf)
        $tcol = ColOf $wsA $hdrA ("타겟테이블" + $suf)
        if ($kc -gt 0 -and $tcol -gt 0) {
            $tgtPairs += ,@{ Ko = $kc; Tbl = $tcol }
            Write-Host ("타겟 매핑: " + (ColLetter $kc) + "열 <-> " + (ColLetter $tcol) + "열")
        }
    }

    # 링크를 걸 열 목록
    $pairs = @()
    if ($aSrcKo -gt 0 -and $aSrc -gt 0) {
        $pairs += ,@{ Label = "원천테이블한글명"; NameCol = $aSrcKo; TblCol = $aSrc }
    } else {
        Write-Warning "원천테이블/원천테이블한글명 열을 찾지 못해 원천 쪽 링크는 건너뜁니다."
    }
    if ($tgtPairs.Count -gt 0) {
        foreach ($tp in $tgtPairs) {
            $pairs += ,@{ Label = "타겟테이블한글명 " + (ColLetter $tp.Ko) + "열"; NameCol = $tp.Ko; TblCol = $tp.Tbl }
        }
    } else {
        Write-Warning "타겟테이블한글명 열을 찾지 못해 타겟 쪽 링크는 건너뜁니다."
    }

    $maxA = (@($aTgt, $aSrc, $aSrcKo) + @($pairs | ForEach-Object { $_.NameCol; $_.TblCol }) | Measure-Object -Maximum).Maximum
    $lastA = 0
    foreach ($col2 in (@($aTgt, $aSrc) + @($tgtPairs | ForEach-Object { $_.Tbl }))) {
        if ($col2 -le 0) { continue }
        $lr = $wsA.Cells($wsA.Rows.Count, $col2).End(-4162).Row
        if ($lr -gt $lastA) { $lastA = $lr }
    }
    if ($lastA -le $hdrA) { throw "$AnaSheet 에 데이터가 없습니다." }
    $dataA = $wsA.Range($wsA.Cells(1, 1), $wsA.Cells($lastA, $maxA)).Value2

    # ERD로 그릴 타겟 테이블 수집 (타겟테이블, 타겟테이블2, ... 모두)
    $erdTables = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $anaKo = @{}
    $noDef = New-Object System.Collections.Generic.List[string]
    $tblColList = @($aTgt) + @($tgtPairs | ForEach-Object { $_.Tbl }) | Select-Object -Unique
    for ($i = $hdrA + 1; $i -le $lastA; $i++) {
        foreach ($tp in $tgtPairs) {
            $t = Norm ($dataA[$i, $tp.Tbl])
            if ($t -and -not $anaKo.ContainsKey($t)) { $anaKo[$t] = Cell $dataA $i $tp.Ko }
        }
        foreach ($tcol in $tblColList) {
            $t = Norm ($dataA[$i, $tcol])
            if (-not $t) { continue }
            if ($seen.ContainsKey($t)) { continue }
            $seen[$t] = 1
            if ($defCols.ContainsKey($t)) { $erdTables.Add($t) }
            else { $noDef.Add($t) }
        }
    }
    if ($erdTables.Count -eq 0) { throw "ERD로 그릴 테이블이 없습니다. (타겟테이블이 테이블정의서에 없음)" }

    # ---------- 3. ERD 시트 생성 ----------
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.Calculation = -4135   # 수동
    $excel.DisplayAlerts = $false
    try { $wb.Worksheets.Item($ErdSheet).Delete() } catch {}

    $lastSheet = $wb.Worksheets.Item($wb.Worksheets.Count)
    $wsE = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet)
    $wsE.Name = $ErdSheet
    $wsE.Cells.Font.Size = 10

    $titleFill = RGBc 221 235 247
    $pos = @{}
    $bandTop = 3
    $curMax = 0
    $idx = 0
    foreach ($t in $erdTables) {
        $bc = $idx % $BoxesPerRow
        if ($bc -eq 0 -and $idx -gt 0) { $bandTop += $curMax + $BandGap; $curMax = 0 }
        $lc = 2 + $bc * 2
        $cols = $defCols[$t]
        $n = $cols.Count
        $title = $defKo[$t]
        if (-not $title -and $anaKo.ContainsKey($t)) { $title = $anaKo[$t] }
        if (-not $title) { $title = $t }

        $rng = $wsE.Range($wsE.Cells($bandTop, $lc), $wsE.Cells($bandTop + $n, $lc))
        $rng.NumberFormat = "@"
        $wsE.Cells($bandTop, $lc).Value2 = $title
        for ($j = 0; $j -lt $n; $j++) {
            $wsE.Cells($bandTop + 1 + $j, $lc).Value2 = $cols[$j]
        }
        $rng.BorderAround(1, 2) | Out-Null
        $tc = $wsE.Cells($bandTop, $lc)
        $tc.Font.Bold = $true
        $tc.Interior.Color = $titleFill
        $tc.BorderAround(1, 2) | Out-Null

        $pos[$t] = "'" + $ErdSheet + "'!" + (ColLetter $lc) + $bandTop
        if ($n + 1 -gt $curMax) { $curMax = $n + 1 }
        $idx++
    }
    $wsE.Columns.Item(1).ColumnWidth = 3
    for ($b = 0; $b -lt $BoxesPerRow; $b++) {
        $wsE.Columns.Item(2 + $b * 2).ColumnWidth = 34
        $wsE.Columns.Item(3 + $b * 2).ColumnWidth = 4
    }
    $null = $wsE.Hyperlinks.Add($wsE.Cells(1, 1), "", "'" + $AnaSheet + "'!A1", "분석 시트로 돌아가기", "◀ $AnaSheet")

    # ---------- 4. 분석 시트의 테이블한글명 -> ERD 하이퍼링크 ----------
    $linkCount = [ordered]@{}
    foreach ($pr in $pairs) {
        $nc = $pr.NameCol
        $tcI = $pr.TblCol
        $rngAll = $wsA.Range($wsA.Cells($hdrA + 1, $nc), $wsA.Cells($lastA, $nc))
        $fn = $wsA.Cells($hdrA + 1, $nc).Font.Name
        $fs = $wsA.Cells($hdrA + 1, $nc).Font.Size
        try { $rngAll.Hyperlinks.Delete() } catch {}
        $cnt = 0
        for ($i = $hdrA + 1; $i -le $lastA; $i++) {
            $t = Norm ($dataA[$i, $tcI])
            if (-not $t -or -not $pos.ContainsKey($t)) { continue }
            if (-not (Cell $dataA $i $nc)) { continue }
            $null = $wsA.Hyperlinks.Add($wsA.Cells($i, $nc), "", $pos[$t], "ERD로 이동: $t")
            $cnt++
        }
        $rngAll.Font.Name = $fn
        $rngAll.Font.Size = $fs
        $linkCount[$pr.Label] = $cnt
    }

    if (-not $own) {
        $wsE.Activate()
        $excel.ActiveWindow.DisplayGridlines = $false
        $excel.ActiveWindow.Zoom = 80
        $wsE.Range("A1").Select() | Out-Null
    }

    Write-Host ""
    Write-Host "완료: ERD 테이블 $($erdTables.Count)개"
    foreach ($k in $linkCount.Keys) { Write-Host "  $k 링크: $($linkCount[$k])개" }
    if ($noDef.Count -gt 0) {
        Write-Warning "타겟테이블이 테이블정의서에 없어 ERD에서 제외된 테이블 $($noDef.Count)개:"
        $noDef | Select-Object -First 30 | ForEach-Object { Write-Host "   $_" }
        if ($noDef.Count -gt 30) { Write-Host "   ... (이하 생략)" }
    }

    # ---------- 5. 저장 ----------
    if ($own) {
        if (-not $OutPath) {
            $dir = Split-Path $full -Parent
            $baseName = [IO.Path]::GetFileNameWithoutExtension($full)
            $ext = [IO.Path]::GetExtension($full)
            $OutPath = Join-Path $dir ($baseName + "_ERD" + $ext)
        }
        if ($OutPath -eq $full) { throw "저장 경로가 원본과 같습니다. 원본을 덮어쓰지 않도록 중단합니다. (`$OutPath 를 다른 이름으로 지정하세요)" }
        if (Test-Path -LiteralPath $OutPath) { Remove-Item -LiteralPath $OutPath -Force }
        $excel.DisplayAlerts = $false
        $wb.SaveAs($OutPath, $wb.FileFormat)
        Write-Host "저장 완료: $OutPath"
    } else {
        Write-Host "엑셀에서 파일을 저장하세요."
    }
}
finally {
    if ($excel) {
        try { $excel.DisplayAlerts = $true } catch {}
        try { $excel.EnableEvents = $true } catch {}
        try { $excel.ScreenUpdating = $true } catch {}
        try { if ($null -ne $prevCalc) { $excel.Calculation = $prevCalc } } catch {}
        if ($own) {
            try { if ($wb) { $wb.Close($false) } } catch {}
            try { $excel.Quit() } catch {}
            try { if ($wb) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } } catch {}
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}
