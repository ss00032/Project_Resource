# Parse-SqlFolder_v2.ps1
# 사용: .\Parse-SqlFolder_v2.ps1 -RootPath "C:\A" -OutFile "C:\temp\result.xlsx"
param(
    [string]$RootPath = 'C:\A',
    [string]$OutFile  = (Join-Path (Get-Location) 'sql_parse_result.xlsx')
)

try { [Text.Encoding]::RegisterProvider([Text.CodePagesEncodingProvider]::Instance) } catch {}

# ---------- 파일 읽기 (UTF-8 시도 후 실패하면 CP949) ----------
function Read-Text($path) {
    $b = [IO.File]::ReadAllBytes($path)
    $utf8 = New-Object Text.UTF8Encoding($false, $true)
    try   { $t = $utf8.GetString($b) }
    catch { $t = [Text.Encoding]::GetEncoding(949).GetString($b) }
    return $t.TrimStart([char]0xFEFF)
}

# ---------- 주석 제거 ----------
function Remove-Comments($s) {
    $s = [regex]::Replace($s, '/\*.*?\*/', ' ', [Text.RegularExpressions.RegexOptions]::Singleline)
    $s = [regex]::Replace($s, '--[^\r\n]*', ' ')
    return $s
}

$id = '[#@\w\.\[\]"]+'
$kw = 'WHERE|ON|SET|INNER|LEFT|RIGHT|FULL|CROSS|OUTER|JOIN|GROUP|ORDER|HAVING|UNION|SELECT|WITH'

# ---------- WITH ... AS (...) 분리 ----------
# "WITH a AS (본문1), b AS (본문2) 본쿼리" ->
#   "@@WITHAS@@ a @@ 본문1", "@@WITHAS@@ b @@ 본문2", "본쿼리" 로 나눠서 반환
function Expand-Cte($raw) {
    $t = ($raw -replace '\s+', ' ').Trim()
    $rx = [regex]'(?i)\bWITH\s+(\w+)\s*(?:\([^)]*\)\s*)?AS\s*\('
    $m = $rx.Match($t)
    if (-not $m.Success) { return $t }

    $res = New-Object System.Collections.ArrayList
    $head = $t.Substring(0, $m.Index)
    $name = $m.Groups[1].Value
    $pos  = $m.Index + $m.Length - 1
    $rxNext = [regex]'^\s*,\s*(\w+)\s*(?:\([^)]*\)\s*)?AS\s*\('
    $rest = ''

    while ($true) {
        $depth = 0; $end = -1
        for ($k = $pos; $k -lt $t.Length; $k++) {
            if     ($t[$k] -eq '(') { $depth++ }
            elseif ($t[$k] -eq ')') { $depth--; if ($depth -eq 0) { $end = $k; break } }
        }
        if ($end -lt 0) { break }
        [void]$res.Add('@@WITHAS@@ ' + $name + ' @@ ' + $t.Substring($pos + 1, $end - $pos - 1))
        $tail = $t.Substring($end + 1)
        $n = $rxNext.Match($tail)
        if ($n.Success) { $name = $n.Groups[1].Value; $pos = $end + $n.Length }
        else            { $rest = $tail; break }
    }

    $main = ($head + ' ' + $rest).Trim()
    if ($main) { [void]$res.Add($main) }
    return $res
}

# ---------- FROM / JOIN 원천 테이블 추출 (콤마 조인 포함) ----------
function Get-Sources($stmt) {
    $list = New-Object System.Collections.ArrayList
    $first = "(?i)\b(?:FROM|JOIN)\s+($id)(?:\s+(?:AS\s+)?(?!(?:$kw)\b)(\w+))?"
    $next  = "(?i)\G\s*,\s*($id)(?:\s+(?:AS\s+)?(?!(?:$kw)\b)(\w+))?"
    $mc = [regex]::Matches($stmt, $first)
    $rx = New-Object System.Text.RegularExpressions.Regex($next)
    for ($i = 0; $i -lt $mc.Count; $i++) {
        $m = $mc.Item($i)
        [void]$list.Add((New-Object psobject -Property @{ Name = $m.Groups[1].Value; Alias = $m.Groups[2].Value }))
        $pos = $m.Index + $m.Length
        while ($true) {
            $n = $rx.Match($stmt, $pos)
            if (-not $n.Success) { break }
            [void]$list.Add((New-Object psobject -Property @{ Name = $n.Groups[1].Value; Alias = $n.Groups[2].Value }))
            $pos = $n.Index + $n.Length
        }
    }
    return $list
}

# ---------- 파일 검색 ----------
$files = Get-ChildItem -Path $RootPath -Recurse -File -Include *.sql, *.txt, *.sas -ErrorAction SilentlyContinue
Write-Host "대상 파일: $($files.Count)개"

$rootFull = (Resolve-Path $RootPath).ProviderPath.TrimEnd('\')
$out = New-Object System.Collections.ArrayList
$cStmt = 0; $cType = 0

foreach ($f in $files) {
    # 경로 분리: C:\A\B\D\file.sql -> [0]=C: [1]=A [2]=B [3]=D
    $p    = $f.FullName.Split('\')
    $dept = if ($p.Count -gt 3) { $p[2] } else { '' }   # 3번째 폴더
    $team = if ($p.Count -gt 4) { $p[3] } else { '' }   # 4번째 폴더

    # 하이퍼링크 주소: RootPath 기준
    $link = Join-Path $RootPath $f.FullName.Substring($rootFull.Length).TrimStart('\')

    $text = Remove-Comments (Read-Text $f.FullName)

    foreach ($raw in ($text.Split(';') | ForEach-Object { Expand-Cte $_ })) {
        $s = ($raw -replace '\s+', ' ').Trim()
        if (-not $s) { continue }
        $cStmt++

        $type = $null; $target = ''
        if     ($s -match '^@@WITHAS@@ (\w+) @@ (.*)$')                     { $type = 'WITH AS';     $target = $Matches[1]; $s = $Matches[2] }
        elseif ($s -match "(?i)\bINSERT\s+(?:INTO\s+)?($id)")               { $type = 'INSERT';      $target = $Matches[1] }
        elseif ($s -match "(?i)\bUPDATE\s+($id)")                           { $type = 'UPDATE';      $target = $Matches[1] }
        elseif ($s -match "(?i)\bCREATE\s+TABLE\s+($id)\s+AS\b")            { $type = 'CREATE AS';   $target = $Matches[1] }
        elseif ($s -match "(?i)\bSELECT\b.*?\bINTO\s+($id)\s+FROM\b")       { $type = 'SELECT INTO'; $target = $Matches[1] }
        elseif ($s -match '(?i)\bSELECT\b' -and $s -notmatch '(?i)\bDELETE\b') { $type = 'SELECT';   $target = '' }
        if (-not $type) { continue }   # DELETE 등은 제외
        $cType++

        $srcs = Get-Sources $s

        # UPDATE a SET ... FROM tableA a ... : 별칭을 실제 테이블로 치환
        if ($type -eq 'UPDATE') {
            $orig = $target
            $hit = $srcs | Where-Object { $_.Alias -ieq $orig } | Select-Object -First 1
            if ($hit) { $target = $hit.Name }
            $srcs = @($srcs | Where-Object { $_.Name -ine $target -and $_.Alias -ine $orig })
        }

        $names = @($srcs | ForEach-Object { $_.Name } | Select-Object -Unique)
        if ($names.Count -eq 0) { $names = @('') }

        foreach ($n in $names) {
            # 스키마/테이블 분리 (마지막 . 기준, 없으면 전부 테이블)
            $nn  = $n -replace '[\[\]"]', ''
            $dot = $nn.LastIndexOf('.')
            if ($dot -ge 0) { $srcSchema = $nn.Substring(0, $dot).TrimEnd('.'); $srcTable = $nn.Substring($dot + 1) }
            else            { $srcSchema = '';                                  $srcTable = $nn }

            [void]$out.Add([pscustomobject]@{
                '부서' = $dept; '팀명' = $team; '파일' = $f.Name; '경로' = $link
                '구분' = $type; '타겟' = $target
                '원천스키마' = $srcSchema; '원천테이블' = $srcTable
            })
        }
    }
}

Write-Host "문장수: $cStmt / 대상 구문: $cType"
Write-Host "추출 행: $($out.Count)"

# ---------- 엑셀 출력 (한 줄씩, 파일 열 하이퍼링크) / 실패 시 CSV ----------
$cols = '부서', '팀명', '파일', '구분', '타겟', '원천스키마', '원천테이블'
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)

    $xl.ScreenUpdating = $false

    # 헤더
    for ($c = 0; $c -lt $cols.Count; $c++) {
        $ws.Cells.Item(1, $c + 1).Value2 = $cols[$c]
    }

    # 데이터
    for ($r = 0; $r -lt $out.Count; $r++) {
        if ($r % 1000 -eq 0) { Write-Host "$r / $($out.Count)" }
        $row = $r + 2
        for ($c = 0; $c -lt $cols.Count; $c++) {
            if ($cols[$c] -eq '파일') {
                [void]$ws.Hyperlinks.Add($ws.Cells.Item($row, $c + 1), $out[$r].'경로', [Type]::Missing, [Type]::Missing, $out[$r].'파일')
            }
            else {
                $ws.Cells.Item($row, $c + 1).Value2 = [string]$out[$r].($cols[$c])
            }
        }
    }

    $xl.ScreenUpdating = $true
    $ws.Rows.Item(1).Font.Bold = $true
    [void]$ws.Columns.AutoFit()
    $wb.SaveAs($OutFile, 51)
    $wb.Close($false); $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Write-Host "저장 완료: $OutFile"
}
catch {
    Write-Host "엑셀 에러: $($_.Exception.Message)"
    try { if ($xl) { $xl.Quit() } } catch {}
    $csv = [IO.Path]::ChangeExtension($OutFile, 'csv')
    $out | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Host "엑셀 실행 실패 -> CSV로 저장: $csv"
}
