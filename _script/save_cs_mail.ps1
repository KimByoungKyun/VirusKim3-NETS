# 고객지원 하위 폴더 메일 → .msg 저장 스크립트 (Windows PowerShell 5.1 / PowerShell 7 호환)
# 실행: powershell -ExecutionPolicy Bypass -File "C:\NETS.co.kr\MA_유지보수\_고객지원본부\고객사별_메일\_script\save_cs_mail.ps1"
# 요구: Outlook 데스크톱이 설치되어 있고 로그인된 상태여야 함

param(
    [string]$BaseDir    = 'C:\NETS.co.kr\MA_유지보수\_고객지원본부\고객사별_메일',
    [string]$FolderName = '고객지원',
    [int]$FirstRunYears = 2,
    [switch]$WhatIf
)

$ErrorActionPreference = 'Continue'
$ts         = Get-Date -Format 'yyyyMMdd_HHmmss'
$listPath   = Join-Path $BaseDir '_저장목록.csv'
$resultPath = Join-Path $BaseDir ("_실행결과_$ts.csv")
$archiveDir = Join-Path $BaseDir ("_정리보관\$ts")
$olFolderInbox = 6
$olMSGUnicode  = 3
$olMail        = 43
$cols = @('고객사폴더','스레드ID','수신일시','발신자','제목','파일명','EntryID','상태')

function Clean-Name([string]$s) {
    $s = [regex]::Replace($s, '[\\/:*?"<>|\r\n\t]', '_')
    $s = $s.Trim().TrimEnd('.')
    if ($s.Length -gt 120) { $s = $s.Substring(0, 120) }
    return $s
}

if (-not (Test-Path $BaseDir)) { New-Item -ItemType Directory -Force -Path $BaseDir | Out-Null }
$scriptDir = Join-Path $BaseDir '_script'
New-Item -ItemType Directory -Force -Path $scriptDir | Out-Null

Write-Host '[1/6] Outlook 연결 중...'
$outlook = New-Object -ComObject Outlook.Application
$ns      = $outlook.GetNamespace('MAPI')
$inbox   = $ns.GetDefaultFolder($olFolderInbox)
$csRoot  = $null
foreach ($f in $inbox.Folders) { if ($f.Name -eq $FolderName) { $csRoot = $f } }
if (-not $csRoot) { Write-Host "받은 편지함에서 '$FolderName' 폴더를 찾지 못했습니다. 중단합니다."; exit 1 }

Write-Host '[2/6] 기존 저장목록 읽는 중...'
$records  = @()
$existing = @{}
$maxRecv  = [datetime]::MinValue
$isFirst  = -not (Test-Path $listPath)
if (-not $isFirst) {
    $records = @(Import-Csv -Encoding UTF8 $listPath)
    foreach ($r in $records) {
        $existing[$r.EntryID] = $r
        $d = [datetime]$r.수신일시
        if ($d -gt $maxRecv) { $maxRecv = $d }
    }
    Write-Host "  기존 기록 $($records.Count)건, 마지막 수신 $($maxRecv.ToString('yyyy-MM-dd HH:mm'))"
} else {
    Write-Host "  _저장목록.csv 없음 → 최근 $FirstRunYears년치 대상"
}

Write-Host '[3/6] 새 메일 수집 중...'
$since = if ($isFirst) { (Get-Date).AddYears(-$FirstRunYears) } else { $maxRecv }
$sinceStr = $since.ToString('MM/dd/yyyy HH:mm')
$candidates = New-Object System.Collections.ArrayList
$failCount = 0
foreach ($sub in $csRoot.Folders) {
    $items = $sub.Items.Restrict("[ReceivedTime] >= '$sinceStr'")
    foreach ($it in $items) {
        try {
            if ($it.Class -ne $olMail) { continue }
            [void]$candidates.Add([pscustomobject]@{
                Folder  = $sub.Name
                Conv    = [string]$it.ConversationID
                Recv    = [datetime]$it.ReceivedTime
                Subject = [string]$it.Subject
                Sender  = [string]$it.SenderName
                EntryID = [string]$it.EntryID
                Item    = $it
            })
        } catch { $failCount++ }
    }
}
Write-Host "  수집 $($candidates.Count)건"

# 같은 고객사 폴더 + 같은 스레드(ConversationID)는 가장 최근 메일 하나만
$latest = @($candidates | Group-Object Folder, Conv | ForEach-Object {
    $_.Group | Sort-Object Recv -Descending | Select-Object -First 1
} | Where-Object { -not $existing.ContainsKey($_.EntryID) })
Write-Host "  스레드 대표 신규 $($latest.Count)건"

Write-Host '[4/6] .msg 저장 중...'
$newRows = @()
$resultRows = @()
$saved = 0
foreach ($c in $latest) {
    $custName = Clean-Name $c.Folder
    $custDir  = Join-Path $BaseDir $custName
    New-Item -ItemType Directory -Force -Path $custDir | Out-Null
    $date = $c.Recv.ToString('yyyyMMdd')
    $name = (Clean-Name ("${date}_" + $c.Subject)) + '.msg'
    $path = Join-Path $custDir $name
    if (Test-Path $path) {
        $name = (Clean-Name ("${date}_" + $c.Subject + '_' + $c.Recv.ToString('HHmmss'))) + '.msg'
        $path = Join-Path $custDir $name
    }
    $status = 'WhatIf'
    if (-not $WhatIf) {
        try {
            $c.Item.SaveAs($path, $olMSGUnicode)
            $status = '저장'; $saved++
        } catch {
            $status = "실패: $($_.Exception.Message)"; $failCount++
        }
    }
    Write-Host "  [$status] $custName\$name"
    $row = [pscustomobject]@{
        '고객사폴더' = $custName; '스레드ID' = $c.Conv
        '수신일시' = $c.Recv.ToString('yyyy-MM-dd HH:mm:ss')
        '발신자' = $c.Sender; '제목' = $c.Subject
        '파일명' = $name; 'EntryID' = $c.EntryID; '상태' = $status
    }
    $resultRows += [pscustomobject]@{ '구분' = '신규'; '고객사폴더' = $custName; '파일명' = $name; '결과' = $status }
    if ($status -eq '저장') { $newRows += $row }
}

Write-Host '[5/6] 정리 대상 확인 중...'
# (a) 같은 스레드의 예전 버전 파일: 새 대표 메일로 교체된 경우
$superseded = @()
foreach ($c in $latest) {
    foreach ($r in $records) {
        if ($r.고객사폴더 -eq (Clean-Name $c.Folder) -and $r.스레드ID -eq $c.Conv -and $r.상태 -eq '저장') {
            $superseded += $r
        }
    }
}
# (b) Outlook에서 옮겨지거나 삭제된 메일의 예전 파일
$orphans = @()
foreach ($r in $records) {
    if ($r.상태 -ne '저장') { continue }
    $alive = $false
    try {
        $it = $ns.GetItemFromID($r.EntryID)
        if ($it.Parent.FolderPath -like "*\$FolderName\*") { $alive = $true }
    } catch { $alive = $false }
    if (-not $alive) { $orphans += $r }
}
$cleanup = @($superseded) + @($orphans) | Sort-Object 파일명 -Unique

$moved = 0
if ($cleanup.Count -gt 0) {
    Write-Host ''
    Write-Host "정리 대상 $($cleanup.Count)건 (삭제하지 않고 _정리보관 폴더로 옮깁니다):"
    $cleanup | ForEach-Object { Write-Host "  - $($_.고객사폴더)\$($_.파일명)  [$($_.제목)]" }
    $ans = Read-Host "위 파일을 _정리보관\$ts 로 옮길까요? (Y/N, 기본 N)"
    if ($ans -match '^[Yy]$') {
        New-Item -ItemType Directory -Force -Path $archiveDir | Out-Null
        foreach ($r in $cleanup) {
            $src = Join-Path (Join-Path $BaseDir $r.고객사폴더) $r.파일명
            if (Test-Path $src) {
                try {
                    Move-Item -LiteralPath $src -Destination (Join-Path $archiveDir $r.파일명) -Force
                    $moved++
                    $r.상태 = '보관이동'
                    $resultRows += [pscustomobject]@{ '구분' = '정리'; '고객사폴더' = $r.고객사폴더; '파일명' = $r.파일명; '결과' = '_정리보관으로 이동' }
                } catch { $failCount++ }
            } else {
                $r.상태 = '파일없음'
            }
        }
    } else {
        Write-Host '  정리는 건너뜁니다.'
    }
}

Write-Host '[6/6] 저장목록 및 실행결과 기록 중...'
$keep = @($records | Where-Object { $_.상태 -ne '보관이동' -and $_.상태 -ne '파일없음' })
$allRows = @($keep) + @($newRows)
$allRows | Select-Object $cols | Export-Csv -LiteralPath $listPath -NoTypeInformation -Encoding UTF8
if ($resultRows.Count -gt 0) {
    $resultRows | Export-Csv -LiteralPath $resultPath -NoTypeInformation -Encoding UTF8
}

$totalFiles = @(Get-ChildItem -LiteralPath $BaseDir -Recurse -Filter *.msg -File | Where-Object { $_.FullName -notlike '*_정리보관*' }).Count

Write-Host ''
Write-Host '=== 완료 ==='
Write-Host "새로 저장한 파일 수 : $saved"
Write-Host "정리한 파일 수      : $moved"
Write-Host "실패 건수           : $failCount"
Write-Host "전체 파일 수        : $totalFiles"
Write-Host "저장목록            : $listPath"
Write-Host "실행결과            : $resultPath"
