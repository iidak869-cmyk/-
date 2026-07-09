<#
.SYNOPSIS
    Dシステムのスケジュールページ (Schedule.aspx) から CS課メンバーの予定を取得し、
    Slack の Incoming Webhook に投稿します。

.DESCRIPTION
    社内ネットワーク上の Windows PC / サーバーで実行してください
    (win2012sv-aws は社外・クラウドからは到達できません)。

    使い方の流れ:
      1) まず -DumpOnly でページのHTMLを保存し、表示内容を確認する
         .\Export-DSystemSchedule.ps1 -DumpOnly
      2) -NoPost で解析結果をコンソールに表示して確認する
         .\Export-DSystemSchedule.ps1 -NoPost
      3) 問題なければ Webhook URL を指定して実際に投稿する
         .\Export-DSystemSchedule.ps1 -WebhookUrl 'https://hooks.slack.com/services/XXX/YYY/ZZZ'

    スケジュール表の解析は「ヘッダー行に日付、先頭列にメンバー名が並ぶ表」という
    一般的なグループウェアの週表示レイアウトを想定した汎用ロジックです。
    Dシステムの実際のHTML構造と合わない場合は解析に失敗し、HTMLダンプを保存して
    終了するので、そのダンプファイルを元にパーサーを調整してください。

.NOTES
    PowerShell 3.0 以降 (Windows Server 2012 標準) で動作するよう記述しています。
    ファイルは UTF-8 (BOM付き) で保存してください。
#>
[CmdletBinding()]
param(
    # DシステムのスケジュールページURL (グループ指定のクエリが必要なら付けてください)
    [string]$ScheduleUrl = 'http://win2012sv-aws/DSystem/Schedule.aspx',

    # 抜き出したいグループ名
    [string]$GroupName = 'CS課',

    # Slack Incoming Webhook URL (未指定なら環境変数 SLACK_WEBHOOK_URL を使用)
    [string]$WebhookUrl = $env:SLACK_WEBHOOK_URL,

    # ページのHTMLを保存するだけで終了する (構造確認用)
    [switch]$DumpOnly,

    # HTMLダンプの保存先 (未指定ならスクリプトと同じフォルダに Schedule_dump.html)
    [string]$DumpPath = '',

    # Slackに投稿せず、整形結果をコンソールに表示する (動作確認用)
    [switch]$NoPost,

    # 表示されている全日付分を投稿する (既定は今日の分のみ)
    [switch]$AllDays
)

$ErrorActionPreference = 'Stop'

# ダンプ保存先が未指定なら、スクリプトのあるフォルダ (取れなければ現在のフォルダ) に保存
if (-not $DumpPath) {
    $baseDir = $PSScriptRoot
    if (-not $baseDir -and $PSCommandPath) { $baseDir = Split-Path -Parent $PSCommandPath }
    if (-not $baseDir) { $baseDir = (Get-Location).Path }
    $DumpPath = Join-Path $baseDir 'Schedule_dump.html'
}

# 古い .NET 既定 (TLS1.0) のままだと Slack への HTTPS 接続に失敗するため TLS1.2 を有効化
try {
    [Net.ServicePointManager]::SecurityProtocol = `
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }

function Get-PageHtml {
    param([string]$Url)
    # 社内のASP.NETサイトは Windows統合認証 が多いため、実行ユーザーの資格情報を使う
    $res = Invoke-WebRequest -Uri $Url -UseDefaultCredentials -UseBasicParsing

    $bytes = $res.RawContentStream.ToArray()

    # 文字コードを Content-Type ヘッダー → HTML内 meta → UTF-8 の順で判定
    $charset = $null
    $ct = $res.Headers['Content-Type']
    if ($ct -and $ct -match 'charset=([\w\-]+)') { $charset = $Matches[1] }
    if (-not $charset) {
        $ascii = [Text.Encoding]::ASCII.GetString($bytes)
        if ($ascii -match 'charset\s*=\s*["'']?([\w\-]+)') { $charset = $Matches[1] }
    }
    $enc = [Text.Encoding]::UTF8
    if ($charset) {
        try { $enc = [Text.Encoding]::GetEncoding($charset) } catch { }
    }
    return $enc.GetString($bytes)
}

function ConvertFrom-HtmlText {
    param([string]$Html)
    $t = $Html -replace '(?i)<br\s*/?>', "`n"
    $t = $t -replace '(?s)<[^>]+>', ''
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    $lines = $t -split "`n" | ForEach-Object { ($_ -replace '[\s ]+', ' ').Trim() } | Where-Object { $_ }
    return ($lines -join "`n")
}

function Get-TableMatches { param([string]$Html)
    [regex]::Matches($Html, '(?is)<table[^>]*>.*?</table>') | ForEach-Object { $_.Value }
}
function Get-RowMatches { param([string]$TableHtml)
    [regex]::Matches($TableHtml, '(?is)<tr[^>]*>.*?</tr>') | ForEach-Object { $_.Value }
}
function Get-CellTexts { param([string]$RowHtml)
    [regex]::Matches($RowHtml, '(?is)<t[dh][^>]*>(.*?)</t[dh]>') |
        ForEach-Object { ConvertFrom-HtmlText $_.Groups[1].Value }
}

# 「7/9」「7月9日」「07/09(木)」などの日付表記にマッチ
$script:DatePattern = '(?<m>\d{1,2})\s*[/月]\s*(?<d>\d{1,2})'

function Find-ScheduleTable {
    <#
      ヘッダー行 (最初の数行のいずれか) に日付らしきセルが2つ以上並ぶ表を
      スケジュール表とみなし、@{ Dates = ...; Members = ... } を返す。
    #>
    param([string]$Html)

    foreach ($table in (Get-TableMatches $Html)) {
        $rows = @(Get-RowMatches $table)
        if ($rows.Count -lt 2) { continue }

        $headerIndex = -1
        $dates = @()
        for ($i = 0; $i -lt [Math]::Min(3, $rows.Count); $i++) {
            $cells = @(Get-CellTexts $rows[$i])
            $dateCells = @($cells | Where-Object { $_ -match $script:DatePattern })
            if ($dateCells.Count -ge 2) {
                $headerIndex = $i
                $dates = $cells
                break
            }
        }
        if ($headerIndex -lt 0) { continue }

        # ヘッダー内で日付が始まる列位置 (それより前は名前などの固定列)
        $firstDateCol = -1
        for ($c = 0; $c -lt $dates.Count; $c++) {
            if ($dates[$c] -match $script:DatePattern) { $firstDateCol = $c; break }
        }

        $members = @()
        for ($r = $headerIndex + 1; $r -lt $rows.Count; $r++) {
            $cells = @(Get-CellTexts $rows[$r])
            if ($cells.Count -le $firstDateCol) { continue }
            $name = ($cells[0..([Math]::Max(0, $firstDateCol - 1))] -join ' ').Trim()
            if (-not $name) { continue }
            $members += ,@{
                Name    = $name
                Entries = @($cells[$firstDateCol..($cells.Count - 1)])
            }
        }
        if ($members.Count -eq 0) { continue }

        return @{
            DateHeaders  = @($dates[$firstDateCol..($dates.Count - 1)])
            Members      = $members
        }
    }
    return $null
}

function Build-SlackText {
    param($Schedule, [datetime]$Today, [bool]$IncludeAllDays, [string]$GroupName)

    $headers = @($Schedule.DateHeaders)

    # 投稿対象の列を決める (既定: 今日の日付列。見つからなければ全列)
    $targetCols = @()
    if (-not $IncludeAllDays) {
        for ($c = 0; $c -lt $headers.Count; $c++) {
            if ($headers[$c] -match $script:DatePattern) {
                if ([int]$Matches['m'] -eq $Today.Month -and [int]$Matches['d'] -eq $Today.Day) {
                    $targetCols += $c
                }
            }
        }
    }
    if ($targetCols.Count -eq 0) { $targetCols = @(0..($headers.Count - 1)) }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(("*【{0} スケジュール】{1}*" -f $GroupName, $Today.ToString('yyyy/MM/dd (ddd)')))

    foreach ($col in $targetCols) {
        if ($targetCols.Count -gt 1) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine(("*■ {0}*" -f $headers[$col]))
        }
        foreach ($m in $Schedule.Members) {
            $entry = ''
            if ($col -lt $m.Entries.Count) { $entry = $m.Entries[$col] }
            [void]$sb.AppendLine(("*{0}*" -f $m.Name))
            if ($entry) {
                foreach ($line in ($entry -split "`n")) {
                    [void]$sb.AppendLine(("  ・{0}" -f $line))
                }
            } else {
                [void]$sb.AppendLine('  ・(予定なし)')
            }
        }
    }
    return $sb.ToString().TrimEnd()
}

function Send-SlackMessage {
    param([string]$Text, [string]$Url)
    # Slack側の文字数制限を考慮し、長い場合は空行区切りで分割投稿する
    $chunks = New-Object System.Collections.ArrayList
    $current = ''
    foreach ($block in ($Text -split "`n`n")) {
        if ($current -and ($current.Length + $block.Length + 2) -gt 3500) {
            [void]$chunks.Add($current); $current = $block
        } elseif ($current) {
            $current = $current + "`n`n" + $block
        } else {
            $current = $block
        }
    }
    if ($current) { [void]$chunks.Add($current) }

    foreach ($chunk in $chunks) {
        $payload = ConvertTo-Json @{ text = $chunk } -Depth 3
        $body = [Text.Encoding]::UTF8.GetBytes($payload)
        Invoke-RestMethod -Uri $Url -Method Post -ContentType 'application/json; charset=utf-8' -Body $body | Out-Null
    }
}

# ---- メイン処理 -------------------------------------------------------------

Write-Host ("ページ取得中: {0}" -f $ScheduleUrl)
$html = Get-PageHtml -Url $ScheduleUrl

if ($DumpOnly) {
    [IO.File]::WriteAllText($DumpPath, $html, [Text.Encoding]::UTF8)
    Write-Host ("HTMLを保存しました: {0}" -f $DumpPath)
    Write-Host 'このファイルを開いて、スケジュール表の構造を確認してください。'
    exit 0
}

$schedule = Find-ScheduleTable -Html $html
if (-not $schedule) {
    [IO.File]::WriteAllText($DumpPath, $html, [Text.Encoding]::UTF8)
    Write-Warning 'スケジュール表を自動検出できませんでした。'
    Write-Warning ("取得したHTMLを保存しました: {0}" -f $DumpPath)
    Write-Warning 'このHTMLダンプを共有してもらえれば、パーサーをページ構造に合わせて調整できます。'
    exit 1
}

Write-Host ("検出: メンバー {0} 名 / 日付列 {1} 列" -f $schedule.Members.Count, $schedule.DateHeaders.Count)

$text = Build-SlackText -Schedule $schedule -Today (Get-Date) -IncludeAllDays:$AllDays -GroupName $GroupName

if ($NoPost) {
    Write-Host '--- 投稿内容 (NoPostモード) ---'
    Write-Host $text
    exit 0
}

if (-not $WebhookUrl) {
    throw '-WebhookUrl か環境変数 SLACK_WEBHOOK_URL に Slack Incoming Webhook URL を設定してください。'
}

Send-SlackMessage -Text $text -Url $WebhookUrl
Write-Host 'Slackに投稿しました。'
