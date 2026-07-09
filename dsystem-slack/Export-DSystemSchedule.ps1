<#
.SYNOPSIS
    Dシステム (D-System 業務カレンダー) から CS課メンバーの予定を取得し、
    Slack の Incoming Webhook に投稿します。

.DESCRIPTION
    社内ネットワーク上の Windows PC / サーバーで実行してください
    (win2012sv-aws は社外・クラウドからは到達できません)。

    ブラウザでの操作をそのまま再現します:
      1) ログインページに ID/パスワードでログイン
      2) 表示順を「グループ順」に切り替え
      3) 左側のグループ一覧から対象グループ (CS課) だけにチェックして「絞込実行」
      4) スケジュール表 (sTable) を解析して、日付ごと・人ごとに整形
      5) Slack Incoming Webhook へ投稿 (既定は今日の分のみ)

    使い方:
      # 動作確認 (Slackに送らずコンソール表示)
      .\Export-DSystemSchedule.ps1 -NoPost -UserId "xxxx" -Password "yyyy"

      # 本番 (Slackに投稿)
      .\Export-DSystemSchedule.ps1 -UserId "xxxx" -Password "yyyy" `
          -WebhookUrl 'https://hooks.slack.com/services/XXX/YYY/ZZZ'

.NOTES
    PowerShell 3.0 以降 (Windows Server 2012 標準) で動作するよう記述しています。
    ファイルは UTF-8 (BOM付き) で保存してください。
#>
[CmdletBinding()]
param(
    # DシステムのスケジュールページURL
    [string]$ScheduleUrl = 'http://win2012sv-aws/DSystem/Schedule.aspx',

    # 抜き出したいグループ名 (左サイドバーのグループ一覧の表記)
    [string]$GroupName = 'CS課',

    # DシステムのログインID (未指定なら環境変数 DSYSTEM_USER を使用)
    [string]$UserId = $env:DSYSTEM_USER,

    # Dシステムのログインパスワード (未指定なら環境変数 DSYSTEM_PASSWORD を使用)
    [string]$Password = $env:DSYSTEM_PASSWORD,

    # Slack Incoming Webhook URL (未指定なら環境変数 SLACK_WEBHOOK_URL を使用)
    [string]$WebhookUrl = $env:SLACK_WEBHOOK_URL,

    # 絞り込み後のページHTMLを保存するだけで終了する (構造確認用)
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

# ログイン状態 (Cookie) を保持するセッション
$script:WebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession

# ---- HTTP / HTML ユーティリティ ---------------------------------------------

function Get-ResponseText {
    param($Res)
    $bytes = $Res.RawContentStream.ToArray()

    # 文字コードを Content-Type ヘッダー → HTML内 meta → UTF-8 の順で判定
    $charset = $null
    $ct = $Res.Headers['Content-Type']
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

function Get-PageHtml {
    param([string]$Url)
    $res = Invoke-WebRequest -Uri $Url -UseBasicParsing -WebSession $script:WebSession
    return Get-ResponseText $res
}

function ConvertFrom-HtmlText {
    param([string]$Html)
    $t = $Html -replace '(?i)<br\s*/?>', "`n"
    $t = $t -replace '(?s)<[^>]+>', ''
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    $lines = $t -split "`n" | ForEach-Object { ($_ -replace '[\s　]+', ' ').Trim() } | Where-Object { $_ }
    return ($lines -join "`n")
}

function Test-IsLoginPage {
    param([string]$Html)
    return ($Html -match 'type="password"' -or $Html -match 'Login Page')
}

# ---- ASP.NET フォーム操作 ----------------------------------------------------

function ConvertTo-FormBody {
    # フォーム値を application/x-www-form-urlencoded 形式に変換する
    # ([Uri]::EscapeDataString は約32,000文字の上限があり巨大な __VIEWSTATE で失敗するため WebUtility を使用)
    param($Fields)
    $pairs = @()
    foreach ($e in $Fields.GetEnumerator()) {
        $pairs += ('{0}={1}' -f `
            [System.Net.WebUtility]::UrlEncode([string]$e.Key), `
            [System.Net.WebUtility]::UrlEncode([string]$e.Value))
    }
    return ($pairs -join '&')
}

function Get-HiddenFields {
    # __VIEWSTATE などASP.NETの隠しフィールドを集める
    param([string]$Html)
    $fields = @{}
    foreach ($m in [regex]::Matches($Html, '(?is)<input[^>]+type="hidden"[^>]*>')) {
        $tag = $m.Value
        if ($tag -match 'name="([^"]+)"') {
            $name = $Matches[1]
            $value = ''
            if ($tag -match 'value="([^"]*)"') { $value = $Matches[1] }
            $fields[$name] = [System.Net.WebUtility]::HtmlDecode($value)
        }
    }
    return $fields
}

function Get-FormFields {
    # ページ内フォームの現在値 (hidden / text / チェック済みradio / チェック済みcheckbox / select) を集める
    param([string]$Html)
    $fields = [ordered]@{}

    foreach ($m in [regex]::Matches($Html, '(?is)<input[^>]*>')) {
        $tag = $m.Value
        if ($tag -notmatch 'name\s*=\s*["'']([^"'']+)["'']') { continue }
        $name = $Matches[1]
        $type = 'text'
        if ($tag -match 'type\s*=\s*["'']([^"'']+)["'']') { $type = $Matches[1].ToLower() }
        $value = ''
        if ($tag -match 'value\s*=\s*["'']([^"'']*)["'']') {
            $value = [System.Net.WebUtility]::HtmlDecode($Matches[1])
        }
        switch ($type) {
            'hidden'   { $fields[$name] = $value }
            'text'     { $fields[$name] = $value }
            'radio'    { if ($tag -match '\bchecked\b') { $fields[$name] = $value } }
            'checkbox' {
                if ($tag -match '\bchecked\b') {
                    if ($value) { $fields[$name] = $value } else { $fields[$name] = 'on' }
                }
            }
            default { }  # submit / button / password などは含めない
        }
    }

    foreach ($m in [regex]::Matches($Html, '(?is)<select[^>]+name\s*=\s*["'']([^"'']+)["''][^>]*>(.*?)</select>')) {
        $name = $m.Groups[1].Value
        $body = $m.Groups[2].Value
        $val = ''
        foreach ($om in [regex]::Matches($body, '(?is)<option([^>]*)>')) {
            $attrs = $om.Groups[1].Value
            if ($attrs -match '\bselected\b') {
                if ($attrs -match 'value\s*=\s*["'']([^"'']*)["'']') { $val = $Matches[1] }
                break
            }
        }
        $fields[$name] = $val
    }
    return $fields
}

function Invoke-PostBack {
    <#
      現在のページ ($Html) のフォーム値を引き継いで POST し、応答HTMLを返す。
      - $Set          : 上書き・追加するフィールド
      - $EventTarget  : __doPostBack のターゲット (ドロップダウン変更など)
      - $SubmitName / $SubmitValue : 押すボタン (絞込実行など)
      - $RemoveLike   : 送信から除外するフィールド名パターン (checkbox のチェック外しに使用)
    #>
    param(
        [string]$Url,
        [string]$Html,
        [hashtable]$Set = @{},
        [string]$EventTarget = '',
        [string]$EventArgument = '',
        [string]$SubmitName = '',
        [string]$SubmitValue = '',
        [string[]]$RemoveLike = @()
    )
    $fields = Get-FormFields $Html
    foreach ($k in @($fields.Keys)) {
        foreach ($pat in $RemoveLike) {
            if ($k -like $pat) { $fields.Remove($k); break }
        }
    }
    foreach ($k in $Set.Keys) { $fields[$k] = $Set[$k] }
    $fields['__EVENTTARGET'] = $EventTarget
    $fields['__EVENTARGUMENT'] = $EventArgument
    if ($SubmitName) { $fields[$SubmitName] = $SubmitValue }

    $body = ConvertTo-FormBody $fields

    $res = Invoke-WebRequest -Uri $Url -Method Post -Body $body `
        -ContentType 'application/x-www-form-urlencoded' `
        -UseBasicParsing -WebSession $script:WebSession
    return Get-ResponseText $res
}

function Invoke-DSystemLogin {
    param([string]$PageUrl, [string]$Html, [string]$UserId, [string]$Password)

    # フォームの action 属性からログインPOST先URLを求める
    $action = './Default.aspx'
    if ($Html -match '(?is)<form[^>]+action="([^"]+)"') {
        $action = [System.Net.WebUtility]::HtmlDecode($Matches[1])
    }
    $postUrl = (New-Object System.Uri((New-Object System.Uri($PageUrl)), $action)).AbsoluteUri

    # ID・パスワード・ログインボタンの入力欄名をHTMLから自動検出 (取れなければ既知の名前)
    $userField = 'ctl00$ContentPlaceHolder1$txtUserID'
    $passField = 'ctl00$ContentPlaceHolder1$txtUserPass'
    $btnField  = 'ctl00$ContentPlaceHolder1$btnLogin'
    $btnValue  = 'ログイン'
    if ($Html -match '(?is)<input\s+name="([^"]+)"[^>]*type="text"')     { $userField = $Matches[1] }
    if ($Html -match '(?is)<input\s+name="([^"]+)"[^>]*type="password"') { $passField = $Matches[1] }
    if ($Html -match '(?is)<input\s+type="submit"\s+name="([^"]+)"\s+value="([^"]*)"') {
        $btnField = $Matches[1]
        $btnValue = [System.Net.WebUtility]::HtmlDecode($Matches[2])
    }

    $fields = Get-HiddenFields $Html
    $fields[$userField] = $UserId
    $fields[$passField] = $Password
    $fields[$btnField]  = $btnValue

    $body = ConvertTo-FormBody $fields

    $res = Invoke-WebRequest -Uri $postUrl -Method Post -Body $body `
        -ContentType 'application/x-www-form-urlencoded' `
        -UseBasicParsing -WebSession $script:WebSession
    return Get-ResponseText $res
}

# ---- Dシステム固有の処理 ------------------------------------------------------

function Get-GroupCheckboxes {
    # 左サイドバーのグループ絞り込みチェックボックス一覧 (チェック状態付き) を取得
    param([string]$Html)
    $list = @()
    $pattern = "(?is)<input type='checkbox' name=""(cbGroup\d+)"" value=""(\d+)""([^>]*)>([^<]*)</label>"
    foreach ($m in [regex]::Matches($Html, $pattern)) {
        $label = [System.Net.WebUtility]::HtmlDecode($m.Groups[4].Value)
        $label = ($label -replace '&nbsp;', ' ' -replace '[\s　]+', ' ').Trim()
        $list += ,@{
            Field   = $m.Groups[1].Value
            Value   = $m.Groups[2].Value
            Checked = ($m.Groups[3].Value -match 'checked')
            Label   = $label
        }
    }
    return $list
}

function Get-GroupState {
    # 指定グループのチェック状態を返す ($true/$false、一覧に無ければ $null)
    param([string]$Html, [string]$Value)
    $m = [regex]::Match($Html, "(?is)<input type='checkbox' name=""cbGroup$Value"" value=""$Value""([^>]*)>")
    if (-not $m.Success) { return $null }
    return ($m.Groups[1].Value -match 'checked')
}

function Set-GroupCheck {
    <#
      グループのチェックON/OFFを切り替える。
      画面のチェックボックスは fncGroupClick() → __doPostBack("GroupClick", "グループ番号,フラグ")
      でサーバー側に記憶される方式のため、それをそのまま再現する。
      フラグの意味 (現在の状態か新しい状態か) はサーバー実装次第なので、
      1回目で変わらなければ逆のフラグでもう一度試し、結果をHTMLで検証する。
    #>
    param([string]$Url, [string]$Html, [string]$Value, [bool]$Desired)

    $cur = Get-GroupState -Html $Html -Value $Value
    if ($null -eq $cur) {
        Write-Warning ("グループ番号 {0} のチェックボックスが画面に見つかりません。" -f $Value)
        return $Html
    }
    if ($cur -eq $Desired) { return $Html }

    $flagCurrent = 0; if ($cur) { $flagCurrent = 1 }
    $h2 = Invoke-PostBack -Url $Url -Html $Html `
        -EventTarget 'GroupClick' -EventArgument ("{0},{1}" -f $Value, $flagCurrent)
    if ((Get-GroupState -Html $h2 -Value $Value) -eq $Desired) { return $h2 }

    $flagDesired = 0; if ($Desired) { $flagDesired = 1 }
    $h3 = Invoke-PostBack -Url $Url -Html $h2 `
        -EventTarget 'GroupClick' -EventArgument ("{0},{1}" -f $Value, $flagDesired)
    if ((Get-GroupState -Html $h3 -Value $Value) -eq $Desired) { return $h3 }

    Write-Warning ("グループ番号 {0} のチェック状態を変更できませんでした。" -f $Value)
    return $h3
}

function Get-DSystemSchedule {
    # sTable (スケジュール表) を解析して日付ごとの予定リストを返す
    param([string]$Html)

    $m = [regex]::Match($Html, '(?is)<table id="sTable".*?</table>')
    if (-not $m.Success) { return $null }
    $table = $m.Value

    # ヘッダー <th> から日付列を取得 (例: "2026/07/09 (木)\n32件")
    $days = @()
    foreach ($th in [regex]::Matches($table, '(?is)<th[^>]*>(.*?)</th>')) {
        $text = ConvertFrom-HtmlText $th.Groups[1].Value
        if ($text -match '(\d{4})/(\d{1,2})/(\d{1,2})') {
            $days += ,@{
                Year  = [int]$Matches[1]
                Month = [int]$Matches[2]
                Day   = [int]$Matches[3]
                Label = ($text -replace "`n", '・')
                Items = @()
            }
        }
    }
    if ($days.Count -eq 0) { return $null }

    # データ行の <td> (日付列と同順) から <li> を抜き出す
    $tds = @([regex]::Matches($table, '(?is)<td[^>]*>(.*?)</td>'))
    for ($i = 0; $i -lt [Math]::Min($tds.Count, $days.Count); $i++) {
        $items = @()
        foreach ($li in [regex]::Matches($tds[$i].Groups[1].Value, '(?is)<li[^>]*>(.*?)</li>')) {
            $t = ConvertFrom-HtmlText $li.Groups[1].Value
            if ($t) { $items += ($t -replace "`n", ' ') }
        }
        $days[$i].Items = $items
    }
    return $days
}

function Split-PersonItem {
    # 予定1件のテキストを「名前」と「内容」に分ける
    # 例: "山下 毎日業務 録音格納" → 山下 / 毎日業務 録音格納
    #     "伊藤そ★1000 Cytekiサポート" → 伊藤そ / ★1000 Cytekiサポート
    param([string]$Item)
    if ($Item -match '^([^\s　★]{1,10})[\s　]*(★.*|[\s　].*)$') {
        $name = $Matches[1]
        $body = $Matches[2].Trim()
        if ($body) { return @($name, $body) }
    }
    if ($Item -match '^([^\s　★]{1,10})$') { return @($Matches[1], '(タイトルなし)') }
    return @('（グループ共通）', $Item)
}

function Build-SlackText {
    param($Days, [datetime]$Today, [bool]$IncludeAllDays, [string]$GroupName)

    # 投稿対象の日付列 (既定: 今日。見つからなければ全列)
    $targets = @()
    if (-not $IncludeAllDays) {
        $targets = @($Days | Where-Object {
            $_.Year -eq $Today.Year -and $_.Month -eq $Today.Month -and $_.Day -eq $Today.Day
        })
    }
    if ($targets.Count -eq 0) { $targets = @($Days) }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(("*【{0} スケジュール】*" -f $GroupName))

    foreach ($day in $targets) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine(("*■ {0}*" -f $day.Label))

        if ($day.Items.Count -eq 0) {
            [void]$sb.AppendLine('  (予定なし)')
            continue
        }

        # 人ごとにまとめる (ページの表示順を維持)
        $order = New-Object System.Collections.ArrayList
        $byPerson = @{}
        foreach ($item in $day.Items) {
            $parts = Split-PersonItem $item
            $name = $parts[0]; $body = $parts[1]
            if (-not $byPerson.ContainsKey($name)) {
                [void]$order.Add($name)
                $byPerson[$name] = New-Object System.Collections.ArrayList
            }
            [void]$byPerson[$name].Add($body)
        }
        foreach ($name in $order) {
            [void]$sb.AppendLine(("*{0}*" -f $name))
            foreach ($body in $byPerson[$name]) {
                [void]$sb.AppendLine(("　・{0}" -f $body))
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

function Save-Dump {
    param([string]$Html, [string]$Path)
    [IO.File]::WriteAllText($Path, $Html, [Text.Encoding]::UTF8)
}

# ---- メイン処理 ---------------------------------------------------------------

Write-Host ("ページ取得中: {0}" -f $ScheduleUrl)
$html = Get-PageHtml -Url $ScheduleUrl

# 1) 必要ならログイン
if (Test-IsLoginPage $html) {
    if (-not $UserId -or -not $Password) {
        throw 'Dシステムへのログインが必要です。-UserId と -Password (または環境変数 DSYSTEM_USER / DSYSTEM_PASSWORD) を指定してください。'
    }
    Write-Host 'ログインページが表示されたため、ログインします...'
    [void](Invoke-DSystemLogin -PageUrl $ScheduleUrl -Html $html -UserId $UserId -Password $Password)

    $html = Get-PageHtml -Url $ScheduleUrl
    if (Test-IsLoginPage $html) {
        throw 'ログインに失敗しました。ログインIDとパスワードが正しいか確認してください。'
    }
    Write-Host 'ログイン成功。'
}

# 2) 表示順を「グループ順」に切り替え (グループの絞り込み一覧を出すため)
Write-Host '表示順を「グループ順」に切り替えます...'
$html = Invoke-PostBack -Url $ScheduleUrl -Html $html -EventTarget 'ddlSort' -Set @{ 'ddlSort' = '2' }

# 3) 「全表示（未チェックも表示）」に切り替えて全グループの一覧を出し、対象グループを探す
#    (「自分のチェックのみ」モードでは登録済みグループしか一覧に出ないため)
Write-Host '表示モードを「全表示（未チェックも表示）」に切り替えます...'
$html = Invoke-PostBack -Url $ScheduleUrl -Html $html `
    -EventTarget 'rblGroupDisp$1' -Set @{ 'rblGroupDisp' = '全表示（未チェックも表示）' }

$groups = @(Get-GroupCheckboxes $html)
$target = $groups | Where-Object { $_.Label -like ("*{0}*" -f $GroupName) } | Select-Object -First 1

if ($target) {
    # 4) 対象グループ以外のチェックを外し、対象グループにチェックを入れる
    #    (チェックは GroupClick ポストバックでサーバー側に記憶される)
    foreach ($g in $groups) {
        if ($g.Value -ne $target.Value -and $g.Checked) {
            Write-Host ("グループ「{0}」({1}) のチェックを外します..." -f $g.Label, $g.Value)
            $html = Set-GroupCheck -Url $ScheduleUrl -Html $html -Value $g.Value -Desired:$false
        }
    }
    Write-Host ("グループ「{0}」({1}) にチェックを入れます..." -f $target.Label, $target.Value)
    $html = Set-GroupCheck -Url $ScheduleUrl -Html $html -Value $target.Value -Desired:$true

    # 5) 「自分のチェックのみ」に戻すと、チェック済みグループ (=対象グループ) だけの表示になる
    Write-Host '表示モードを「自分のチェックのみ」に戻します...'
    $html = Invoke-PostBack -Url $ScheduleUrl -Html $html `
        -EventTarget 'rblGroupDisp$2' -Set @{ 'rblGroupDisp' = '自分のチェックのみ' }
} else {
    $labels = (($groups | ForEach-Object { $_.Label }) | Select-Object -Unique) -join ' / '
    Write-Warning ("グループ「{0}」がグループ一覧に見つかりませんでした。見つかったグループ: {1}" -f $GroupName, $labels)
    Write-Warning '絞り込みなしで続行します。ダンプを保存するので共有してください。'
    Save-Dump -Html $html -Path $DumpPath
}

if ($DumpOnly) {
    Save-Dump -Html $html -Path $DumpPath
    Write-Host ("HTMLを保存しました: {0}" -f $DumpPath)
    Write-Host 'このファイルを開いて、スケジュール表の構造を確認してください。'
    exit 0
}

# 4) スケジュール表を解析
$days = Get-DSystemSchedule -Html $html
if (-not $days) {
    Save-Dump -Html $html -Path $DumpPath
    Write-Warning 'スケジュール表 (sTable) を検出できませんでした。'
    Write-Warning ("取得したHTMLを保存しました: {0}" -f $DumpPath)
    Write-Warning 'このHTMLダンプを共有してもらえれば、パーサーをページ構造に合わせて調整できます。'
    exit 1
}

$total = 0
foreach ($d in $days) { $total += $d.Items.Count }
Write-Host ("検出: {0} 日分 / 合計 {1} 件" -f $days.Count, $total)

# 5) 整形して投稿
$text = Build-SlackText -Days $days -Today (Get-Date) -IncludeAllDays:$AllDays -GroupName $GroupName

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
