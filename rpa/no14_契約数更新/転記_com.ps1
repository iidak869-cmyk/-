# Excel本体(COM)を使ってプロダクト日報にセルを追記するスクリプト
# 契約数更新.js から自動で呼ばれる（単体実行は不要）
#
# JSON形式: [ { "A": "2026/07/15", "C": "会社名", "M": 2640000, ... }, ... ]
param(
  [Parameter(Mandatory = $true)][string]$JsonPath,
  [Parameter(Mandatory = $true)][string]$ExcelPath,
  [Parameter(Mandatory = $true)][string]$SheetName
)

$rows = Get-Content $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $rows -or $rows.Count -eq 0) {
  Write-Host "追記対象がありません"
  exit 0
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null

try {
  $wb = $excel.Workbooks.Open($ExcelPath)
  $ws = $wb.Worksheets.Item($SheetName)

  # A列の最終データ行を探す（xlUp = -4162）
  $last = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row

  foreach ($row in $rows) {
    $last += 1
    foreach ($prop in $row.PSObject.Properties) {
      $v = $prop.Value
      # 日付らしき文字列はExcelの日付として入れる
      if ($v -is [string] -and $v -match '^\d{4}/\d{1,2}/\d{1,2}$') {
        $v = [datetime]$v
      }
      $ws.Range("$($prop.Name)$last").Value2 = $v
    }
  }

  $wb.Save()
  Write-Host "Excelへ追記完了: $($rows.Count)行（$($SheetName)シート 〜$last行目）"
}
finally {
  if ($wb) { $wb.Close($true) }
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
