param(
  [Parameter(Mandatory = $true)]
  [string]$Path
)

$resolved = (Resolve-Path -LiteralPath $Path).Path
$before = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$results = @()
try {
  Get-ChildItem -LiteralPath $resolved -Filter 'ZCJY*.xlsx' | Sort-Object Name | ForEach-Object {
    $fileName = $_.Name
    $workbook = $excel.Workbooks.Open($_.FullName, 0, $true)
    try {
      $links = $workbook.LinkSources(1)
      $results += [pscustomobject]@{
        File = $fileName
        Opened = $true
        Links = if ($null -eq $links) { 0 } else { @($links).Count }
        WindowView = $excel.ActiveWindow.View
      }
    } finally {
      $workbook.Close($false)
    }
  }
} finally {
  $excel.Quit()
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
}
$after = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
$results
[pscustomobject]@{ NewRepairLogs = @($after | Where-Object { $_ -notin $before }).Count }
