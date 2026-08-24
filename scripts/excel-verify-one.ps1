param(
  [Parameter(Mandatory = $true)]
  [string]$File
)

$resolved = (Resolve-Path -LiteralPath $File).Path
$before = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
try {
  $workbook = $excel.Workbooks.Open($resolved, 0, $true)
  try {
    $sheet = $workbook.Worksheets.Item(1)
    $links = $workbook.LinkSources(1)
    [pscustomobject]@{
      File = [IO.Path]::GetFileName($resolved)
      Opened = $true
      Links = if ($null -eq $links) { 0 } else { @($links).Count }
      WindowView = $excel.ActiveWindow.View
      PrintArea = $sheet.PageSetup.PrintArea
      PrintTitleRows = $sheet.PageSetup.PrintTitleRows
      CenterFooter = $sheet.PageSetup.CenterFooter
      HorizontalPageBreaks = $sheet.HPageBreaks.Count
    }
  } finally {
    $workbook.Close($false)
  }
} finally {
  $excel.Quit()
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
}
$after = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
[pscustomobject]@{ NewRepairLogs = @($after | Where-Object { $_ -notin $before }).Count }
