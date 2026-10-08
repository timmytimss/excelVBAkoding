Import-Module "$PSScriptRoot\get-testfil-wb.psm1" -Force
$excelApp = Get-TestfilExcelApp
$excelApp.Visible = $false
try { $excelApp.VBE.MainWindow.Visible = $false } catch {}
$wb = Get-TestfilWorkbook
try { $wb.Worksheets.Item("StatTest").Delete() } catch {}
$ws = $wb.Worksheets.Add(); $ws.Name = "StatTest"
$ws.Range("A1").Value2 = "Status"; $ws.Range("B1").Value2 = "Kommentar"
$vals = @("Alpha","Bravo","Charlie","Delta","Echo","Foxtrot","Zulu")
$subs = @("Ring","Purr","","Avvent","Noter")
$r = 2
foreach ($v in $vals) { for ($k=0; $k -lt (2 + ($vals.IndexOf($v) % 3)); $k++) {
  $ws.Cells.Item($r,1).Value2 = $v; $ws.Cells.Item($r,2).Value2 = $subs[($r) % 5]; $r++ } }
$cols = @(255, 65280, 16711680, 65535, 16744448, 8421631, 12632256)
for ($i=0; $i -lt $vals.Count; $i++) {
  for ($rr=2; $rr -lt $r; $rr++) { if ($ws.Cells.Item($rr,1).Value2 -eq $vals[$i]) { $ws.Cells.Item($rr,1).Interior.Color = $cols[$i] } } }
$wb.Worksheets.Item("StatTest").Activate()
$tbl = $ws.ListObjects.Add(1, $ws.Range("A1:B$($r-1)"), $null, 1); $tbl.Name = "StatTestTbl"
$o = $excelApp.Run("modStatistikkern.GetOppsettSheet")
$o.Range("A2").Value2="StatTest"; $o.Range("B2").Value2="StatTestTbl"; $o.Range("C2").Value2="Status"; $o.Range("D2").Value2="Kommentar"
Write-Output "ok rader=$($r-2)"
