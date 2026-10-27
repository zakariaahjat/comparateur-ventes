$ErrorActionPreference = 'Stop'

$files = @(
    @{ Month = 'Mois 08'; Path = 'C:\Users\zakaria01\Downloads\Mois_08.xlsx' },
    @{ Month = 'Mois 09'; Path = 'C:\Users\zakaria01\Downloads\Mois_09.xlsx' },
    @{ Month = 'Mois 10'; Path = 'C:\Users\zakaria01\Downloads\Mois_10.xlsx' }
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# ---- 1. READ the 3 files, merge by EAN ----
$products = @{}   # EAN -> hashtable { Name, Qty per month }
$paths = @($files | ForEach-Object { $_.Path })

foreach ($f in $files) {
    $wb = $excel.Workbooks.Open($f.Path, 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $last = $ws.UsedRange.Rows.Count
    $qi = [array]::IndexOf($paths, $f.Path)
    for ($r = 2; $r -le $last; $r++) {
        $txt = ($ws.Cells.Item($r, 1).Text).Trim()
        $qt  = $ws.Cells.Item($r, 2).Value2
        if ([string]::IsNullOrWhiteSpace($txt)) { continue }
        if ($txt -match '\[(\d+)\]') {
            $ean = $Matches[1]
            $name = ($txt -replace '\[\d+\]', '').Trim()
        } else {
            $ean = $txt                    # no EAN -> use full text as key
            $name = $txt
        }
        if (-not $products.ContainsKey($ean)) {
            $products[$ean] = @{ Name = $name; Q1 = 0; Q2 = 0; Q3 = 0 }
        }
        $products[$ean].Name = $name
        if ($null -ne $qt -and $qt -ne '') {
            $q = [int]$qt
            if ($qi -eq 0) { $products[$ean].Q1 += $q }
            elseif ($qi -eq 1) { $products[$ean].Q2 += $q }
            else { $products[$ean].Q3 += $q }
        }
    }
    $wb.Close($false)
    Write-Output ("Read {0}: {1} rows" -f $f.Month, ($last - 1))
}

# ---- 2. WRITE the comparison table ----
$out = New-Object -ComObject Excel.Application
$out.Visible = $false
$out.DisplayAlerts = $false
$wbo = $out.Workbooks.Add()
$ws = $wbo.Worksheets.Item(1)
$ws.Name = 'Comparaison'

$headers = 'EAN', 'Produit', 'Mois 08 (Qté)', 'Mois 09 (Qté)', 'Mois 10 (Qté)', 'Total', 'Meilleur Mois (Qté Max)'
for ($c = 0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(1, $c + 1) = $headers[$c] }

$i = 1
foreach ($e in $products.GetEnumerator()) {
    $p = $e.Value
    $i++
    $best = ($p.Q1, $p.Q2, $p.Q3 | Measure-Object -Maximum).Maximum
    $bestMonth = if ($p.Q1 -eq $best) { 'Mois 08' } elseif ($p.Q2 -eq $best) { 'Mois 09' } else { 'Mois 10' }
    if ($best -eq 0) { $bestMonth = '-' }
    $ws.Cells.Item($i, 1).NumberFormat = '@'
    $ws.Cells.Item($i, 1) = $e.Key
    $ws.Cells.Item($i, 2) = $p.Name
    $ws.Cells.Item($i, 3) = $p.Q1
    $ws.Cells.Item($i, 4) = $p.Q2
    $ws.Cells.Item($i, 5) = $p.Q3
    $ws.Cells.Item($i, 6) = $p.Q1 + $p.Q2 + $p.Q3
    $ws.Cells.Item($i, 7) = if ($best -eq 0) { '-' } else { "$bestMonth ($best)" }
}

# formatting
$ws.Rows.Item(1).Font.Bold = $true
$ws.Rows.Item(1).Interior.Color = 15982569   # light yellow
$ws.Activate()
$ws.Application.ActiveWindow.FreezePanes = $false
$ws.Application.Goto($ws.Cells.Item(2, 1), $true)
$ws.Application.ActiveWindow.FreezePanes = $true
$ws.Range("A1:G$i").AutoFilter() | Out-Null
$ws.Columns.AutoFit() | Out-Null

$dest = 'C:\Users\zakaria01\Downloads\sepraration project\Comparaison_Ventes.xlsx'
$wbo.SaveAs($dest, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
Write-Output "Saved: $dest ($($i - 1) products)"

$wbo.Close($false)
$out.Quit()
$excel.Quit()
