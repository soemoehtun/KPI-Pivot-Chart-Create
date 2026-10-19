param([Parameter(Mandatory=$true)][string]$ConfigPath)
$ErrorActionPreference = "Stop"
$cfg = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$missing = [System.Reflection.Missing]::Value
# category-axis date format: date only, or date + 12-hour time when a time column is used
$dtFmt = if ($cfg.hasTime) { "dd-mmm hh:mm AM/PM" } else { "dd-mmm" }

function Set-NoGrid($sheet) {
    try { $sheet.Activate(); $excel.ActiveWindow.DisplayGridlines = $false } catch {}
}
function Style-Pivot($pt) {
    try { $pt.TableStyle2 = "PivotStyleMedium2" } catch {}
    $pr = $pt.TableRange2
    try { $pr.HorizontalAlignment = -4108 } catch {}
    try { $pr.EntireColumn.AutoFit() } catch {}
    try { $pr.Borders.LineStyle = 1; $pr.Borders.Weight = 2 } catch {}
}
function Style-Chart($co, $title) {
    try { $co.Placement = 3 } catch {}   # xlFreeFloating = don't move or size with cells
    $co.Chart.ChartType = [int]$cfg.chartType
    $co.Chart.HasTitle = $true
    try { $co.Chart.ChartTitle.Text = $title } catch {}
    try { $co.Chart.ChartTitle.Font.Size = 11 } catch {}   # heading font size
    $co.Chart.ShowAllFieldButtons = $false
    try { $co.Chart.HasLegend = $true; $co.Chart.Legend.Position = -4107 } catch {}
    try { $co.Chart.Legend.Font.Size = 7 } catch {}
    try { $co.Chart.Axes(1,1).TickLabels.Orientation = -4171 } catch {}
    try { $co.Chart.Axes(1,1).TickLabels.Font.Size = 7 } catch {}
    try { $co.Chart.Axes(1,1).TickLabels.NumberFormat = $dtFmt } catch {}
    try { $co.Chart.Axes(2,1).TickLabels.Font.Size = 7 } catch {}
    try { $co.Chart.Axes(1,1).AxisTitle.Font.Size = 7 } catch {}
    try { $co.Chart.Axes(2,1).AxisTitle.Font.Size = 7 } catch {}
    try {
        $sc = $co.Chart.SeriesCollection()
        for ($s = 1; $s -le $sc.Count; $s++) {
            try { $sc.Item($s).Format.Line.Weight = 1.5 } catch {}   # series line width (1 1/2 pt)
            try { $sc.Item($s).DataLabels.Font.Size = 7 } catch {}
        }
    } catch {}
}
function Slicer-Count($field) {
    try { return [int]$cfg.slicerItems.$field } catch {}
    try {
        $prop = $cfg.slicerItems.PSObject.Properties | Where-Object { $_.Name -eq $field }
        if ($prop) { return [int]$prop.Value }
    } catch {}
    return 1
}
function Add-PivotFields($pt, $col) {
    # build with ManualUpdate so the real-date field is NOT auto-grouped into
    # Years/Quarters/Months (keeps chronological order), then format the labels.
    $pt.ManualUpdate = $true
    $pt.PivotFields("DateTime").Orientation = 1
    if ($cfg.seriesField) { $pt.PivotFields($cfg.seriesField).Orientation = 2 }
    $fld = $pt.AddDataField($pt.PivotFields($col), "$($cfg.aggLabel) of $col", [int]$cfg.aggFunc)
    $fld.NumberFormat = "0.00"
    try { $pt.PivotFields("DateTime").NumberFormat = $dtFmt } catch {}
    $pt.ManualUpdate = $false
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
$wb = $null
try {
    $wb = $excel.Workbooks.Open($cfg.dataPath)
    try { $excel.Calculation = -4135 } catch {}   # xlCalculationManual for speed
    $wsData = $wb.Worksheets.Item("Raw Data")
    $pc = $wb.PivotCaches().Create(1, $wsData.Range("A1").CurrentRegion)

    $slicers = @($cfg.slicers | Where-Object { $_ })
    $caches = @{}
    $pts = @()
    $slicerW = 2.22 * 72
    $gap = 10
    $slicerHBase = 22
    $slicerHItem = 19

    if ($cfg.layout -eq "per_sheet") {
        # ---- one KPI per sheet (pivot + chart + slicers side by side) ----
        $chartW = 520
        $chartH = 300
        for ($i = 0; $i -lt $cfg.kpis.Count; $i++) {
            $col   = $cfg.kpis[$i].col
            $sheet = $cfg.kpis[$i].sheet
            $idx   = $i + 1

            $ws = $wb.Worksheets.Add($missing, $wb.Worksheets.Item($wb.Worksheets.Count))
            $ws.Name = $sheet
            Set-NoGrid $ws

            $pt = $pc.CreatePivotTable($ws.Range("B3"), "PT_$idx")
            Add-PivotFields $pt $col
            $pts += $pt
            Style-Pivot $pt

            $pr = $pt.TableRange2
            $lastCol = $pr.Column + $pr.Columns.Count - 1
            $chartLeft = $ws.Cells(1, $lastCol + 2).Left
            $chartTop  = $ws.Range("B3").Top
            $co = $ws.ChartObjects().Add($chartLeft, $chartTop, $chartW, $chartH)
            [void]$co.Chart.SetSourceData($pt.TableRange1)
            Style-Chart $co $col

            $sLeft = $chartLeft + $chartW + 12
            $s = 0
            foreach ($field in $slicers) {
                if (-not $caches.ContainsKey($field)) {
                    $cache = $wb.SlicerCaches.Add2($pt, $field)
                    try { $cache.ShowAllItems = $false } catch {}
                    $caches[$field] = $cache
                }
                $n = Slicer-Count $field
                $h = $slicerHBase + ($n * $slicerHItem)
                $left = $sLeft + $s * ($slicerW + $gap)
                $sn = "Slicer_" + ($field -replace '\W','') + "_$idx"
                $sl = $caches[$field].Slicers.Add($ws, $missing, $sn, $field, $chartTop, $left, $slicerW, $h)
                try { $sl.Shape.Placement = 3 } catch {}   # don't move or size with cells
                $s++
            }
        }

        foreach ($field in $slicers) {
            foreach ($p in $pts) { [void]$caches[$field].PivotTables.AddPivotTable($p) }
        }
    }
    else {
        # ---- all KPIs on one Dashboard; pivots on one sheet (all_pivot) or one per sheet (one_pivot) ----
        if ($cfg.layout -eq "all_pivot") {
            $firstPivotSheet = $cfg.pivotSheetName
            $wsPivot = $wb.Worksheets.Add($missing, $wb.Worksheets.Item($wb.Worksheets.Count))
            $wsPivot.Name = $cfg.pivotSheetName
            Set-NoGrid $wsPivot
            $rowCursor = 3
            for ($i = 0; $i -lt $cfg.kpis.Count; $i++) {
                $col = $cfg.kpis[$i].col
                $idx = $i + 1
                $pt = $pc.CreatePivotTable($wsPivot.Cells($rowCursor, 2), "PT_$idx")
                Add-PivotFields $pt $col
                Style-Pivot $pt
                $pts += $pt
                $rowCursor = $pt.TableRange2.Row + $pt.TableRange2.Rows.Count + 3
            }
        }
        else {
            $firstPivotSheet = $cfg.kpis[0].sheet
            for ($i = 0; $i -lt $cfg.kpis.Count; $i++) {
                $col   = $cfg.kpis[$i].col
                $sheet = $cfg.kpis[$i].sheet
                $idx   = $i + 1

                $ws = $wb.Worksheets.Add($missing, $wb.Worksheets.Item($wb.Worksheets.Count))
                $ws.Name = $sheet
                Set-NoGrid $ws

                $pt = $pc.CreatePivotTable($ws.Range("B3"), "PT_$idx")
                Add-PivotFields $pt $col
                Style-Pivot $pt
                $pts += $pt
            }
        }

        # one unique slicer cache per field, linked to every pivot
        $wb.Worksheets.Item($firstPivotSheet).Activate()
        foreach ($field in $slicers) {
            if (-not $caches.ContainsKey($field)) {
                $cache = $wb.SlicerCaches.Add2($pts[0], $field)
                try { $cache.ShowAllItems = $false } catch {}
                $caches[$field] = $cache
            }
            foreach ($p in $pts) {
                try { [void]$caches[$field].PivotTables.AddPivotTable($p) } catch {}
            }
        }

        # dashboard: slicers on top, then one chart per KPI with 4 charts per row
        $wsDash = $wb.Worksheets.Add($missing, $wb.Worksheets.Item($wb.Worksheets.Count))
        $wsDash.Name = $cfg.dashboardSheetName
        Set-NoGrid $wsDash

        # optional summary: per-group comparison of two periods
        $sTop = 4
        if ($cfg.hasSummary -and $cfg.summaryRows -and @($cfg.summaryRows).Count -gt 0) {
            $nKpi = @($cfg.summaryRows).Count
            $groups = @($cfg.summaryGroups)
            $nG = $groups.Count
            $hdr = 15917529
            $sub = 15921906
            $tcell = $wsDash.Cells(1, 1)
            $tcell.Value2 = $cfg.groupField
            $tcell.Font.Bold = $true
            $tcell.Interior.Color = $hdr
            $tcell.HorizontalAlignment = -4108
            for ($g = 0; $g -lt $nG; $g++) {
                $c0 = 2 + $g * 4
                $grng = $wsDash.Range($wsDash.Cells(1, $c0), $wsDash.Cells(1, $c0 + 3))
                $grng.Merge()
                $gc = $wsDash.Cells(1, $c0)
                $gc.Value2 = $groups[$g]
                $gc.Font.Bold = $true
                $gc.Interior.Color = $hdr
                $gc.HorizontalAlignment = -4108
            }
            $hcell = $wsDash.Cells(2, 1)
            $hcell.Value2 = "KPI"
            $hcell.Font.Bold = $true
            $hcell.Interior.Color = $sub
            $hcell.HorizontalAlignment = -4108
            for ($g = 0; $g -lt $nG; $g++) {
                $c0 = 2 + $g * 4
                $labels = @($cfg.periodA, $cfg.periodB, "Diff", "Diff %")
                for ($k = 0; $k -lt 4; $k++) {
                    $c = $wsDash.Cells(2, $c0 + $k)
                    $c.Value2 = $labels[$k]
                    $c.Font.Bold = $true
                    $c.Interior.Color = $sub
                    $c.HorizontalAlignment = -4108
                }
            }
            for ($i = 0; $i -lt $nKpi; $i++) {
                $r = 3 + $i
                $wsDash.Cells($r, 1).Value2 = $cfg.summaryRows[$i].kpi
                for ($g = 0; $g -lt $nG; $g++) {
                    $c0 = 2 + $g * 4
                    $row = $cfg.summaryRows[$i]
                    $a = $row.a[$g]; $b = $row.b[$g]; $d = $row.diff[$g]; $p = $row.pct[$g]
                    $ca = $wsDash.Cells($r, $c0)
                    $cb = $wsDash.Cells($r, $c0 + 1)
                    $cd = $wsDash.Cells($r, $c0 + 2)
                    $cp = $wsDash.Cells($r, $c0 + 3)
                    if ($null -eq $a) { $ca.Value2 = "n/a" } else { $ca.Value2 = [double]$a; $ca.NumberFormat = "0.00" }
                    if ($null -eq $b) { $cb.Value2 = "n/a" } else { $cb.Value2 = [double]$b; $cb.NumberFormat = "0.00" }
                    if ($null -eq $d) { $cd.Value2 = "n/a" } else { $cd.Value2 = [double]$d; $cd.NumberFormat = "0.00" }
                    if ($null -eq $p) { $cp.Value2 = "n/a" } else { $cp.Value2 = [double]$p; $cp.NumberFormat = "0.00" }
                }
            }
            $lastCol = 1 + $nG * 4
            $tbl = $wsDash.Range($wsDash.Cells(1, 1), $wsDash.Cells(2 + $nKpi, $lastCol))
            try { $tbl.Borders.LineStyle = 1; $tbl.Borders.Weight = 2 } catch {}
            try { $tbl.HorizontalAlignment = -4108 } catch {}
            try { $tbl.VerticalAlignment = -4108 } catch {}
            try { $tbl.EntireColumn.AutoFit() } catch {}
            for ($r = 1; $r -le (2 + $nKpi); $r++) { try { $wsDash.Rows($r).RowHeight = 16 } catch {} }
            $sTop = (2 + $nKpi) * 16 + 16
        }

        $sLeft = 4
        $s = 0
        $maxSlicerH = 0
        foreach ($field in $slicers) {
            $n = Slicer-Count $field
            $h = $slicerHBase + ($n * $slicerHItem)
            if ($h -gt $maxSlicerH) { $maxSlicerH = $h }
            $left = $sLeft + $s * ($slicerW + $gap)
            $sn = "Slicer_" + ($field -replace '\W','') + "_d"
            $sl = $caches[$field].Slicers.Add($wsDash, $missing, $sn, $field, $sTop, $left, $slicerW, $h)
            try { $sl.Shape.Placement = 3 } catch {}   # don't move or size with cells
            $s++
        }

        $chartW = 380
        $chartH = 240
        $gapX = 14
        $gapY = 14
        $chartTop0 = $sTop + $maxSlicerH + 16
        for ($i = 0; $i -lt $pts.Count; $i++) {
            $gcol = $i % 4
            $grow = [math]::Floor($i / 4)
            $left = 4 + $gcol * ($chartW + $gapX)
            $top = $chartTop0 + $grow * ($chartH + $gapY)
            $co = $wsDash.ChartObjects().Add($left, $top, $chartW, $chartH)
            [void]$co.Chart.SetSourceData($pts[$i].TableRange1)
            Style-Chart $co $cfg.kpis[$i].col
        }

        # in the one-pivot-per-sheet variant, hide the per-KPI pivot sheets
        if ($cfg.layout -eq "one_pivot") {
            for ($i = 0; $i -lt $cfg.kpis.Count; $i++) {
                try { $wb.Worksheets.Item($cfg.kpis[$i].sheet).Visible = 0 } catch {}
            }
        }
    }

    $wsData.Visible = 0
    $excel.Calculation = -4105   # xlCalculationAutomatic so pivot values/charts refresh
    $excel.CalculateFull()
    $wb.SaveAs($cfg.outPath, 51)
    Write-Output "OK"
}
finally {
    if ($null -ne $wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
