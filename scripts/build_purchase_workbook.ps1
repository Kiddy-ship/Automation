[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$TemplatePath,
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [string]$BaselinePath = '',
    [double]$VatDivisor = 1.08
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

function Release-ComObject([AllowNull()][object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}

if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) { throw "Template file not found: $TemplatePath" }
if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { throw "Source file not found: $SourcePath" }
if (-not [string]::IsNullOrWhiteSpace($BaselinePath) -and -not (Test-Path -LiteralPath $BaselinePath -PathType Leaf)) { throw "Baseline file not found: $BaselinePath" }
if ($VatDivisor -le 0) { throw 'VatDivisor must be greater than zero.' }

$excel = $null; $template = $null; $source = $null; $baseline = $null
$templateSheet = $null; $sourceSheet = $null; $baselineSheet = $null
$sourceUsed = $null; $templateUsed = $null; $baselineUsed = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $template = $excel.Workbooks.Open((Resolve-Path -LiteralPath $TemplatePath).Path, $false, $false)
    $source = $excel.Workbooks.Open((Resolve-Path -LiteralPath $SourcePath).Path, $false, $true)
    $templateSheet = $template.Worksheets.Item(1)
    $sourceSheet = $source.Worksheets.Item(1)
    $sourceUsed = $sourceSheet.UsedRange
    $templateUsed = $templateSheet.UsedRange

    # The source shipment sheet is fixed: B=SKU, C=name, F=delivered quantity,
    # I=buying unit price including VAT. J is the line total (I * F), not the
    # unit price and must never be used for the import price. Rows with
    # blank/dash/non-positive F are
    # deliberately excluded.
    $lastSourceRow = [int]($sourceUsed.Row + $sourceUsed.Rows.Count - 1)
    $items = New-Object 'System.Collections.Generic.List[object]'
    for ($row = 11; $row -le $lastSourceRow; $row++) {
        $sku = [string]$sourceSheet.Cells.Item($row, 2).Value2
        $name = [string]$sourceSheet.Cells.Item($row, 3).Value2
        $quantityRaw = $sourceSheet.Cells.Item($row, 6).Value2
        $priceRaw = $sourceSheet.Cells.Item($row, 9).Value2
        if ([string]::IsNullOrWhiteSpace($sku)) { continue }
        if ($null -eq $quantityRaw -or [string]::IsNullOrWhiteSpace([string]$quantityRaw) -or [string]$quantityRaw -eq '-') { continue }
        try { $quantity = [double]$quantityRaw } catch { continue }
        if ($quantity -le 0) { continue }
        try { $unitPriceIncludingVat = [double]$priceRaw } catch { continue }
        if ($unitPriceIncludingVat -lt 0) { continue }
        $items.Add([PSCustomObject]@{
            source_row = $row
            product_code = $sku.Trim()
            product_name = $name.Trim()
            unit_price = [math]::Round($unitPriceIncludingVat / $VatDivisor, 6)
            quantity = $quantity
        })
    }
    $sourcePositiveItemCount = $items.Count
    $comparison = $null
    if (-not [string]::IsNullOrWhiteSpace($BaselinePath)) {
        # The 12.09-style correction files are cumulative shipment files: a
        # SKU may show today's total quantity, not only the newly added
        # quantity.  Aggregate positive delivered quantities per SKU from the
        # baseline and source, then write only a strictly positive difference.
        # Do not guess when one source SKU has conflicting names/prices.
        $baseline = $excel.Workbooks.Open((Resolve-Path -LiteralPath $BaselinePath).Path, $false, $true)
        $baselineSheet = $baseline.Worksheets.Item(1)
        $baselineUsed = $baselineSheet.UsedRange
        $baselineLastRow = [int]($baselineUsed.Row + $baselineUsed.Rows.Count - 1)
        $baselineTotals = @{}
        for ($row = 11; $row -le $baselineLastRow; $row++) {
            $sku = [string]$baselineSheet.Cells.Item($row, 2).Value2
            $quantityRaw = $baselineSheet.Cells.Item($row, 6).Value2
            if ([string]::IsNullOrWhiteSpace($sku)) { continue }
            if ($null -eq $quantityRaw -or [string]::IsNullOrWhiteSpace([string]$quantityRaw) -or [string]$quantityRaw -eq '-') { continue }
            try { $quantity = [double]$quantityRaw } catch { continue }
            if ($quantity -le 0) { continue }
            $key = $sku.Trim().ToUpperInvariant()
            if (-not $baselineTotals.ContainsKey($key)) { $baselineTotals[$key] = 0.0 }
            $baselineTotals[$key] = [double]$baselineTotals[$key] + $quantity
        }

        $currentGroups = @{}
        foreach ($item in $items) {
            $key = ([string]$item.product_code).Trim().ToUpperInvariant()
            if (-not $currentGroups.ContainsKey($key)) {
                $currentGroups[$key] = [PSCustomObject]@{
                    product_code = [string]$item.product_code
                    product_name = [string]$item.product_name
                    unit_price = [double]$item.unit_price
                    quantity = [double]$item.quantity
                    source_row = [int]$item.source_row
                }
                continue
            }
            $existing = $currentGroups[$key]
            if ($existing.product_name -ne [string]$item.product_name) {
                throw "SKU $($existing.product_code) has conflicting product names in the source; cannot calculate a safe incremental receipt."
            }
            if ([math]::Abs(([double]$existing.unit_price) - ([double]$item.unit_price)) -gt 0.000001) {
                throw "SKU $($existing.product_code) has conflicting purchase prices in the source; cannot calculate a safe incremental receipt."
            }
            $existing.quantity = [double]$existing.quantity + [double]$item.quantity
        }

        $incrementalItems = New-Object 'System.Collections.Generic.List[object]'
        $newSkus = New-Object 'System.Collections.Generic.List[string]'
        $increasedSkus = New-Object 'System.Collections.Generic.List[string]'
        $unchangedOrDecreasedSkus = New-Object 'System.Collections.Generic.List[string]'
        $decreasedSkus = New-Object 'System.Collections.Generic.List[string]'
        foreach ($group in $currentGroups.Values) {
            $key = ([string]$group.product_code).Trim().ToUpperInvariant()
            $baselineQuantity = if ($baselineTotals.ContainsKey($key)) { [double]$baselineTotals[$key] } else { 0.0 }
            $deltaQuantity = [double]$group.quantity - $baselineQuantity
            if ($deltaQuantity -gt 0.000001) {
                $incrementalItems.Add([PSCustomObject]@{
                    source_row = [int]$group.source_row
                    product_code = [string]$group.product_code
                    product_name = [string]$group.product_name
                    unit_price = [double]$group.unit_price
                    quantity = [math]::Round($deltaQuantity, 6)
                })
                if ($baselineTotals.ContainsKey($key)) {
                    $increasedSkus.Add([string]$group.product_code)
                } else {
                    $newSkus.Add([string]$group.product_code)
                }
            } else {
                $unchangedOrDecreasedSkus.Add([string]$group.product_code)
                if ($deltaQuantity -lt -0.000001) {
                    $decreasedSkus.Add([string]$group.product_code)
                }
            }
        }
        $items = $incrementalItems
        $comparison = [ordered]@{
            mode = 'positive_quantity_delta'
            baseline_path = (Resolve-Path -LiteralPath $BaselinePath).Path
            baseline_sku_count = $baselineTotals.Count
            source_sku_count = $currentGroups.Count
            incremental_sku_count = $incrementalItems.Count
            new_skus = @($newSkus.ToArray())
            increased_skus = @($increasedSkus.ToArray())
            unchanged_or_decreased_skus = @($unchangedOrDecreasedSkus.ToArray())
            decreased_skus = @($decreasedSkus.ToArray())
        }
    }
    if ($items.Count -lt 1 -and [string]::IsNullOrWhiteSpace($BaselinePath)) {
        throw 'Source has no rows with a positive delivered quantity.'
    }

    # Preserve the sample workbook formatting, but remove sample values below
    # the header before writing the filtered source rows.
    $templateLastRow = [int]($templateUsed.Row + $templateUsed.Rows.Count - 1)
    if ($templateLastRow -ge 2) {
        $templateSheet.Range("A2:N$templateLastRow").ClearContents() | Out-Null
    }
    $outRow = 2
    foreach ($item in $items) {
        $templateSheet.Cells.Item($outRow, 1).Value2 = $item.product_code
        $templateSheet.Cells.Item($outRow, 2).Value2 = $item.product_name
        # Unit, discounts, VAT and lot columns intentionally retain the empty
        # template defaults.  KiotViet VAT is enabled after upload.
        # Write numeric values, never locale-dependent text. A decimal point
        # stored as text can be interpreted as a thousands separator by the
        # KiotViet importer (for example 5,555.555 -> 5,555,555).
        $templateSheet.Cells.Item($outRow, 4).Value2 = [double]$item.unit_price
        $templateSheet.Cells.Item($outRow, 7).Value2 = [double]$item.quantity
        $outRow++
    }
    if ($outRow -gt 2) {
        $templateSheet.Range("D2:D$($outRow - 1)").NumberFormat = '#,##0.######'
        $templateSheet.Range("G2:G$($outRow - 1)").NumberFormat = '#,##0.######'
    }
    $output = [IO.Path]::GetFullPath($OutputPath)
    $parent = Split-Path -Parent $output
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
    if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output -Force }
    $template.SaveAs($output, 51)
    $itemArray = @($items.ToArray())
    $metadata = [ordered]@{
        success = $true
        output_path = $output
        item_count = $itemArray.Count
        source_positive_item_count = $sourcePositiveItemCount
        skipped_non_delivery = ($lastSourceRow - 10 - $sourcePositiveItemCount)
        vat_divisor = $VatDivisor
        items = $itemArray
        comparison = $comparison
    }
    $json = ConvertTo-Json -Compress -Depth 6 -InputObject $metadata
    Write-Output $json
}
finally {
    if ($null -ne $baseline) { $baseline.Close($false) }
    if ($null -ne $source) { $source.Close($false) }
    if ($null -ne $template) { $template.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    Release-ComObject $baselineUsed; Release-ComObject $sourceUsed; Release-ComObject $templateUsed
    Release-ComObject $baselineSheet; Release-ComObject $sourceSheet; Release-ComObject $templateSheet
    Release-ComObject $baseline; Release-ComObject $source; Release-ComObject $template; Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
