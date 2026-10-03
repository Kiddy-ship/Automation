[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OrdersJsonPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$xlSheetVisible = -1
$xlSheetHidden = 0
$xlCellTypeConstants = 2
$xlCellTypeFormulas = -4123
$xlErrors = 16
$xlLinkTypeExcelLinks = 1
$xlOpenXMLWorkbook = 51
$msoAutomationSecurityForceDisable = 3

$expectedHeaders = @(
    'Mã hàng',
    'Tên hàng',
    'Đơn giá',
    'Giảm giá',
    'Giảm giá (%)',
    'Số lượng',
    'Lô',
    'Hạn sử dụng',
    'Mã khách hàng',
    'Tên khách hàng',
    'Số điện thoại',
    'Loại khách',
    'Địa chỉ',
    'Khu vực',
    'Phường xã'
)

function Release-ComObject {
    param([AllowNull()][object]$ComObject)

    if ($null -eq $ComObject) {
        return
    }

    try {
        if ([Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
    }
    catch {
        # Cleanup must remain best-effort and must not hide the original error.
    }
}

function Get-JsonProperty {
    param(
        [Parameter(Mandatory = $true)][object]$Object,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Context,
        [switch]$Required
    )

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        if ($Required) {
            throw "$Context thiếu thuộc tính '$Name'."
        }
        return $null
    }

    if ($Required -and $null -eq $property.Value) {
        throw "$Context có thuộc tính '$Name' nhưng giá trị null."
    }

    return $property.Value
}

function Get-RequiredText {
    param(
        [Parameter(Mandatory = $true)][object]$Object,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Context,
        [switch]$PreserveWhitespace
    )

    $value = Get-JsonProperty -Object $Object -Name $Name -Context $Context -Required
    $text = [string]$value
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw "$Context có '$Name' rỗng."
    }

    if ($PreserveWhitespace) {
        return $text
    }

    return $text.Trim()
}

function Convert-ToFiniteDouble {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory = $true)][string]$Field,
        [Parameter(Mandatory = $true)][string]$Context
    )

    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace([string]$Value))) {
        throw "$Context có '$Field' rỗng."
    }

    try {
        $number = [Convert]::ToDouble($Value, [Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        throw "$Context có '$Field' không phải số hợp lệ: '$Value'."
    }

    if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) {
        throw "$Context có '$Field' không phải số hữu hạn."
    }

    return [double]$number
}

function Convert-ToNormalizedOrders {
    param([Parameter(Mandatory = $true)][object]$Payload)

    if ($Payload -is [System.Array]) {
        $rawOrders = @($Payload)
    }
    else {
        $ordersValue = Get-JsonProperty -Object $Payload -Name 'orders' -Context 'JSON gốc' -Required
        $rawOrders = @($ordersValue)
    }

    if ($rawOrders.Count -lt 1) {
        throw 'JSON phải chứa ít nhất 1 đơn hàng.'
    }
    if ($rawOrders.Count -gt 19) {
        throw "Một file chỉ được chứa tối đa 19 đơn hàng; nhận được $($rawOrders.Count)."
    }

    $seenOrderCodes = @{}
    $normalized = New-Object 'System.Collections.Generic.List[object]'

    for ($orderIndex = 0; $orderIndex -lt $rawOrders.Count; $orderIndex++) {
        $rawOrder = $rawOrders[$orderIndex]
        $context = "orders[$orderIndex]"
        if ($null -eq $rawOrder) {
            throw "$context là null."
        }

        $orderCode = Get-RequiredText -Object $rawOrder -Name 'order_code' -Context $context
        $orderKey = $orderCode.ToUpperInvariant()
        if ($seenOrderCodes.ContainsKey($orderKey)) {
            throw "Mã đơn '$orderCode' xuất hiện nhiều lần trong cùng batch."
        }
        $seenOrderCodes[$orderKey] = $true

        $customerName = Get-RequiredText -Object $rawOrder -Name 'customer_name' -Context $context -PreserveWhitespace
        $explicitKiotCode = Get-JsonProperty -Object $rawOrder -Name 'kiot_customer_code' -Context $context
        if ($null -ne $explicitKiotCode -and -not [string]::IsNullOrWhiteSpace([string]$explicitKiotCode)) {
            $kiotCustomerCode = ([string]$explicitKiotCode).Trim()
        }
        else {
            $sourceCustomerCode = Get-RequiredText -Object $rawOrder -Name 'customer_code' -Context $context
            if ($sourceCustomerCode.EndsWith('DN1', [StringComparison]::OrdinalIgnoreCase)) {
                $kiotCustomerCode = $sourceCustomerCode
            }
            else {
                $kiotCustomerCode = $sourceCustomerCode + 'DN1'
            }
        }

        $productsValue = Get-JsonProperty -Object $rawOrder -Name 'products' -Context $context -Required
        $rawProducts = @($productsValue)
        if ($rawProducts.Count -lt 1) {
            throw "$context phải có ít nhất 1 sản phẩm."
        }
        if ($rawProducts.Count -gt 250) {
            throw "Đơn '$orderCode' có $($rawProducts.Count) sản phẩm; KiotViet chỉ cho phép tối đa 250 dòng/sheet."
        }

        $products = New-Object 'System.Collections.Generic.List[object]'
        for ($productIndex = 0; $productIndex -lt $rawProducts.Count; $productIndex++) {
            $rawProduct = $rawProducts[$productIndex]
            $productContext = "$context.products[$productIndex]"
            if ($null -eq $rawProduct) {
                throw "$productContext là null."
            }

            $productCode = Get-RequiredText -Object $rawProduct -Name 'product_code' -Context $productContext
            $productName = Get-RequiredText -Object $rawProduct -Name 'product_name' -Context $productContext -PreserveWhitespace
            $unitPrice = Convert-ToFiniteDouble -Value (Get-JsonProperty -Object $rawProduct -Name 'unit_price' -Context $productContext -Required) -Field 'unit_price' -Context $productContext
            $discountPercent = Convert-ToFiniteDouble -Value (Get-JsonProperty -Object $rawProduct -Name 'discount_percent' -Context $productContext -Required) -Field 'discount_percent' -Context $productContext
            $quantity = Convert-ToFiniteDouble -Value (Get-JsonProperty -Object $rawProduct -Name 'quantity' -Context $productContext -Required) -Field 'quantity' -Context $productContext

            if ($unitPrice -lt 0) {
                throw "$productContext có unit_price âm."
            }
            if ($discountPercent -lt 0 -or $discountPercent -gt 100) {
                throw "$productContext có discount_percent ngoài khoảng 0..100."
            }
            if ($quantity -le 0) {
                throw "$productContext có quantity không lớn hơn 0."
            }

            $products.Add([PSCustomObject]@{
                    product_code       = $productCode
                    product_name       = $productName
                    unit_price         = $unitPrice
                    discount_percent   = $discountPercent
                    quantity           = $quantity
                })
        }

        $normalized.Add([PSCustomObject]@{
                order_code         = $orderCode
                customer_code      = $kiotCustomerCode
                customer_name      = $customerName
                products           = @($products.ToArray())
            })
    }

    return @($normalized.ToArray())
}

function Get-ExcelLinkSources {
    param([Parameter(Mandatory = $true)][object]$Workbook)

    $result = New-Object 'System.Collections.Generic.List[string]'
    $rawLinks = $null
    try {
        $rawLinks = $Workbook.LinkSources($xlLinkTypeExcelLinks)
    }
    catch {
        return @()
    }

    if ($null -eq $rawLinks) {
        return @()
    }

    if ($rawLinks -is [System.Array]) {
        foreach ($link in $rawLinks) {
            if ($null -ne $link -and -not [string]::IsNullOrWhiteSpace([string]$link)) {
                $result.Add([string]$link)
            }
        }
    }
    else {
        $linkText = [string]$rawLinks
        if (-not [string]::IsNullOrWhiteSpace($linkText)) {
            $result.Add($linkText)
        }
    }

    return @($result.ToArray())
}

function Assert-ExpectedHeaders {
    param(
        [Parameter(Mandatory = $true)][object]$Worksheet,
        [Parameter(Mandatory = $true)][string]$Context
    )

    for ($columnIndex = 1; $columnIndex -le $expectedHeaders.Count; $columnIndex++) {
        $cell = $null
        try {
            $cell = $Worksheet.Cells.Item(1, $columnIndex)
            $actual = [string]$cell.Value2
            $expected = $expectedHeaders[$columnIndex - 1]
            if ($actual -cne $expected) {
                throw "$Context sai header cột ${columnIndex}: cần '$expected', nhận '$actual'."
            }
        }
        finally {
            Release-ComObject $cell
        }
    }
}

function Assert-NoCellFormulasOrErrors {
    param([Parameter(Mandatory = $true)][object]$Workbook)

    $worksheets = $null
    try {
        $worksheets = $Workbook.Worksheets
        for ($sheetIndex = 1; $sheetIndex -le $worksheets.Count; $sheetIndex++) {
            $worksheet = $null
            $usedRange = $null
            $formulaCells = $null
            $constantErrors = $null
            $formulaErrors = $null
            try {
                $worksheet = $worksheets.Item($sheetIndex)
                $usedRange = $worksheet.UsedRange

                try {
                    $formulaCells = $usedRange.SpecialCells($xlCellTypeFormulas)
                }
                catch {
                    $formulaCells = $null
                }
                if ($null -ne $formulaCells) {
                    $address = $formulaCells.Address($false, $false)
                    throw "Sheet '$($worksheet.Name)' còn công thức tại $address."
                }

                try {
                    $constantErrors = $usedRange.SpecialCells($xlCellTypeConstants, $xlErrors)
                }
                catch {
                    $constantErrors = $null
                }
                if ($null -ne $constantErrors) {
                    $address = $constantErrors.Address($false, $false)
                    throw "Sheet '$($worksheet.Name)' chứa giá trị lỗi Excel tại $address."
                }

                try {
                    $formulaErrors = $usedRange.SpecialCells($xlCellTypeFormulas, $xlErrors)
                }
                catch {
                    $formulaErrors = $null
                }
                if ($null -ne $formulaErrors) {
                    $address = $formulaErrors.Address($false, $false)
                    throw "Sheet '$($worksheet.Name)' chứa lỗi công thức Excel tại $address."
                }
            }
            finally {
                Release-ComObject $formulaErrors
                Release-ComObject $constantErrors
                Release-ComObject $formulaCells
                Release-ComObject $usedRange
                Release-ComObject $worksheet
            }
        }
    }
    finally {
        Release-ComObject $worksheets
    }
}

function Test-BlankCell {
    param([AllowNull()][object]$Value)
    return ($null -eq $Value -or [string]$Value -eq '')
}

function Assert-NumericEqual {
    param(
        [AllowNull()][object]$Actual,
        [double]$Expected,
        [string]$Context
    )

    if ($null -eq $Actual) {
        throw "$Context bị rỗng; cần $Expected."
    }
    $actualNumber = [Convert]::ToDouble($Actual, [Globalization.CultureInfo]::InvariantCulture)
    if ([Math]::Abs($actualNumber - $Expected) -gt 0.0000001) {
        throw "$Context sai giá trị: cần $Expected, nhận $actualNumber."
    }
}

function Assert-OutputWorkbook {
    param(
        [Parameter(Mandatory = $true)][object]$Workbook,
        [Parameter(Mandatory = $true)][object[]]$Orders,
        [Parameter(Mandatory = $true)][string]$ExpectedTableStyle
    )

    $links = @(Get-ExcelLinkSources -Workbook $Workbook)
    if ($links.Count -gt 0) {
        throw "Workbook output còn external link: $($links -join '; ')."
    }

    Assert-NoCellFormulasOrErrors -Workbook $Workbook

    $worksheets = $null
    $level1 = $null
    $level2 = $null
    try {
        $worksheets = $Workbook.Worksheets
        $level1 = $worksheets.Item('Level1')
        $level2 = $worksheets.Item('Level2')
        if ($level1.Visible -eq $xlSheetVisible -or $level2.Visible -eq $xlSheetVisible) {
            throw 'Level1 và Level2 phải được giữ ở trạng thái hidden.'
        }

        $visibleCount = 0
        for ($sheetIndex = 1; $sheetIndex -le $worksheets.Count; $sheetIndex++) {
            $candidate = $null
            try {
                $candidate = $worksheets.Item($sheetIndex)
                if ($candidate.Visible -eq $xlSheetVisible) {
                    $visibleCount++
                }
            }
            finally {
                Release-ComObject $candidate
            }
        }
        if ($visibleCount -ne $Orders.Count) {
            throw "Workbook output có $visibleCount sheet visible; cần $($Orders.Count)."
        }

        for ($orderIndex = 0; $orderIndex -lt $Orders.Count; $orderIndex++) {
            $order = $Orders[$orderIndex]
            $visibleIndex = $orderIndex + 1
            $expectedSheetName = if ($visibleIndex -eq 1) { 'ImportProductTemplate' } else { "ImportProductTemplate ($visibleIndex)" }
            $expectedTableName = 'ImportProductTemplate_{0:D2}' -f $visibleIndex
            $productCount = $order.products.Count
            $lastRow = $productCount + 1
            $expectedTableAddress = "A1:O$lastRow"

            $worksheet = $null
            $listObjects = $null
            $table = $null
            $tableRange = $null
            try {
                $worksheet = $worksheets.Item($expectedSheetName)
                if ($worksheet.Visible -ne $xlSheetVisible) {
                    throw "Sheet '$expectedSheetName' không visible."
                }
                Assert-ExpectedHeaders -Worksheet $worksheet -Context "Sheet '$expectedSheetName'"

                $listObjects = $worksheet.ListObjects
                if ($listObjects.Count -ne 1) {
                    throw "Sheet '$expectedSheetName' phải có đúng 1 ListObject; nhận $($listObjects.Count)."
                }
                $table = $listObjects.Item(1)
                if ([string]$table.Name -cne $expectedTableName) {
                    throw "Sheet '$expectedSheetName' sai table name: '$($table.Name)'."
                }
                $tableRange = $table.Range
                $actualAddress = $tableRange.Address($false, $false)
                if ($actualAddress -cne $expectedTableAddress) {
                    throw "Sheet '$expectedSheetName' sai table ref: cần $expectedTableAddress, nhận $actualAddress."
                }
                if ([string]$table.TableStyle.Name -cne $ExpectedTableStyle) {
                    throw "Sheet '$expectedSheetName' không giữ table style '$ExpectedTableStyle'."
                }

                for ($productIndex = 0; $productIndex -lt $productCount; $productIndex++) {
                    $row = $productIndex + 2
                    $product = $order.products[$productIndex]
                    $actualProductCode = [string]$worksheet.Cells.Item($row, 1).Value2
                    $actualProductName = [string]$worksheet.Cells.Item($row, 2).Value2
                    $actualCustomerCode = [string]$worksheet.Cells.Item($row, 9).Value2
                    $actualCustomerName = [string]$worksheet.Cells.Item($row, 10).Value2

                    if ($actualProductCode -cne $product.product_code) {
                        throw "Sheet '$expectedSheetName' A$row sai mã hàng."
                    }
                    if ($actualProductName -cne $product.product_name) {
                        throw "Sheet '$expectedSheetName' B$row sai tên hàng."
                    }
                    if ($actualCustomerCode -cne $order.customer_code) {
                        throw "Sheet '$expectedSheetName' I$row sai mã khách hàng."
                    }
                    if ($actualCustomerName -cne $order.customer_name) {
                        throw "Sheet '$expectedSheetName' J$row sai tên khách hàng."
                    }

                    Assert-NumericEqual -Actual $worksheet.Cells.Item($row, 3).Value2 -Expected $product.unit_price -Context "Sheet '$expectedSheetName' C$row"
                    Assert-NumericEqual -Actual $worksheet.Cells.Item($row, 5).Value2 -Expected $product.discount_percent -Context "Sheet '$expectedSheetName' E$row"
                    Assert-NumericEqual -Actual $worksheet.Cells.Item($row, 6).Value2 -Expected $product.quantity -Context "Sheet '$expectedSheetName' F$row"

                    foreach ($blankColumn in @(4, 7, 8, 11, 12, 13, 14, 15)) {
                        $blankValue = $worksheet.Cells.Item($row, $blankColumn).Value2
                        if (-not (Test-BlankCell -Value $blankValue)) {
                            throw "Sheet '$expectedSheetName' ô hàng $row cột $blankColumn phải trống."
                        }
                    }
                }

                foreach ($validationRow in @(2, $lastRow) | Select-Object -Unique) {
                    $nCell = $null
                    $oCell = $null
                    try {
                        $nCell = $worksheet.Cells.Item($validationRow, 14)
                        $oCell = $worksheet.Cells.Item($validationRow, 15)
                        if ($nCell.Validation.Type -ne 3 -or $oCell.Validation.Type -ne 3) {
                            throw "Sheet '$expectedSheetName' thiếu list validation tại N/O hàng $validationRow."
                        }
                        $dependentFormula = [string]$oCell.Validation.Formula1
                        if ($dependentFormula.IndexOf("N$validationRow", [StringComparison]::OrdinalIgnoreCase) -lt 0) {
                            throw "Sheet '$expectedSheetName' có validation O$validationRow không tham chiếu N$validationRow."
                        }
                    }
                    finally {
                        Release-ComObject $oCell
                        Release-ComObject $nCell
                    }
                }
            }
            finally {
                Release-ComObject $tableRange
                Release-ComObject $table
                Release-ComObject $listObjects
                Release-ComObject $worksheet
            }
        }
    }
    finally {
        Release-ComObject $level2
        Release-ComObject $level1
        Release-ComObject $worksheets
    }
}

$templateFullPath = $null
$ordersFullPath = $null
$outputFullPath = $null
$stagingPath = $null
$orders = $null
$excel = $null
$buildWorkbook = $null
$validationWorkbook = $null
$buildSheetNames = @()
$mappings = @()
$tableStyleName = $null
$scriptError = $null
$buildValidated = $false

try {
    $templateFullPath = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $TemplatePath).Path)
    $ordersFullPath = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $OrdersJsonPath).Path)
    $outputFullPath = [IO.Path]::GetFullPath($OutputPath)

    if ([IO.Path]::GetExtension($templateFullPath) -ine '.xlsx') {
        throw 'TemplatePath phải là file .xlsx.'
    }
    if ([IO.Path]::GetExtension($outputFullPath) -ine '.xlsx') {
        throw 'OutputPath phải có đuôi .xlsx.'
    }
    if ($templateFullPath -ieq $outputFullPath) {
        throw 'OutputPath không được trùng TemplatePath.'
    }

    $jsonText = [IO.File]::ReadAllText($ordersFullPath, [Text.Encoding]::UTF8)
    if ([string]::IsNullOrWhiteSpace($jsonText)) {
        throw 'OrdersJsonPath là file rỗng.'
    }
    try {
        $payload = $jsonText | ConvertFrom-Json
    }
    catch {
        throw "JSON không hợp lệ: $($_.Exception.Message)"
    }
    $orders = @(Convert-ToNormalizedOrders -Payload $payload)

    $outputDirectory = [IO.Path]::GetDirectoryName($outputFullPath)
    if ([string]::IsNullOrWhiteSpace($outputDirectory)) {
        throw 'Không xác định được thư mục OutputPath.'
    }
    [void][IO.Directory]::CreateDirectory($outputDirectory)
    # OutputPath is already a short unique partial path supplied by the Python
    # wrapper.  Appending another full GUID can push Excel SaveAs beyond its
    # practical Windows path limit.
    $stagingName = '{0}.staging.xlsx' -f [IO.Path]::GetFileNameWithoutExtension($outputFullPath)
    $stagingPath = [IO.Path]::Combine($outputDirectory, $stagingName)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable

    $workbooks = $null
    try {
        $workbooks = $excel.Workbooks
        $buildWorkbook = $workbooks.Open($templateFullPath, 0, $true)
    }
    finally {
        Release-ComObject $workbooks
    }

    $worksheets = $null
    $level1 = $null
    $level2 = $null
    $templateSheet = $null
    $templateTable = $null
    $templateTableRange = $null
    try {
        $worksheets = $buildWorkbook.Worksheets
        $level1 = $worksheets.Item('Level1')
        $level2 = $worksheets.Item('Level2')
        $level1.Visible = $xlSheetHidden
        $level2.Visible = $xlSheetHidden

        $visibleTemplateCount = 0
        for ($sheetIndex = 1; $sheetIndex -le $worksheets.Count; $sheetIndex++) {
            $candidate = $null
            try {
                $candidate = $worksheets.Item($sheetIndex)
                if ($candidate.Visible -eq $xlSheetVisible -and $candidate.Name -cne 'Level1' -and $candidate.Name -cne 'Level2') {
                    $visibleTemplateCount++
                    if ($visibleTemplateCount -eq 1) {
                        $templateSheet = $candidate
                        $candidate = $null
                    }
                }
            }
            finally {
                Release-ComObject $candidate
            }
        }
        if ($visibleTemplateCount -ne 1 -or $null -eq $templateSheet) {
            throw "Template sạch phải có đúng 1 sheet visible ngoài Level1/Level2; nhận $visibleTemplateCount."
        }

        Assert-ExpectedHeaders -Worksheet $templateSheet -Context 'Template'
        $listObjects = $null
        try {
            $listObjects = $templateSheet.ListObjects
            if ($listObjects.Count -ne 1) {
                throw "Template sheet phải có đúng 1 ListObject; nhận $($listObjects.Count)."
            }
            $templateTable = $listObjects.Item(1)
        }
        finally {
            Release-ComObject $listObjects
        }

        $templateTableRange = $templateTable.Range
        if ($templateTableRange.Columns.Count -ne 15 -or $templateTableRange.Row -ne 1 -or $templateTableRange.Column -ne 1) {
            throw "Template table phải bắt đầu tại A1 và có 15 cột; nhận $($templateTableRange.Address($false, $false))."
        }
        $tableStyleName = [string]$templateTable.TableStyle.Name

        $usedRange = $null
        $clearRange = $null
        $resizeRange = $null
        try {
            $usedRange = $templateSheet.UsedRange
            $usedLastRow = $usedRange.Row + $usedRange.Rows.Count - 1
            $tableLastRow = $templateTableRange.Row + $templateTableRange.Rows.Count - 1
            $clearLastRow = [Math]::Max(2, [Math]::Max($usedLastRow, $tableLastRow))
            $clearRange = $templateSheet.Range("A2:O$clearLastRow")
            [void]$clearRange.ClearContents()
            $resizeRange = $templateSheet.Range('A1:O2')
            [void]$templateTable.Resize($resizeRange)
        }
        finally {
            Release-ComObject $resizeRange
            Release-ComObject $clearRange
            Release-ComObject $usedRange
        }

        $n2 = $null
        $o2 = $null
        try {
            $n2 = $templateSheet.Range('N2')
            $o2 = $templateSheet.Range('O2')
            if ($n2.Validation.Type -ne 3 -or $o2.Validation.Type -ne 3) {
                throw 'Template thiếu list validation tại N2/O2.'
            }
        }
        finally {
            Release-ComObject $o2
            Release-ComObject $n2
        }

        $templateSheet.Name = 'ImportProductTemplate'
        $buildSheetNames += 'ImportProductTemplate'

        # Copying a worksheet can invalidate the PowerShell RCW that was used
        # as the source.  Keep stable sheet names and reacquire a fresh COM
        # wrapper for every copy/write operation instead of retaining RCWs in
        # an array across Worksheet.Copy calls.
        Release-ComObject $templateSheet
        $templateSheet = $null

        for ($visibleIndex = 2; $visibleIndex -le $orders.Count; $visibleIndex++) {
            $copySource = $null
            $afterSheet = $null
            $newSheet = $null
            try {
                $copySource = $worksheets.Item('ImportProductTemplate')
                $afterSheet = $worksheets.Item($worksheets.Count)
                [void]$copySource.Copy([Type]::Missing, $afterSheet)
                $newSheet = $worksheets.Item($worksheets.Count)
                $newSheet.Name = "ImportProductTemplate ($visibleIndex)"
                $buildSheetNames += [string]$newSheet.Name
            }
            finally {
                Release-ComObject $newSheet
                Release-ComObject $afterSheet
                Release-ComObject $copySource
            }
        }

        for ($orderIndex = 0; $orderIndex -lt $orders.Count; $orderIndex++) {
            $order = $orders[$orderIndex]
            $visibleIndex = $orderIndex + 1
            $worksheet = $null
            $productCount = $order.products.Count
            $lastRow = $productCount + 1
            $tableName = 'ImportProductTemplate_{0:D2}' -f $visibleIndex

            $listObjects = $null
            $table = $null
            $newTableRange = $null
            $styleSource = $null
            $styleDestination = $null
            $dataRange = $null
            try {
                $worksheet = $worksheets.Item($buildSheetNames[$orderIndex])
                $listObjects = $worksheet.ListObjects
                if ($listObjects.Count -ne 1) {
                    throw "Sheet '$($worksheet.Name)' phải có đúng 1 ListObject."
                }
                $table = $listObjects.Item(1)
                $table.Name = $tableName
                $newTableRange = $worksheet.Range("A1:O$lastRow")
                [void]$table.Resize($newTableRange)

                $styleSource = $worksheet.Range('A2:O2')
                $styleDestination = $worksheet.Range("A2:O$lastRow")
                if ($productCount -gt 1) {
                    [void]$styleSource.Copy($styleDestination)
                    # Excel's typed COM wrapper does not accept Boolean false
                    # for CutCopyMode.  CANCEL.COPY clears the clipboard state
                    # without SendKeys and prevents SaveAs from failing after
                    # copying template formatting across multiple rows.
                    [void]$excel.ExecuteExcel4Macro('CANCEL.COPY()')
                }
                [void]$styleDestination.ClearContents()

                $lengths = [int[]]@($productCount, 15)
                $matrix = [Array]::CreateInstance([object], $lengths)
                for ($productIndex = 0; $productIndex -lt $productCount; $productIndex++) {
                    $product = $order.products[$productIndex]
                    $matrix[$productIndex, 0] = $product.product_code
                    $matrix[$productIndex, 1] = $product.product_name
                    $matrix[$productIndex, 2] = [double]$product.unit_price
                    $matrix[$productIndex, 3] = $null
                    $matrix[$productIndex, 4] = [double]$product.discount_percent
                    $matrix[$productIndex, 5] = [double]$product.quantity
                    $matrix[$productIndex, 6] = $null
                    $matrix[$productIndex, 7] = $null
                    $matrix[$productIndex, 8] = $order.customer_code
                    $matrix[$productIndex, 9] = $order.customer_name
                    $matrix[$productIndex, 10] = $null
                    $matrix[$productIndex, 11] = $null
                    $matrix[$productIndex, 12] = $null
                    $matrix[$productIndex, 13] = $null
                    $matrix[$productIndex, 14] = $null
                }

                $dataRange = $worksheet.Range("A2:O$lastRow")
                $dataRange.Value2 = $matrix

                $mappings += [PSCustomObject]@{
                    sheet_index  = $visibleIndex
                    sheet_name   = [string]$worksheet.Name
                    order_code   = $order.order_code
                    product_count = $productCount
                }
            }
            finally {
                Release-ComObject $dataRange
                Release-ComObject $styleDestination
                Release-ComObject $styleSource
                Release-ComObject $newTableRange
                Release-ComObject $table
                Release-ComObject $listObjects
                Release-ComObject $worksheet
            }
        }

        $preSaveLinks = @(Get-ExcelLinkSources -Workbook $buildWorkbook)
        foreach ($link in $preSaveLinks) {
            [void]$buildWorkbook.BreakLink($link, $xlLinkTypeExcelLinks)
        }

        Assert-NoCellFormulasOrErrors -Workbook $buildWorkbook
        [void]$excel.ExecuteExcel4Macro('CANCEL.COPY()')
        $buildWorkbook.CheckCompatibility = $false
        [void]$buildWorkbook.SaveAs($stagingPath, $xlOpenXMLWorkbook)
        [void]$buildWorkbook.Save()
    }
    finally {
        Release-ComObject $templateTableRange
        Release-ComObject $templateTable
        Release-ComObject $level2
        Release-ComObject $level1
        Release-ComObject $worksheets
    }

    if ($null -ne $buildWorkbook) {
        [void]$buildWorkbook.Close($false)
        $buildSheetNames = @()
        Release-ComObject $buildWorkbook
        $buildWorkbook = $null
    }

    $workbooks = $null
    try {
        $workbooks = $excel.Workbooks
        $validationWorkbook = $workbooks.Open($stagingPath, 0, $true)
    }
    finally {
        Release-ComObject $workbooks
    }

    Assert-OutputWorkbook -Workbook $validationWorkbook -Orders $orders -ExpectedTableStyle $tableStyleName
    $buildValidated = $true
}
catch {
    $scriptError = $_
}
finally {
    if ($null -ne $validationWorkbook) {
        try { [void]$validationWorkbook.Close($false) } catch {}
        Release-ComObject $validationWorkbook
        $validationWorkbook = $null
    }

    if ($null -ne $buildWorkbook) {
        try { [void]$buildWorkbook.Close($false) } catch {}
        Release-ComObject $buildWorkbook
        $buildWorkbook = $null
    }

    $buildSheetNames = @()

    if ($null -ne $excel) {
        try {
            $excel.DisplayAlerts = $false
            [void]$excel.Quit()
        }
        catch {}
        Release-ComObject $excel
        $excel = $null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if (-not $buildValidated -and $null -ne $stagingPath -and [IO.File]::Exists($stagingPath)) {
        try { [IO.File]::Delete($stagingPath) } catch {}
    }
}

if ($null -ne $scriptError) {
    $errorPayload = [ordered]@{
        success = $false
        error   = [string]$scriptError.Exception.Message
        line    = [int]$scriptError.InvocationInfo.ScriptLineNumber
        stack   = [string]$scriptError.ScriptStackTrace
    }
    [Console]::Error.WriteLine(($errorPayload | ConvertTo-Json -Compress -Depth 4))
    exit 1
}

try {
    if ([IO.File]::Exists($outputFullPath)) {
        [IO.File]::Replace($stagingPath, $outputFullPath, $null, $true)
    }
    else {
        [IO.File]::Move($stagingPath, $outputFullPath)
    }
}
catch {
    if ($null -ne $stagingPath -and [IO.File]::Exists($stagingPath)) {
        try { [IO.File]::Delete($stagingPath) } catch {}
    }
    $errorPayload = [ordered]@{
        success = $false
        error   = "Không thể hoàn tất OutputPath: $($_.Exception.Message)"
    }
    [Console]::Error.WriteLine(($errorPayload | ConvertTo-Json -Compress -Depth 4))
    exit 1
}

$outputInfo = Get-Item -LiteralPath $outputFullPath
$resultPayload = [ordered]@{
    success             = $true
    output_path         = $outputInfo.FullName
    order_count         = $orders.Count
    visible_sheet_count = $orders.Count
    mappings            = @($mappings)
    file_size           = [long]$outputInfo.Length
}

Write-Output ($resultPayload | ConvertTo-Json -Compress -Depth 6)
