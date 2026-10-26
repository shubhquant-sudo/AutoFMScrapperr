param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = 'Stop'
$workbookPath = (Resolve-Path -LiteralPath $Path).Path
$excelApplication = $null
$excelWorkbook = $null

try {
    $excelApplication = New-Object -ComObject Excel.Application
    $excelApplication.Visible = $false
    $excelApplication.DisplayAlerts = $false
    $excelApplication.EnableEvents = $false
    $excelApplication.AskToUpdateLinks = $false
    $excelApplication.AutomationSecurity = 3  # Disable workbook macros.

    $excelWorkbook = $excelApplication.Workbooks.Open($workbookPath, 0, $false)
    $excelApplication.Calculation = -4105  # xlCalculationAutomatic
    $excelApplication.Iteration = $false
    $excelApplication.CalculateBeforeSave = $true
    $excelApplication.CalculateFullRebuild()

    foreach ($worksheet in $excelWorkbook.Worksheets) {
        $worksheet.Activate()
        $excelApplication.ActiveWindow.Zoom = 85
        $worksheet.PageSetup.Orientation = 2  # xlLandscape
        $worksheet.PageSetup.Zoom = $false
        $worksheet.PageSetup.FitToPagesWide = 1
        $worksheet.PageSetup.FitToPagesTall = $false
        $worksheet.PageSetup.PrintTitleRows = '$1:$7'
        $worksheet.PageSetup.PrintGridlines = $false
    }

    $excelWorkbook.Worksheets.Item('Assumptions').Activate()
    [void]$excelWorkbook.Worksheets.Item('Assumptions').Range('F9').Select()
    $excelWorkbook.Save()
    Write-Output 'Calculated and saved by Microsoft Excel.'
}
finally {
    if ($excelWorkbook) {
        $excelWorkbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelWorkbook)
    }
    if ($excelApplication) {
        $excelApplication.Quit()
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApplication)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
