<# :
@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_FILE=%~f0"
cls
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "iex (Get-Content -LiteralPath $env:SCRIPT_FILE -Raw)"
set "exitCode=%ERRORLEVEL%"
pause
exit /b %exitCode%
#>
# Convert Excel workbooks beside this script to ACIC fixed-width text files beside this script.
# Requires Windows PowerShell 5.0 or later and Microsoft Excel installed.

$ErrorActionPreference = 'Stop'
$scriptDirectory = $env:SCRIPT_DIR
$inputDirectory = $scriptDirectory
$outputDirectory = $scriptDirectory
$invariant = [Globalization.CultureInfo]::InvariantCulture
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$requiredFields = @('AccountNumber', 'CheckNumber', 'CheckAmount', 'Payee', 'CheckDate')

function ConvertTo-PaddedDigits([object]$Value, [int]$Width, [string]$FieldName) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        throw "$FieldName is empty."
    }

    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
        $digits = ([decimal]$Value).ToString('0', $invariant)
    }
    else {
        $text = ([string]$Value).Trim()
        if ($text -notmatch '^[0-9\s-]+$') {
            throw "$FieldName must contain digits only (hyphens and spaces are allowed); got '$Value'."
        }
        $digits = [regex]::Replace($text, '\D', '')
    }

    if ($digits.Length -eq 0 -or $digits.Length -gt $Width) {
        throw "$FieldName must contain at most $Width digits; got '$Value'."
    }
    return $digits.PadLeft($Width, '0')
}

function ConvertTo-CheckDate([object]$Value) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        throw 'CheckDate is empty.'
    }

    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
        try { return [datetime]::FromOADate([double]$Value).ToString('MMddyyyy', $invariant) }
        catch { throw "Invalid Excel date serial for CheckDate: '$Value'." }
    }

    $parsed = [datetime]::MinValue
    $text = ([string]$Value).Trim()
    if ([datetime]::TryParse($text, $invariant, [Globalization.DateTimeStyles]::None, [ref]$parsed) -or
        [datetime]::TryParse($text, [Globalization.CultureInfo]::CurrentCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.ToString('MMddyyyy', $invariant)
    }
    throw "Invalid CheckDate: '$Value'."
}

function ConvertTo-CheckAmountCents([object]$Value) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        throw 'CheckAmount is empty.'
    }

    $amount = [decimal]0
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
        $amount = [decimal]$Value
    }
    else {
        $text = ([string]$Value).Trim() -replace '[,$\s]', ''
        if (-not [decimal]::TryParse($text, [Globalization.NumberStyles]::Number, $invariant, [ref]$amount)) {
            throw "Invalid CheckAmount: '$Value'."
        }
    }

    $cents = [decimal]::Round(($amount * 100), 0, [MidpointRounding]::AwayFromZero)
    if ($cents -lt 0 -or $cents -gt 9999999999999999) {
        throw "CheckAmount must fit the 16-digit amount field; got '$Value'."
    }
    return $cents
}

if (-not (Test-Path -LiteralPath $inputDirectory -PathType Container)) {
    throw "Excel input folder not found: $inputDirectory"
}
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

$workbooks = @(Get-ChildItem -LiteralPath $inputDirectory -File | Where-Object {
    $_.Extension -match '^\.xlsx?$' -and $_.Name -notlike '~$*'
})
if ($workbooks.Count -eq 0) {
    Write-Warning "No .xls or .xlsx files found in $inputDirectory"
    exit 0
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($inputFile in $workbooks) {
        $workbook = $null
        $didConvert = $false
        try {
            $workbook = $excel.Workbooks.Open($inputFile.FullName, 0, $true)
            for ($sheetIndex = 1; $sheetIndex -le $workbook.Worksheets.Count; $sheetIndex++) {
                $worksheet = $null
                $usedRange = $null
                try {
                    $worksheet = $workbook.Worksheets.Item($sheetIndex)
                    $usedRange = $worksheet.UsedRange
                    $rowCount = [int]$usedRange.Rows.Count
                    $columnCount = [int]$usedRange.Columns.Count
                    if ($rowCount -lt 2 -or $columnCount -lt 1) { continue }

                    $cells = $usedRange.Value2
                    $fieldColumns = @{}
                    for ($column = 1; $column -le $columnCount; $column++) {
                        $header = ([string]$cells[1, $column]).Trim().TrimStart([char]0xFEFF)
                        if ($header) { $fieldColumns[$header] = $column }
                    }
                    $missingFields = @($requiredFields | Where-Object { -not $fieldColumns.ContainsKey($_) })
                    if ($missingFields.Count -gt 0) { continue }

                    $lines = New-Object 'System.Collections.Generic.List[string]'
                    $calculationTotal = [decimal]0
                    $amountTotalCents = [decimal]0
                    $conversionTime = Get-Date
                    $transactionDate = $conversionTime.ToString('MMddyyyy', $invariant)
                    $transactionTime = $conversionTime.ToString('HHmmss', $invariant)
                    $dataCount = 0

                    for ($row = 2; $row -le $rowCount; $row++) {
                        $accountValue = $cells[$row, $fieldColumns['AccountNumber']]
                        $checkNumberValue = $cells[$row, $fieldColumns['CheckNumber']]
                        $amountValue = $cells[$row, $fieldColumns['CheckAmount']]
                        $payeeValue = $cells[$row, $fieldColumns['Payee']]
                        $checkDateValue = $cells[$row, $fieldColumns['CheckDate']]

                        if (($null -eq $accountValue -or [string]::IsNullOrWhiteSpace([string]$accountValue)) -and
                            ($null -eq $checkNumberValue -or [string]::IsNullOrWhiteSpace([string]$checkNumberValue)) -and
                            ($null -eq $amountValue -or [string]::IsNullOrWhiteSpace([string]$amountValue)) -and
                            ($null -eq $payeeValue -or [string]::IsNullOrWhiteSpace([string]$payeeValue)) -and
                            ($null -eq $checkDateValue -or [string]::IsNullOrWhiteSpace([string]$checkDateValue))) {
                            continue
                        }

                        $account = ConvertTo-PaddedDigits $accountValue 10 'AccountNumber'
                        $checkNumber = ConvertTo-PaddedDigits $checkNumberValue 10 'CheckNumber'
                        $amountCents = ConvertTo-CheckAmountCents $amountValue
                        $checkDate = ConvertTo-CheckDate $checkDateValue
                        $payee = ([string]$payeeValue).Trim()
                        if ([string]::IsNullOrWhiteSpace($payee)) { throw "Payee is empty at row $row in $($inputFile.Name)." }
                        if ($payee.Length -gt 40) { throw "Payee exceeds 40 characters at row $row in $($inputFile.Name)." }

                        # Fixed values: CheckStatus=00, CheckNew=1, CheckUpdate=0, TransactionCode=001.
                        $amountDigits = ([long]$amountCents).ToString('D16', $invariant)
                        $line = $account + $checkNumber + $transactionDate + $transactionTime + $amountDigits + $payee.PadRight(40) + $checkDate + '00' + '1' + '0' + '001'
                        if ($line.Length -ne 105) { throw "Generated record has unexpected length $($line.Length) at row $row in $($inputFile.Name)." }
                        $lines.Add($line + (' ' * 10))

                        $accountFactor = [decimal]::Parse($account.Substring(4, 6), $invariant)
                        $checkAmount = $amountCents / 100
                        $calculationTotal += ($accountFactor * $checkAmount) + 0 + 1 + 0 + 1
                        $amountTotalCents += $amountCents
                        $dataCount++
                    }

                    if ($dataCount -eq 0) { throw "No data rows found in worksheet '$($worksheet.Name)' of $($inputFile.Name)." }
                    if ($dataCount -gt 999999) { throw "Entry count exceeds the six-digit footer field in $($inputFile.Name)." }

                    $calculationCents = [decimal]::Round(($calculationTotal * 100), 0, [MidpointRounding]::AwayFromZero)
                    if ($calculationCents -lt 0 -or $calculationCents -gt 99999999999999999999) { throw "Calculated footer value exceeds the 20-digit field in $($inputFile.Name)." }
                    $calculationDigits = $calculationCents.ToString('0', $invariant).PadLeft(20, '0')
                    $sumAmountText = $amountTotalCents.ToString('0', $invariant)
                    if ($sumAmountText.Length -gt 16) { $sumAmountText = $sumAmountText.Substring(0, 16) }
                    $sumAmountDigits = $sumAmountText.PadLeft(16, '0')
                    $lines.Add(('9999999999' + $calculationDigits + $dataCount.ToString('D6', $invariant) + $sumAmountDigits))

                    $outputFile = Join-Path $outputDirectory ($inputFile.BaseName + '.txt')
                    [System.IO.File]::WriteAllLines($outputFile, $lines.ToArray(), $utf8NoBom)
                    Write-Host ("Created {0} ({1} records)" -f $outputFile, $dataCount)
                    $didConvert = $true
                    break
                }
                finally {
                    if ($usedRange) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) }
                    if ($worksheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) }
                }
            }

            if (-not $didConvert) {
                Write-Warning ("No worksheet in {0} contains all required headers: {1}" -f $inputFile.Name, ($requiredFields -join ', '))
            }
        }
        finally {
            if ($workbook) {
                try { $workbook.Close($false) } catch { }
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
            }
        }
    }
}
catch {
    Write-Error ("Conversion failed: {0}" -f $_.Exception.Message)
    exit 1
}
finally {
    if ($excel) {
        try { $excel.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}
