<# :
@echo off & cd /d "%~dp0" & cls
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex (gc '%~f0' -Raw)"
pause & exit /b
#>

$xlUp = -4162

Write-Host "Please wait for the conversion process to finish...`n" -ForegroundColor Blue

$excel = New-Object -ComObject Excel.Application

try {
    Get-ChildItem -Filter "*.xls*" | ForEach-Object {
        $file = "$($_.BaseName).csv"
        $path = $_.FullName
        $book = $null
        $sheet = $null
        
        try {
            $book = $excel.Workbooks.Open($path)
            $sheet = $book.Worksheets.Item(1)
            $rowCount = ($sheet.Rows.Count)
            $rowLast = (@('A','B','C') | ForEach-Object { ($sheet.Range("$_$rowCount").End($xlUp).Row) } | Measure-Object -Maximum).Maximum

            if (Test-Path ".\$file") {
                Remove-Item ".\$file" -Force -Confirm:$false
            }

            # Read entire range at once as 2D array - much faster than row-by-row
            $range = $sheet.Range("A2:C$rowLast")
            $data = $range.Value2
            
            # Build output in memory using ArrayList (much faster than repeated Out-File)
            $outputLines = [System.Collections.ArrayList]::new()
            
            # Process array directly (rows are 1-indexed in the array)
            for ($i = 1; $i -le ($rowLast - 1); $i++) {
                $account = ("$($data[$i, 1])" -replace "\W",'')
                $name = ("$($data[$i, 2])".Trim() -replace "\s+",' ')
                $amount = 0
                [double]::TryParse("$($data[$i, 3])", [ref]$amount) | Out-Null
                $amount = (("{0:0.00}" -f $amount) -replace "\.",'')

                [void]$outputLines.Add("$account,`"$name`",$amount")
            }
            
            # Write all lines at once
            $outputLines | Out-File ".\$file" -Encoding utf8
        }
        catch {
            Write-Host "Error processing $($_.Name): $_`n" -ForegroundColor Red
        }
        finally {
            # Release COM objects for this workbook
            if ($sheet) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
            }
            if ($book) {
                $book.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book) | Out-Null
            }
        }
    }
}
finally {
    # Clean up Excel application
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Conversion process finished`n" -ForegroundColor Green
