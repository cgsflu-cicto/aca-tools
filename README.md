# Excel Converters

This repository contains two Windows tools for converting Excel workbooks:

- **PACS converter** (`pacs-converter/ConvertAll.bat`) creates CSV files.
- **ACIC converter** (`acic-converter/ConvertAll.bat`) creates ACIC fixed-width text files.

## Requirements

- Windows with PowerShell 5.1 or higher
- Microsoft Excel installed

## PACS Converter

1. Place the PACS Excel files (`.xls` or `.xlsx`) in the `pacs-converter` folder.
2. Double-click `pacs-converter/ConvertAll.bat`.
3. CSV files are created beside the workbooks with the same base names.

## Excel File Format

PACS workbooks should have these columns in order:

| Account | Name | Amount |
|---------|------|--------|
| 123456 | John Doe | 1500.00 |
| 789012 | Jane Smith | 2000.00 |

- **Row 1**: Headers (Account, Name, Amount)
- **Row 2+**: Data rows

## Output Format

CSV files are created with the format:
```
AccountNumber,"Full Name",AmountInCents
```

Example:
```
123456,"John Doe",150000
789012,"Jane Smith",200000
```

**Note**: Amounts are converted to cents (no decimal point).

## ACIC Converter

1. Place the ACIC Excel files (`.xls` or `.xlsx`) in the `acic-converter` folder.
2. Double-click `acic-converter/ConvertAll.bat`.
3. The converter finds a worksheet with the required column headers and creates a `.txt` file beside each workbook.

The ACIC worksheet must contain these headers (column order can vary):

| Header | Description |
|--------|-------------|
| `AccountNumber` | Account number; converted to a 10-digit field |
| `CheckNumber` | Check number; converted to a 10-digit field |
| `CheckAmount` | Check amount; written in cents |
| `Payee` | Payee name, up to 40 characters |
| `CheckDate` | Check date; written as `MMddyyyy` |

The ACIC output uses 105-character detail records followed by a totals record. The converter skips worksheets that do not contain all five headers and reports a warning if no matching worksheet is found in a workbook.
