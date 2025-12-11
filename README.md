# PACS Salary Converter

Converts Excel salary files to CSV format.

## Requirements

- Windows with PowerShell 5.1 or higher
- Microsoft Excel installed

## Usage

1. Place your Excel files (`.xls` or `.xlsx`) in the same folder as `ConvertAll.bat`
2. Double-click `ConvertAll.bat` to run the conversion
3. CSV files will be created with the same names as your Excel files

## Excel File Format

Your Excel files should have the following structure:

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