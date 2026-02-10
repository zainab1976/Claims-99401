# QHSLab Claims Automation

Automated claims processing script for QHSLab that reads data from Excel files and processes each row through the QHSLab web interface.

## Features

- ✅ Reads data from Excel files (`.xlsx`)
- ✅ Validates row data (MRN, dates, CPT codes)
- ✅ Processes each row through the QHSLab automation workflow
- ✅ Generates detailed processing reports (Excel format)
- ✅ Supports dry-run mode for validation
- ✅ Auto-detects Excel files in project directory
- ✅ Flexible column name matching (case-insensitive, handles variations)

## Prerequisites

- Node.js (v14 or higher)
- npm or yarn
- QHSLab account credentials

## Installation

1. Install dependencies:
```bash
npm install
```

2. Create a `.env` file in the project root:
```env
QHSLAB_EMAIL=your-email@example.com
QHSLAB_PASSWORD=your-password
EXCEL_FILE_PATH=path/to/your/file.xlsx  # Optional: specify exact file
EXCEL_DIR=.                              # Optional: directory to search for Excel files
EXCEL_FILE_NAME=your-file.xlsx          # Optional: preferred file name
EXCEL_SHEET_NAME=Input                   # Optional: sheet name (defaults to "Input" or first sheet)
DRY_RUN=false                            # Set to "true" for validation only
```

## How to Run Using Excel Input

### 1. Prepare Your Excel File

Place your Excel file (`.xlsx`) in the project folder or specify the path in `.env`.

**Required Columns:**
- `MRN` (required) - Patient Medical Record Number
- `Custom ID` or `CustomID` (optional) - Custom identifier
- `DOS` or `Date of Service` (optional) - Date of service (MM/DD/YYYY or YYYY-MM-DD)
- `Appointment Date` (optional) - Appointment date (MM/DD/YYYY or YYYY-MM-DD)
- `CPT` or `CPT Code` or `CPT Codes` (optional) - CPT codes (comma or space separated)

**Example Excel Structure:**

| MRN | Custom ID | DOS | Appointment Date | CPT Codes |
|-----|-----------|-----|------------------|-----------|
| 12345 | CUST001 | 12/31/2025 | 12/31/2025 | 96136 |
| 12346 | CUST002 | 01/15/2026 | 01/15/2026 | 96136, 96137 |
| 12347 | | 02/20/2026 | | 96136 |

**Notes:**
- Column names are case-insensitive and spaces are normalized
- Dates can be in MM/DD/YYYY or YYYY-MM-DD format
- Multiple CPT codes can be separated by commas or spaces
- The script will use the first sheet or a sheet named "Input" if it exists

### 2. Run the Script

**Normal execution:**
```bash
node index.js
```

**Dry-run mode (validation only, no automation):**
```bash
node index.js --dry-run
```

Or set in `.env`:
```env
DRY_RUN=true
```

### 3. Review the Output

After processing, the script generates a report file:
- `claims_processing_report_YYYY-MM-DDTHH-MM-SS.xlsx`

The report includes:
- **Processing Report** sheet: Status for each row (Success/Failed/Skipped), error messages, processing time
- **Summary** sheet: Overall statistics

## Configuration Options

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `QHSLAB_EMAIL` | QHSLab login email | Required |
| `QHSLAB_PASSWORD` | QHSLab login password | Required |
| `EXCEL_FILE_PATH` | Full path to Excel file | Auto-detect |
| `EXCEL_DIR` | Directory to search for Excel files | Project root |
| `EXCEL_FILE_NAME` | Preferred file name (partial match) | None |
| `EXCEL_SHEET_NAME` | Sheet name to read | "Input" or first sheet |
| `DRY_RUN` | Enable dry-run mode | false |

### Excel File Detection

If `EXCEL_FILE_PATH` is not specified, the script will:
1. Look for a file matching `EXCEL_FILE_NAME` (if provided)
2. Otherwise, use the newest `.xlsx` file in `EXCEL_DIR`
3. Log which file was selected

## Output Report Format

The generated report includes:

- **Row Number**: Original row number from Excel
- **MRN**: Patient MRN
- **Custom ID**: Custom identifier
- **DOS**: Date of Service
- **Appointment Date**: Appointment date
- **CPT Codes**: CPT codes used
- **Status**: Success / Failed / Skipped
- **Error Message**: Error details (if failed)
- **Generated ID**: Any ID generated during processing
- **Processing Time (ms)**: Time taken to process the row
- **Timestamp**: When the row was processed
- **Notes**: Additional notes

## Validation Rules

Rows are validated before processing:

- ✅ **MRN**: Required (non-empty)
- ✅ **DOS**: If provided, must be a valid date
- ✅ **Appointment Date**: If provided, must be a valid date
- ✅ **CPT Codes**: Optional, but if provided will be parsed and validated

Invalid rows are skipped and included in the report with error details.

## Troubleshooting

### Excel file not found
- Ensure the file is in the project directory or specify `EXCEL_FILE_PATH` in `.env`
- Check file permissions

### Login fails
- Verify `QHSLAB_EMAIL` and `QHSLAB_PASSWORD` in `.env`
- Check network connectivity

### Column not found
- Column names are case-insensitive and handle variations
- Check that your column names match: MRN, Custom ID, DOS, Appointment Date, CPT

### Processing errors
- Check the generated report for specific error messages
- Verify that the QHSLab interface hasn't changed
- Try dry-run mode first to validate data

## Development

### Project Structure

```
Claims/
├── index.js              # Main automation script
├── excelReader.js        # Excel reading and validation
├── reportGenerator.js    # Report generation
├── package.json          # Dependencies
├── .env                  # Environment variables (create this)
└── README.md            # This file
```

### Adding New Fields

To add support for additional Excel columns:

1. Update `excelReader.js` to normalize and map the new column
2. Update `processRow()` in `index.js` to use the new field
3. Update this README with the new column name

## License

ISC



