/**
 * Excel Reader Module
 * Reads and validates Excel data for Claims automation
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

/**
 * Normalize column names (trim, lowercase, handle variations)
 */
function normalizeColumnName(name) {
  if (!name) return '';
  return name.toString().trim().toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/custom\s*id/gi, 'customid')
    .replace(/date\s*of\s*service/gi, 'dos')
    .replace(/appointment\s*date/gi, 'appointmentdate')
    .replace(/orig\s*appt\.?\s*date/gi, 'orig appt. date')
    .replace(/original\s*appointment\s*date/gi, 'original appointment date');
}

/**
 * Parse date from various formats (MM/DD/YYYY, YYYY-MM-DD, etc.)
 */
function parseDate(dateStr) {
  if (!dateStr) return null;
  
  const str = dateStr.toString().trim();
  if (!str) return null;

  // Try MM/DD/YYYY format
  const mmddyyyy = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(str);
  if (mmddyyyy) {
    const month = parseInt(mmddyyyy[1], 10);
    const day = parseInt(mmddyyyy[2], 10);
    const year = parseInt(mmddyyyy[3], 10);
    const date = new Date(year, month - 1, day);
    if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) {
      return date;
    }
  }

  // Try YYYY-MM-DD format
  const yyyymmdd = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(str);
  if (yyyymmdd) {
    const year = parseInt(yyyymmdd[1], 10);
    const month = parseInt(yyyymmdd[2], 10);
    const day = parseInt(yyyymmdd[3], 10);
    const date = new Date(year, month - 1, day);
    if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) {
      return date;
    }
  }

  // Try Excel serial date number
  if (/^\d+$/.test(str)) {
    const serial = parseInt(str, 10);
    if (serial > 0 && serial < 1000000) {
      // Excel epoch is 1900-01-01, but Excel incorrectly treats 1900 as leap year
      const excelEpoch = new Date(1899, 11, 30);
      const date = new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
  }

  // Try native Date parsing
  const date = new Date(str);
  if (!isNaN(date.getTime())) {
    return date;
  }

  return null;
}

/**
 * Format date to MM/DD/YYYY
 */
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') {
    const parsed = parseDate(date);
    if (!parsed) return date;
    date = parsed;
  }
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();
  return `${month}/${day}/${year}`;
}

/**
 * Parse CPT codes (can be comma/space separated)
 */
function parseCPTCodes(cptStr) {
  if (!cptStr) return [];
  return cptStr.toString()
    .split(/[,\s]+/)
    .map(code => code.trim())
    .filter(code => code.length > 0);
}

/**
 * Validate a row
 */
function validateRow(row, rowIndex) {
  const errors = [];

  if (!row.mrn || !row.mrn.toString().trim()) {
    errors.push('MRN is required');
  }

  if (row.dos) {
    const dosDate = parseDate(row.dos);
    if (!dosDate) {
      errors.push(`DOS "${row.dos}" is not a valid date`);
    } else {
      row.dosDate = dosDate;
      row.dosFormatted = formatDate(dosDate);
    }
  }

  // Check for appointment date in various column name variations
  const appointmentDateValue = row.appointmentdate || row['appointment date'] || row['orig appt. date'] || row['orig appt date'] || row['original appointment date'];
  
  if (appointmentDateValue) {
    const apptDate = parseDate(appointmentDateValue);
    if (!apptDate) {
      errors.push(`Appointment Date "${appointmentDateValue}" is not a valid date`);
    } else {
      row.appointmentDate = apptDate;
      row.appointmentDateFormatted = formatDate(apptDate);
    }
  }

  if (row.cpt) {
    row.cptCodes = parseCPTCodes(row.cpt);
  } else {
    row.cptCodes = [];
  }

  return {
    isValid: errors.length === 0,
    errors
  };
}

/**
 * Find Excel file in directory
 */
function findExcelFile(directory, preferredName = null) {
  const files = fs.readdirSync(directory);
  // Filter out report files and only get actual data files
  const xlsxFiles = files.filter(f => {
    const lower = f.toLowerCase();
    const isExcel = lower.endsWith('.xlsx') || lower.endsWith('.xls');
    const isReport = lower.includes('report') || lower.includes('claims_processing_report');
    return isExcel && !isReport;
  });

  if (xlsxFiles.length === 0) {
    // If no non-report files found, fall back to all Excel files
    const allXlsxFiles = files.filter(f => 
    f.toLowerCase().endsWith('.xlsx') || f.toLowerCase().endsWith('.xls')
  );
    if (allXlsxFiles.length === 0) {
    throw new Error('No Excel files found in directory');
    }
    // Use all files if no non-report files
    xlsxFiles.push(...allXlsxFiles);
  }

  // If preferred name is specified, try to find it
  if (preferredName) {
    const preferred = xlsxFiles.find(f => 
      f.toLowerCase() === preferredName.toLowerCase() ||
      f.toLowerCase().includes(preferredName.toLowerCase())
    );
    if (preferred) {
      return path.join(directory, preferred);
    }
  }

  // Otherwise, get the newest file (excluding reports)
  const fileStats = xlsxFiles.map(f => ({
    name: f,
    path: path.join(directory, f),
    mtime: fs.statSync(path.join(directory, f)).mtime
  }));

  fileStats.sort((a, b) => b.mtime - a.mtime);
  return fileStats[0].path;
}

/**
 * Read Excel file and return normalized rows
 */
function readExcelFile(filePath, sheetName = null) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Excel file not found: ${filePath}`);
  }

  console.log(`📖 Reading Excel file: ${path.basename(filePath)}`);

  const workbook = XLSX.readFile(filePath);
  
  // Find the sheet to use
  let sheet;
  if (sheetName) {
    sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${workbook.SheetNames.join(', ')}`);
    }
  } else {
    // Try "Input" sheet first, otherwise use first sheet
    sheet = workbook.Sheets['Input'] || workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) {
      throw new Error('No sheets found in workbook');
    }
  }

  const usedSheetName = sheetName || (workbook.Sheets['Input'] ? 'Input' : workbook.SheetNames[0]);
  console.log(`📄 Using sheet: "${usedSheetName}"`);

  // Convert to JSON
  const rows = XLSX.utils.sheet_to_json(sheet, { raw: false });

  if (rows.length === 0) {
    throw new Error('No data rows found in sheet');
  }

  // Normalize column names
  const normalizedRows = rows.map((row, index) => {
    const normalized = {};
    for (const [key, value] of Object.entries(row)) {
      const normalizedKey = normalizeColumnName(key);
      normalized[normalizedKey] = value;
    }
    // Preserve original row number (1-indexed, +1 for header)
    normalized._rowNumber = index + 2;
    return normalized;
  });

  // Map to expected field names
  const mappedRows = normalizedRows.map(row => ({
    mrn: row.mrn || row['patient mrn'] || '',
    customId: row.customid || row['custom id'] || '',
    dos: row.dos || row['date of service'] || '',
    appointmentDate: row.appointmentdate || row['appointment date'] || row['orig appt. date'] || row['orig appt date'] || row['original appointment date'] || '',
    cpt: row.cpt || row['cpt code'] || row['cpt codes'] || row['cpt(s)'] || row['cpts'] || '',
    icd: row.icd || row['icd code'] || row['icd codes'] || '',
    // Include all other columns for potential use
    ...Object.fromEntries(
      Object.entries(row).filter(([key]) => 
        !['mrn', 'customid', 'custom id', 'dos', 'date of service', 
          'appointmentdate', 'appointment date', 'orig appt. date', 'orig appt date', 'original appointment date',
          'cpt', 'cpt code', 'cpt codes', 'icd', 'icd code', 'icd codes', '_rownumber'].includes(key)
      )
    ),
    _rowNumber: row._rowNumber
  }));

  console.log(`✅ Loaded ${mappedRows.length} rows from Excel`);
  console.log(`📋 Sample columns found: ${Object.keys(normalizedRows[0] || {}).join(', ')}`);

  return mappedRows;
}

/**
 * Main function to read and validate Excel data
 */
function loadExcelData(config) {
  const excelPath = config.excelPath || process.env.EXCEL_FILE_PATH;
  const excelDir = config.excelDir || process.env.EXCEL_DIR || __dirname;
  const preferredFileName = config.excelFileName || process.env.EXCEL_FILE_NAME;

  let filePath;
  
  if (excelPath) {
    // Use explicit path
    filePath = path.isAbsolute(excelPath) ? excelPath : path.join(excelDir, excelPath);
  } else {
    // Auto-find Excel file
    filePath = findExcelFile(excelDir, preferredFileName);
    console.log(`📁 Auto-selected Excel file: ${path.basename(filePath)}`);
  }

  const rows = readExcelFile(filePath, config.excelSheetName || process.env.EXCEL_SHEET_NAME);

  // Validate each row
  const validatedRows = rows.map((row, index) => {
    const validation = validateRow(row, index);
    return {
      ...row,
      _isValid: validation.isValid,
      _errors: validation.errors,
      _filePath: filePath
    };
  });

  const validRows = validatedRows.filter(r => r._isValid);
  const invalidRows = validatedRows.filter(r => !r._isValid);

  console.log(`✅ Valid rows: ${validRows.length}`);
  if (invalidRows.length > 0) {
    console.log(`⚠️  Invalid rows: ${invalidRows.length}`);
    invalidRows.forEach(row => {
      console.log(`   Row ${row._rowNumber}: ${row._errors.join('; ')}`);
    });
  }

  return {
    filePath,
    allRows: validatedRows,
    validRows,
    invalidRows
  };
}

module.exports = {
  loadExcelData,
  parseDate,
  formatDate,
  parseCPTCodes,
  validateRow,
  findExcelFile,
  readExcelFile
};

