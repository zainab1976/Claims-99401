/**
 * Report Generator Module
 * Creates output Excel/CSV reports with processing status per row
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

/**
 * Generate output report
 */
function generateReport(results, outputPath = null) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
  const defaultFileName = `claims_processing_report_${timestamp}.xlsx`;
  const reportPath = outputPath || path.join(__dirname, defaultFileName);

  // Prepare report data
  const reportRows = results.map(result => ({
    'Row Number': result.rowNumber || result._rowNumber || '',
    'MRN': result.mrn || '',
    'Custom ID': result.customId || '',
    'DOS': result.dos || '',
    'Appointment Date': result.appointmentDate || '',
    'CPT Codes': Array.isArray(result.cptCodes) ? result.cptCodes.join(', ') : (result.cpt || ''),
    'Status': result.status || 'Unknown',
    'Error Message': result.errorMessage || '',
    'Generated ID': result.generatedId || '',
    'Processing Time (ms)': result.processingTime || '',
    'Timestamp': result.timestamp || new Date().toISOString(),
    'Notes': result.notes || ''
  }));

  // Create workbook
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(reportRows);

  // Set column widths
  const colWidths = [
    { wch: 12 }, // Row Number
    { wch: 15 }, // MRN
    { wch: 15 }, // Custom ID
    { wch: 12 }, // DOS
    { wch: 18 }, // Appointment Date
    { wch: 20 }, // CPT Codes
    { wch: 12 }, // Status
    { wch: 40 }, // Error Message
    { wch: 15 }, // Generated ID
    { wch: 20 }, // Processing Time
    { wch: 25 }, // Timestamp
    { wch: 30 }  // Notes
  ];
  worksheet['!cols'] = colWidths;

  XLSX.utils.book_append_sheet(workbook, worksheet, 'Processing Report');

  // Add summary sheet
  const summary = {
    'Total Rows Processed': results.length,
    'Successful': results.filter(r => r.status === 'Success').length,
    'Failed': results.filter(r => r.status === 'Failed').length,
    'Skipped': results.filter(r => r.status === 'Skipped').length,
    'Report Generated': new Date().toISOString()
  };

  const summarySheet = XLSX.utils.json_to_sheet([summary]);
  summarySheet['!cols'] = [{ wch: 25 }, { wch: 15 }];
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');

  // Write file
  XLSX.writeFile(workbook, reportPath);
  console.log(`📊 Report generated: ${path.basename(reportPath)}`);

  return reportPath;
}

/**
 * Generate CSV report (alternative format)
 */
function generateCSVReport(results, outputPath = null) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
  const defaultFileName = `claims_processing_report_${timestamp}.csv`;
  const reportPath = outputPath || path.join(__dirname, defaultFileName);

  const headers = [
    'Row Number', 'MRN', 'Custom ID', 'DOS', 'Appointment Date', 
    'CPT Codes', 'Status', 'Error Message', 'Generated ID', 
    'Processing Time (ms)', 'Timestamp', 'Notes'
  ];

  const csvRows = [
    headers.join(','),
    ...results.map(result => [
      result.rowNumber || result._rowNumber || '',
      `"${(result.mrn || '').toString().replace(/"/g, '""')}"`,
      `"${(result.customId || '').toString().replace(/"/g, '""')}"`,
      `"${(result.dos || '').toString().replace(/"/g, '""')}"`,
      `"${(result.appointmentDate || '').toString().replace(/"/g, '""')}"`,
      `"${(Array.isArray(result.cptCodes) ? result.cptCodes.join(', ') : (result.cpt || '')).toString().replace(/"/g, '""')}"`,
      `"${(result.status || 'Unknown').toString().replace(/"/g, '""')}"`,
      `"${(result.errorMessage || '').toString().replace(/"/g, '""')}"`,
      `"${(result.generatedId || '').toString().replace(/"/g, '""')}"`,
      result.processingTime || '',
      `"${(result.timestamp || new Date().toISOString()).toString().replace(/"/g, '""')}"`,
      `"${(result.notes || '').toString().replace(/"/g, '""')}"`
    ].join(','))
  ];

  fs.writeFileSync(reportPath, csvRows.join('\n'), 'utf8');
  console.log(`📊 CSV Report generated: ${path.basename(reportPath)}`);

  return reportPath;
}

/**
 * Update original Excel file with processing status
 */
function updateOriginalExcel(originalFilePath, results, sheetName = null) {
  const XLSX = require('xlsx');
  const fs = require('fs');
  
  if (!fs.existsSync(originalFilePath)) {
    throw new Error(`Original Excel file not found: ${originalFilePath}`);
  }

  console.log(`📝 Updating original Excel file: ${path.basename(originalFilePath)}`);

  // Read the original workbook
  const workbook = XLSX.readFile(originalFilePath);
  
  // Find the sheet to update
  let sheet;
  let usedSheetName;
  if (sheetName) {
    sheet = workbook.Sheets[sheetName];
    usedSheetName = sheetName;
  } else {
    // Try "Input" sheet first, otherwise use first sheet
    sheet = workbook.Sheets['Input'] || workbook.Sheets[workbook.SheetNames[0]];
    usedSheetName = workbook.Sheets['Input'] ? 'Input' : workbook.SheetNames[0];
  }

  if (!sheet) {
    throw new Error('Sheet not found in workbook');
  }

  // Convert to JSON to get headers and data
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  
  if (rows.length === 0) {
    throw new Error('No data found in sheet');
  }

  // Find or create Status column - always at the end
  const headers = rows[0];
  let statusColumnIndex = headers.findIndex(h => 
    h && h.toString().toLowerCase() === 'status'
  );

  // If Status column exists but not at the end, remove it first
  if (statusColumnIndex !== -1 && statusColumnIndex < headers.length - 1) {
    // Remove Status column from all rows
    for (let i = 0; i < rows.length; i++) {
      rows[i].splice(statusColumnIndex, 1);
    }
    statusColumnIndex = -1; // Reset to add at end
  }

  // If Status column doesn't exist, add it at the end
  if (statusColumnIndex === -1) {
    statusColumnIndex = headers.length;
    headers.push('Status');
    // Ensure all data rows have enough columns
    for (let i = 1; i < rows.length; i++) {
      while (rows[i].length < headers.length) {
        rows[i].push('');
      }
    }
  }

  // Create a map of results by row number - group multiple results per row
  const resultsMap = new Map();
  results.forEach(result => {
    const rowNum = result.rowNumber || result._rowNumber;
    if (rowNum) {
      if (!resultsMap.has(rowNum)) {
        resultsMap.set(rowNum, []);
      }
      resultsMap.get(rowNum).push(result);
    }
  });

  // Helper function to aggregate statuses from multiple results
  // Priority: Billed > Appt. Cancelled > Success > Failed > Others
  function aggregateStatus(results) {
    if (!results || results.length === 0) return 'Unknown';
    
    const statuses = results.map(r => r.status || 'Unknown');
    const uniqueStatuses = [...new Set(statuses)];
    
    // Priority order: Billed > Appt. Cancelled > Success > Failed
    if (uniqueStatuses.includes('Billed')) {
      // If all are Billed, return Billed; otherwise show combination
      if (uniqueStatuses.length === 1) return 'Billed';
      return `Billed (${results.length} patients)`;
    }
    if (uniqueStatuses.includes('Appt. Cancelled')) {
      // If all are Appt. Cancelled, return Appt. Cancelled; otherwise show combination
      if (uniqueStatuses.length === 1) return 'Appt. Cancelled';
      return `Appt. Cancelled (${results.length} patients)`;
    }
    if (uniqueStatuses.includes('Success')) {
      // If all are Success, return Success; otherwise show combination
      if (uniqueStatuses.length === 1) return 'Success';
      return `Success (${results.length} patients)`;
    }
    if (uniqueStatuses.includes('Failed')) {
      // If any failed, prioritize showing failed status
      const failedCount = statuses.filter(s => s === 'Failed').length;
      if (failedCount === results.length) return 'Failed';
      return `Failed (${failedCount}/${results.length} patients)`;
    }
    
    // If multiple different statuses, combine them
    if (uniqueStatuses.length > 1) {
      return uniqueStatuses.join(', ');
    }
    
    // Single status
    return uniqueStatuses[0];
  }

  // Update rows with status
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const excelRowNumber = i + 1; // Excel row number (1-indexed, +1 for header)
    
    // Ensure row has enough columns
    while (row.length <= statusColumnIndex) {
      row.push('');
    }
    
    // Update status if we have results for this row
    const rowResults = resultsMap.get(excelRowNumber);
    if (rowResults && rowResults.length > 0) {
      // Aggregate status from all results for this row
      const statusValue = aggregateStatus(rowResults);
      row[statusColumnIndex] = statusValue;
      console.log(`   📝 Row ${excelRowNumber}: Updated Status to "${statusValue}" (${rowResults.length} patient(s))`);
    } else {
      // If no result found for this row, mark as "Not Processed"
      if (!row[statusColumnIndex] || row[statusColumnIndex] === '') {
        row[statusColumnIndex] = 'Not Processed';
      }
    }
  }

  // Convert back to worksheet
  const updatedSheet = XLSX.utils.aoa_to_sheet(rows);
  
  // Preserve column widths if they exist
  if (sheet['!cols']) {
    updatedSheet['!cols'] = [...sheet['!cols']];
    // Ensure Status column has proper width (always at the end)
    if (statusColumnIndex >= updatedSheet['!cols'].length) {
      updatedSheet['!cols'].push({ wch: 15 });
    } else {
      // Update width for Status column if it already existed
      updatedSheet['!cols'][statusColumnIndex] = { wch: 15 };
    }
  } else {
    // If no column widths exist, create default widths
    updatedSheet['!cols'] = headers.map(() => ({ wch: 15 }));
    updatedSheet['!cols'][statusColumnIndex] = { wch: 15 }; // Status column width
  }

  // Update the sheet in workbook
  workbook.Sheets[usedSheetName] = updatedSheet;

  // Write back to original file
  XLSX.writeFile(workbook, originalFilePath);
  console.log(`✅ Updated original Excel file with status column`);

  return originalFilePath;
}

module.exports = {
  generateReport,
  generateCSVReport,
  updateOriginalExcel
};



