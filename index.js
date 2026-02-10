/**
 * QHSLab Automation Script
 * -----------------------
 * - Launches a Chromium browser
 * - Logs into QHSLab
 * - Navigates to the Accounts page
 *
 * Environment Variables Required:
 *   QHSLAB_EMAIL
 *   QHSLAB_PASSWORD
 */

require("dotenv").config();
const { chromium } = require("playwright");
const path = require("path");
const { loadExcelData, formatDate } = require("./excelReader");
const { generateReport, updateOriginalExcel } = require("./reportGenerator");

/* =======================
   Claims Button State
======================= */
let claimsClickedOnce = false;
let accountsPageNavigated = false;
let entitySourceFilterApplied = false;
let assessmentResourceFilterApplied = false;

/* =======================
   Configuration
======================= */
const CONFIG = {
  BASE_URL: "https://my.qhslab.com",

  BROWSER: {
    HEADLESS: false,
    SLOW_MO: 2000,
    ARGS: [
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-dev-shm-usage",
      "--disable-web-security",
      "--disable-extensions",
      "--disable-blink-features=AutomationControlled",
      "--disable-popup-blocking",
    ],
  },

  TIMEOUTS: {
    PAGE_LOAD: 30_000,
    ELEMENT_WAIT: 10_000,
    SHORT: 1_000,
  },

  CREDS: {
    email: process.env.QHSLAB_EMAIL,
    password: process.env.QHSLAB_PASSWORD,
  },

  EXCEL: {
    filePath: process.env.EXCEL_FILE_PATH,
    fileDir: process.env.EXCEL_DIR || __dirname,
    fileName: process.env.EXCEL_FILE_NAME,
    sheetName: process.env.EXCEL_SHEET_NAME || null,
  },

  DRY_RUN: process.env.DRY_RUN === 'true' || process.argv.includes('--dry-run'),
};

/* =======================
   Utilities
======================= */
function attachPageListeners(page) {
  page.on("crash", () => console.error("❌ Page crashed"));
  page.on("close", () => console.error("❌ Page closed unexpectedly"));
  page.on("error", (err) =>
    console.error("❌ Page error:", err.message)
  );
  page.on("console", (msg) => {
    if (msg.type() === "error") {
      console.error("❌ Browser console error:", msg.text());
    }
  });
}

/* =======================
   Claims Button Handler
======================= */
async function clickClaimsOnce(page) {
  if (claimsClickedOnce) {
    console.log('ℹ️ Claims button already clicked once — skipping');
    return;
  }

  const claimsBtn = page.getByRole('button', { name: 'Dash@3x Claims' });

  try {
    await claimsBtn.waitFor({ state: 'visible', timeout: 5000 });
    await claimsBtn.click();
    claimsClickedOnce = true;
    console.log('✅ Claims button clicked (one-time)');
  } catch (error) {
    console.error('❌ Failed to click Claims button:', error.message);
    throw error;
  }
}

/* =======================
   Filter Click Handler
======================= */
async function clickCustomIdFilter(page) {
  console.log("⏳ Waiting for Claims page to load...");
  
  // Wait for Claims page to be visible (wait for "Claims" heading)
  try {
    await page.waitForSelector('text=Claims', { state: 'visible', timeout: CONFIG.TIMEOUTS.ELEMENT_WAIT });
    await page.waitForTimeout(3000); // Additional wait for table to render
  } catch (error) {
    console.log("⚠️ Claims heading not found, continuing anyway...");
  }

  // Try multiple locator strategies - specifically targeting Custom ID column (before MRN)
  const strategies = [
    // Strategy 1: Find table, locate Custom ID header, get its column index, then find filter in that exact column
    async () => {
      const table = page.locator('table').first();
      await table.waitFor({ state: 'visible', timeout: 5000 });
      const headers = table.locator('th, [role="columnheader"]');
      const headerCount = await headers.count();
      let customIdIndex = -1;
      
      for (let i = 0; i < headerCount; i++) {
        const headerText = await headers.nth(i).textContent();
        // Match "Custom ID" exactly (not MRN or Patient MRN)
        if (headerText && /^Custom ID$/i.test(headerText.trim())) {
          customIdIndex = i;
          break;
        }
      }
      
      if (customIdIndex === -1) {
        throw new Error('Custom ID header not found');
      }
      
      // Find the filter row and get the filter at the same column index
      const filterRow = table.locator('tr').filter({ has: page.locator('input[placeholder*="filter" i]') }).first();
      await filterRow.waitFor({ state: 'visible', timeout: 5000 });
      const filters = filterRow.locator('input[placeholder*="filter" i]');
      const filter = filters.nth(customIdIndex);
      
      // Verify this is Custom ID by checking the header above
      const headerAbove = await filter.evaluate((el) => {
        const cell = el.closest('td, th');
        if (!cell) return false;
        const row = cell.closest('tr');
        if (!row) return false;
        const table = row.closest('table');
        if (!table) return false;
        const headerRow = table.querySelector('tr');
        if (!headerRow) return false;
        const cells = Array.from(headerRow.querySelectorAll('th, td, [role="columnheader"]'));
        const cellIndex = Array.from(row.querySelectorAll('td, th')).indexOf(cell);
        if (cellIndex >= 0 && cellIndex < cells.length) {
          const headerText = cells[cellIndex].textContent;
          return /^Custom ID$/i.test(headerText?.trim() || '');
        }
        return false;
      });
      
      if (!headerAbove) {
        throw new Error('Filter is not under Custom ID header');
      }
      
      return filter;
    },
    // Strategy 2: Find by cell role with exact Custom ID name match
    () => page.getByRole('cell', { name: /^Custom ID.*Sort by Custom ID$/i }).getByPlaceholder('filter'),
    // Strategy 3: Find all filter inputs and verify which one is directly under Custom ID header
    async () => {
      const filters = page.locator('input[placeholder*="filter" i]');
      const count = await filters.count();
      
      for (let i = 0; i < count; i++) {
        const filter = filters.nth(i);
        // Verify the header above this filter is Custom ID (not MRN or Patient MRN)
        const isCustomId = await filter.evaluate((el) => {
          const cell = el.closest('td, th');
          if (!cell) return false;
          const row = cell.closest('tr');
          if (!row) return false;
          const table = row.closest('table');
          if (!table) return false;
          const headerRows = table.querySelectorAll('tr');
          if (headerRows.length === 0) return false;
          
          // Find which column this filter is in
          const filterCells = Array.from(row.querySelectorAll('td, th'));
          const columnIndex = filterCells.indexOf(cell);
          if (columnIndex === -1) return false;
          
          // Check the header row for this column
          for (const headerRow of headerRows) {
            const headerCells = Array.from(headerRow.querySelectorAll('th, td, [role="columnheader"]'));
            if (columnIndex < headerCells.length) {
              const headerText = headerCells[columnIndex].textContent;
              // Must be exactly "Custom ID", not "Patient MRN" or "MRN"
              if (headerText && /^Custom ID$/i.test(headerText.trim())) {
                return true;
              }
            }
          }
          return false;
        });
        
        if (isCustomId) {
          return filter;
        }
      }
      throw new Error('Custom ID filter not found by header verification');
    },
    // Strategy 4: Find Custom ID header, then navigate to filter in same column using DOM structure
    async () => {
      const table = page.locator('table').first();
      await table.waitFor({ state: 'visible', timeout: 5000 });
      
      const customIdHeader = table.locator('th, [role="columnheader"]').filter({ hasText: /^Custom ID$/i }).first();
      await customIdHeader.waitFor({ state: 'visible', timeout: 5000 });
      
      // Get the column index
      const columnIndex = await customIdHeader.evaluate((el) => {
        const row = el.closest('tr');
        if (!row) return -1;
        const cells = Array.from(row.querySelectorAll('th, td, [role="columnheader"]'));
        return cells.indexOf(el);
      });
      
      if (columnIndex === -1) {
        throw new Error('Could not determine Custom ID column index');
      }
      
      // Find filter in the same column
      const filterRow = table.locator('tr').filter({ has: page.locator('input[placeholder*="filter" i]') }).first();
      const filters = filterRow.locator('input[placeholder*="filter" i]');
      return filters.nth(columnIndex);
    },
    // Strategy 5: Direct search for filter in cell containing "Custom ID" text (but verify it's not MRN)
    async () => {
      const filter = page.locator('td, th').filter({ hasText: /Custom ID/i }).locator('input[placeholder*="filter" i]').first();
      // Verify the header doesn't say MRN
      const headerText = await filter.evaluate((el) => {
        const cell = el.closest('td, th');
        if (!cell) return '';
        const row = cell.closest('tr');
        if (!row) return '';
        const table = row.closest('table');
        if (!table) return '';
        const headerRow = table.querySelector('tr');
        if (!headerRow) return '';
        const cells = Array.from(headerRow.querySelectorAll('th, td, [role="columnheader"]'));
        const cellIndex = Array.from(row.querySelectorAll('td, th')).indexOf(cell);
        if (cellIndex >= 0 && cellIndex < cells.length) {
          return cells[cellIndex].textContent || '';
        }
        return '';
      });
      
      if (headerText && /^Custom ID$/i.test(headerText.trim()) && !/MRN/i.test(headerText)) {
        return filter;
      }
      throw new Error('Filter is not under Custom ID header');
    },
  ];

  for (let i = 0; i < strategies.length; i++) {
    try {
      console.log(`🔍 Trying strategy ${i + 1}...`);
      const filter = await strategies[i]();
      await filter.waitFor({ state: 'visible', timeout: 5000 });
      await filter.click();
      console.log(`✅ Custom ID filter clicked using strategy ${i + 1}`);
      return;
    } catch (error) {
      console.log(`⚠️ Strategy ${i + 1} failed: ${error.message}`);
      if (i === strategies.length - 1) {
        throw new Error(`All strategies failed to find Custom ID filter. Last error: ${error.message}`);
      }
    }
  }
}

/* =======================
   Login Logic
======================= */
async function login(page) {
  try {
    console.log("🌐 Opening login page...");

    await page.goto(`${CONFIG.BASE_URL}/login`, {
      waitUntil: "domcontentloaded",
      timeout: CONFIG.TIMEOUTS.PAGE_LOAD,
    });

    if (page.isClosed()) {
      throw new Error("Page closed during navigation");
    }

    const loginContainer = page
      .locator(
        "div.MuiGrid-root.MuiGrid-container.MuiGrid-align-items-xs-center.MuiGrid-justify-content-xs-center"
      )
      .first();

    await loginContainer.waitFor({
      state: "visible",
      timeout: CONFIG.TIMEOUTS.ELEMENT_WAIT,
    });

    const emailInput = loginContainer.locator("input").nth(0);
    const passwordInput = loginContainer.locator("input").nth(1);

    await emailInput.fill(CONFIG.CREDS.email);
    
    // Click password field first to make it editable (removes readonly attribute)
    await passwordInput.click();
    await page.waitForTimeout(500); // Wait for field to become editable
    
    // Use type() instead of fill() for readonly fields, or remove readonly and fill
    await passwordInput.evaluate((el) => {
      el.removeAttribute('readonly');
    });
    await passwordInput.fill(CONFIG.CREDS.password);

    await page.getByRole("button", { name: "Login" }).click();
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);

    if (page.isClosed()) {
      throw new Error("Page closed after login submission");
    }

    console.log("✅ Login successful");
  } catch (error) {
    console.error("❌ Login failed:", error.message);
    throw error;
  }
}

/* =======================
   Session Bootstrap
======================= */
async function startQHSLabSession() {
  if (!CONFIG.CREDS.email || !CONFIG.CREDS.password) {
    throw new Error("Missing QHSLAB_EMAIL or QHSLAB_PASSWORD");
  }

  console.log("🚀 Launching browser...");

  const browser = await chromium.launch({
    headless: CONFIG.BROWSER.HEADLESS,
    slowMo: CONFIG.BROWSER.SLOW_MO,
    args: CONFIG.BROWSER.ARGS,
  });

  const context = await browser.newContext();
  const page = await context.newPage();

  attachPageListeners(page);
  await login(page);

  return { browser, context, page };
}

/* =======================
   Process Single Patient Claim
   Processes one patient/claim row from the table
======================= */
async function processPatientClaim(page, excelRow, patientIndex, totalPatients) {
  const startTime = Date.now();
  const result = {
    rowNumber: excelRow._rowNumber,
    mrn: excelRow.mrn,
    customId: excelRow.customId,
    dos: excelRow.dos,
    appointmentDate: excelRow.appointmentDate,
    cptCodes: excelRow.cptCodes || [],
    patientIndex: patientIndex + 1,
    totalPatients: totalPatients,
    status: 'Unknown',
    billingStatus: '', // Will be set based on Excel value
    errorMessage: '',
    generatedId: '',
    processingTime: 0,
    timestamp: new Date().toISOString(),
    notes: ''
  };

  try {
    console.log(`   📝 Processing patient ${result.patientIndex}/${totalPatients} for MRN ${excelRow.mrn}...`);

    // Find all table rows with "IntellyChart" text (these are the patient/claim rows)
    const table = page.locator('table').first();
    await table.waitFor({ state: 'visible', timeout: 5000 });
    
    // Get all data rows that contain "IntellyChart" (skip header and filter rows)
    const allRows = table.locator('tbody tr, tr[role="row"]')
      .filter({ hasNot: page.locator('input[placeholder*="filter" i]') })
      .filter({ has: page.locator('text=IntellyChart') });
    const rowCount = await allRows.count();
    
    if (rowCount === 0) {
      throw new Error('No patient rows with IntellyChart found in table');
    }

    if (patientIndex >= rowCount) {
      throw new Error(`Patient index ${patientIndex} exceeds available rows (${rowCount})`);
    }

    // Get the specific row for this patient
    const patientRow = allRows.nth(patientIndex);
    await patientRow.waitFor({ state: 'visible', timeout: 5000 });

    // Scroll the row into view if needed
    await patientRow.scrollIntoViewIfNeeded();
    await page.waitForTimeout(500);

    // Click on IntellyChart in this row to open the form
    const intellyChartLink = patientRow.getByText('IntellyChart').first();
    await intellyChartLink.waitFor({ state: 'visible', timeout: 5000 });
    await intellyChartLink.click();
    await page.waitForTimeout(1000);

    // Open the form (click the icon button) - try multiple strategies
    let formOpened = false;
    const formButtonStrategies = [
      () => page.locator('.MuiButtonBase-root-773.MuiIconButton-root-765.jss781'),
      () => page.locator('button[aria-label*="edit" i], button[aria-label*="open" i]').first(),
      () => page.locator('.MuiIconButton-root').first(),
      () => page.locator('button').filter({ has: page.locator('svg') }).first()
    ];
    
    for (const strategy of formButtonStrategies) {
      try {
        const button = strategy();
        if (await button.isVisible({ timeout: 3000 }).catch(() => false)) {
          await button.click();
          await page.waitForTimeout(1500);
          formOpened = true;
          console.log(`      ✅ Form opened successfully`);
          break;
        }
      } catch (e) {
        continue;
      }
    }
    
    if (!formOpened) {
      console.log(`      ⚠️  Form button not found, but continuing...`);
      await page.waitForTimeout(1000);
    }

    // Check billing status early to determine if we should skip DOS/ICD (for cancelled appointments)
    const excelBillingStatus = (
      excelRow.billingStatus || 
      excelRow['billing status'] || 
      excelRow.billingstatus ||
      excelRow['qhs billing status'] ||
      excelRow['qhsbillingstatus'] ||
      ''
    ).toString().trim();
    
    const isCancelledAppointment = excelBillingStatus.toUpperCase() !== 'QHSLAB';
    
    if (isCancelledAppointment) {
      console.log(`      ⏭️  Appointment is being cancelled (billing status: "${excelBillingStatus || '(empty)'}") - skipping DOS and ICD entry`);
      console.log(`      ⏭️  Proceeding directly to billing status selection...`);
    } else {
      // Fill Date of Service (DOS) from Excel
    const dosValue = excelRow.dosFormatted || excelRow.dos || '';
    if (dosValue) {
      try {
        console.log(`      📅 Entering Date of Service from Excel: ${dosValue}`);
        
        // Find DOS field
        let dosField = null;
        
        // Strategy 3: Use exact selector provided
        if (!dosField) {
          try {
            const field = page.locator('.jss2486 > form > .jss2315 > .MuiFormControl-root-2238 > .jss2312 > .MuiInputBase-root-2266 > .MuiInputBase-input-2274').first();
            if (await field.isVisible({ timeout: 3000 }).catch(() => false)) {
              dosField = field;
              console.log(`      ✅ Found Date of Service field using exact selector`);
            }
          } catch (e) {
            // Continue
          }
        }
        
        // Strategy 4: Find by label "Date of Service"
        if (!dosField) {
          try {
            const field = page.locator('label').filter({ hasText: /Date of Service/i }).locator('..').locator('input').first();
            if (await field.isVisible({ timeout: 3000 }).catch(() => false)) {
              dosField = field;
              console.log(`      ✅ Found Date of Service field by label`);
            }
          } catch (e) {
            // Continue
          }
        }
        
        if (!dosField) {
          throw new Error('Date of Service field not found');
        }
        
        await dosField.waitFor({ state: 'visible', timeout: 10000 });
        await dosField.scrollIntoViewIfNeeded();
        await dosField.click();
        await page.waitForTimeout(500);
        
        // Clear the field first
        await dosField.fill('');
        await page.waitForTimeout(300);
        
        // Enter the DOS value from Excel
        await dosField.fill(dosValue);
        await page.waitForTimeout(500);
        
        // Press Enter or Tab to confirm
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);
        
        console.log(`      ✅ Date of Service entered: ${dosValue}`);
      } catch (error) {
        console.log(`      ⚠️  Date of Service entry failed: ${error.message}. Continuing...`);
        // Continue even if DOS entry fails
      }
    } else {
      console.log(`      ℹ️  No Date of Service provided in Excel row for this patient`);
    }

    // Fill ICD field from Excel (skip one field after DOS, then ICD)
    const icdValue = excelRow.icd || excelRow['icd code'] || excelRow['icd codes'] || '';
    if (icdValue) {
      try {
        console.log(`      📋 Entering ICD from Excel: ${icdValue}`);
        
        // Find ICD field - it's one field after DOS (skip one field)
        let icdField = null;
        
        // Strategy 1: Find DOS field first, then skip one field and get the next one (ICD)
        try {
          const allInputs = page.locator('form input, form .MuiInputBase-input');
          const inputCount = await allInputs.count();
          
          for (let i = 0; i < inputCount; i++) {
            const input = allInputs.nth(i);
            const isDos = await input.evaluate((el) => {
              const label = el.closest('.MuiFormControl-root')?.querySelector('label');
              return label?.textContent?.includes('Date of Service') || label?.textContent?.includes('DOS');
            });
            
            if (isDos && i + 2 < inputCount) {
              // Skip one field (i+1) and get the next one (i+2) which should be ICD
              icdField = allInputs.nth(i + 2);
              console.log(`      ✅ Found ICD field (skipped one field after DOS)`);
              break;
            }
          }
        } catch (e) {
          // Continue to other strategies
        }
        
        // Strategy 2: Find by label "ICD"
        if (!icdField) {
          try {
            const field = page.locator('label').filter({ hasText: /^ICD/i }).locator('..').locator('input').first();
            if (await field.isVisible({ timeout: 3000 }).catch(() => false)) {
              icdField = field;
              console.log(`      ✅ Found ICD field by label`);
            }
          } catch (e) {
            // Continue
          }
        }
        
        // Strategy 3: Find input near "ICD" text
        if (!icdField) {
          try {
            const field = page.locator('div, label').filter({ hasText: /^ICD/i }).locator('input').first();
            if (await field.isVisible({ timeout: 3000 }).catch(() => false)) {
              icdField = field;
              console.log(`      ✅ Found ICD field near ICD text`);
            }
          } catch (e) {
            // Continue
          }
        }
        
        if (!icdField) {
          throw new Error('ICD field not found');
        }
        
        await icdField.waitFor({ state: 'visible', timeout: 10000 });
        await icdField.scrollIntoViewIfNeeded();
        await icdField.click();
        await page.waitForTimeout(1500); // Wait for dropdown/search modal to open
        
        // The ICD field is readonly, so we need to find the search input field in the modal/popover
        // Look for the editable search input that appears in the dropdown
        let searchInput = null;
        
        // Strategy 1: Find input with "Search by ICD code" placeholder or label
        try {
          const input1 = page.locator('input').filter({ hasNotText: /readonly/i }).first();
          const isReadonly = await input1.evaluate((el) => el.hasAttribute('readonly') || el.readOnly);
          if (!isReadonly) {
            const placeholder = await input1.getAttribute('placeholder').catch(() => '');
            const label = await input1.evaluate((el) => {
              return el.closest('div')?.querySelector('label')?.textContent || '';
            }).catch(() => '');
            if (placeholder?.toLowerCase().includes('search') || label?.toLowerCase().includes('search') || 
                placeholder?.toLowerCase().includes('icd') || label?.toLowerCase().includes('icd')) {
              if (await input1.isVisible({ timeout: 2000 }).catch(() => false)) {
                searchInput = input1;
                console.log(`      ✅ Found ICD search input by placeholder/label`);
              }
            }
          }
        } catch (e) {
          // Continue
        }
        
        // Strategy 2: Find all inputs and check which one is NOT readonly and is in a modal/popover
        if (!searchInput) {
          try {
            const allInputs = page.locator('input[type="text"]');
            const inputCount = await allInputs.count();
            
            for (let i = 0; i < inputCount; i++) {
              const input = allInputs.nth(i);
              const isReadonly = await input.evaluate((el) => el.hasAttribute('readonly') || el.readOnly).catch(() => true);
              if (!isReadonly) {
                const isVisible = await input.isVisible({ timeout: 1000 }).catch(() => false);
                if (isVisible) {
                  // Check if it's in a modal/popover/dialog
                  const inModal = await input.evaluate((el) => {
                    return !!el.closest('[role="dialog"], [role="menu"], .MuiPopover-root, .MuiModal-root, [class*="popover"], [class*="modal"]');
                  }).catch(() => false);
                  
                  if (inModal) {
                    searchInput = input;
                    console.log(`      ✅ Found editable ICD search input in modal`);
                    break;
                  }
                }
              }
            }
          } catch (e) {
            // Continue
          }
        }
        
        // Strategy 3: Find by placeholder text "Search by ICD code"
        if (!searchInput) {
          try {
            const input = page.getByPlaceholder(/Search.*ICD/i).first();
            if (await input.isVisible({ timeout: 2000 }).catch(() => false)) {
              const isReadonly = await input.evaluate((el) => el.hasAttribute('readonly') || el.readOnly).catch(() => true);
              if (!isReadonly) {
                searchInput = input;
                console.log(`      ✅ Found ICD search input by placeholder`);
              }
            }
          } catch (e) {
            // Continue
          }
        }
        
        // Strategy 4: Find input that becomes visible after clicking ICD field (in popover/modal)
        if (!searchInput) {
          try {
            // Wait a bit more for modal to fully render
            await page.waitForTimeout(500);
            const inputs = page.locator('input').filter({ hasNot: page.locator('[readonly]') });
            const inputCount = await inputs.count();
            
            for (let i = 0; i < inputCount; i++) {
              const input = inputs.nth(i);
              const isVisible = await input.isVisible({ timeout: 1000 }).catch(() => false);
              if (isVisible) {
                // Double check it's not readonly
                const isReadonly = await input.evaluate((el) => el.hasAttribute('readonly') || el.readOnly).catch(() => true);
                if (!isReadonly) {
                  searchInput = input;
                  console.log(`      ✅ Found editable input after ICD click`);
                  break;
                }
              }
            }
          } catch (e) {
            // Continue
          }
        }
        
        if (searchInput) {
          // Use MuiAutocomplete-option method to select from dropdown
          try {
            console.log(`      🔍 Opening ICD dropdown and selecting option...`);
            
            // Open ICD dropdown - click the search input we already found
            await searchInput.click();
            await page.waitForTimeout(500);
            
            // Fill with ICD value
            const icdPartial = icdValue.split('.')[0]; // Get part before decimal if exists
            await searchInput.fill(icdPartial || icdValue);
            await page.waitForTimeout(1000);
            
            // Press Enter to trigger dropdown
            await page.keyboard.press('Enter');
            await page.waitForTimeout(3000); // Wait longer for dropdown to appear
            
            // Take screenshot after pressing Enter to see the box
            try {
              await page.screenshot({ path: 'icd_dropdown_after_enter.png', fullPage: false });
              console.log(`      📸 Screenshot saved: icd_dropdown_after_enter.png`);
            } catch (screenshotError) {
              console.log(`      ⚠️  Could not take screenshot: ${screenshotError.message}`);
            }
            
            // Find the ICD dropdown container first (the popover/modal that appears)
            await page.waitForTimeout(2000); // Wait for dropdown to fully render
            
            // Find the dropdown container - try multiple strategies
            let dropdownContainer = null;
            
            // Strategy 1: Look for listbox with ICD code options
            const dropdownContainerSelectors = [
              '[role="listbox"]',
              '.MuiAutocomplete-paper',
              '.MuiPopover-paper',
              '[class*="Popover"]',
              '[class*="Autocomplete"]',
              '.MuiPaper-root'
            ];
            
            for (const selector of dropdownContainerSelectors) {
              try {
                // First try with filter
                const container = page.locator(selector).filter({ has: page.locator('li, [role="option"]') }).first();
                if (await container.isVisible({ timeout: 3000 }).catch(() => false)) {
                  dropdownContainer = container;
                  console.log(`      ✅ Found ICD dropdown container using selector "${selector}" (with filter)`);
                  break;
                }
              } catch (e) {
                // Try without filter
                try {
                  const container = page.locator(selector).first();
                  if (await container.isVisible({ timeout: 2000 }).catch(() => false)) {
                    // Check if it contains list items
                    const hasOptions = await container.locator('li, [role="option"]').count().catch(() => 0);
                    if (hasOptions > 0) {
                      dropdownContainer = container;
                      console.log(`      ✅ Found ICD dropdown container using selector "${selector}" (without filter)`);
                      break;
                    }
                  }
                } catch (e2) {
                  continue;
                }
              }
            }
            
            // Strategy 2: If container not found, try to find options directly and get their parent
            if (!dropdownContainer) {
              console.log(`      🔍 Container not found, trying to find options directly...`);
              const directOptions = page.locator('li[role="option"], li.MuiMenuItem-root, .MuiAutocomplete-option').first();
              const isVisible = await directOptions.isVisible({ timeout: 3000 }).catch(() => false);
              if (isVisible) {
                // Get the parent container
                dropdownContainer = directOptions.locator('..').locator('..');
                console.log(`      ✅ Found dropdown container by finding options first`);
              }
            }
            
            // Strategy 3: Use page as container if nothing else works
            if (!dropdownContainer) {
              console.log(`      ⚠️  Dropdown container not found, will search page-wide for options`);
              dropdownContainer = page; // Use page as container
            }
            
            // Find options WITHIN the dropdown container only - prioritize visible options
            const optionSelectors = [
              'li[role="option"]',
              'li.MuiMenuItem-root',
              '.MuiAutocomplete-option',
              'li',
              '[role="option"]'
            ];
            
            let allOptions = null;
            
            // Try each selector within the container
            for (const selector of optionSelectors) {
              try {
                const options = dropdownContainer.locator(selector);
                const count = await options.count();
                if (count > 0) {
                  // Filter to only include VISIBLE options that look like ICD codes
                  let visibleIcdOptions = [];
                  for (let i = 0; i < count; i++) {
                    const option = options.nth(i);
                    const isVisible = await option.isVisible({ timeout: 500 }).catch(() => false);
                    if (isVisible) {
                      const optionText = await option.textContent().catch(() => '');
                      const trimmedText = optionText.trim();
                      // Check if it looks like an ICD code (e.g., "F33.41", "F33.42", etc.)
                      if (trimmedText && /^[A-Z]\d{2,3}\.?\d*/.test(trimmedText)) {
                        visibleIcdOptions.push(i);
                      }
                    }
                  }
                  
                  if (visibleIcdOptions.length > 0) {
                    // Create a filtered locator for visible ICD options
                    // We'll use the indices to access them
                    allOptions = {
                      locator: options,
                      indices: visibleIcdOptions
                    };
                    console.log(`      📋 Found ${visibleIcdOptions.length} visible ICD code option(s) using selector "${selector}"`);
                    break;
                  }
                }
              } catch (e) {
                continue;
              }
            }
            
            // If no visible ICD options found, try to find any visible options
            if (!allOptions) {
              console.log(`      🔍 Trying to find visible options without ICD pattern filter...`);
              for (const selector of optionSelectors) {
                try {
                  const options = dropdownContainer.locator(selector);
                  const count = await options.count();
                  if (count > 0) {
                    // Find visible options
                    let visibleIndices = [];
                    for (let i = 0; i < count; i++) {
                      const option = options.nth(i);
                      const isVisible = await option.isVisible({ timeout: 500 }).catch(() => false);
                      if (isVisible) {
                        visibleIndices.push(i);
                      }
                    }
                    
                    if (visibleIndices.length > 0) {
                      allOptions = {
                        locator: options,
                        indices: visibleIndices
                      };
                      console.log(`      📋 Found ${visibleIndices.length} visible option(s) using selector "${selector}"`);
                      break;
                    }
                  }
                } catch (e) {
                  continue;
                }
              }
            }
            
            if (!allOptions || !allOptions.indices || allOptions.indices.length === 0) {
              throw new Error('No visible dropdown options found in ICD dropdown container');
            }
            
            const optionCount = allOptions.indices.length;
            const optionsLocator = allOptions.locator;
            console.log(`      🔍 Looking for option that starts with "${icdValue}" among ${optionCount} visible options...`);
            
            // Log all visible options for debugging
            console.log(`      📝 Available visible ICD options:`);
            for (let i = 0; i < Math.min(optionCount, 10); i++) {
              try {
                const actualIndex = allOptions.indices[i];
                const option = optionsLocator.nth(actualIndex);
                const optionText = await option.textContent().catch(() => '');
                console.log(`         ${i} (actual index ${actualIndex}): "${optionText.trim().substring(0, 80)}"`);
              } catch (e) {
                console.log(`         ${i}: (error reading option)`);
              }
            }
            
            let matchingOption = null;
            let matchingActualIndex = -1;
            
            // Find the first visible option that starts with the Excel ICD code (and looks like an ICD code)
            for (let i = 0; i < optionCount; i++) {
              try {
                const actualIndex = allOptions.indices[i];
                const option = optionsLocator.nth(actualIndex);
                const optionText = await option.textContent().catch(() => '');
                const trimmedText = optionText.trim();
                
                // Only consider options that look like ICD codes (start with letter + numbers)
                if (trimmedText && /^[A-Z]\d{2,3}\.?\d*/.test(trimmedText)) {
                  // Check if option text starts with the Excel ICD code (case-insensitive)
                  if (trimmedText.toUpperCase().startsWith(icdValue.toUpperCase())) {
                    matchingOption = option;
                    matchingActualIndex = actualIndex;
                    console.log(`      ✅ Found matching visible option at index ${i} (actual ${actualIndex}): "${trimmedText.substring(0, 80)}..."`);
                    break;
                  }
                }
              } catch (e) {
                continue;
              }
            }
            
            // If no exact match, try with prefix (e.g., "F33" for "F33.1")
            if (!matchingOption && icdValue.includes('.')) {
              const icdPrefix = icdValue.split('.')[0];
              console.log(`      🔍 No exact match, trying prefix "${icdPrefix}"...`);
              for (let i = 0; i < optionCount; i++) {
                try {
                  const actualIndex = allOptions.indices[i];
                  const option = optionsLocator.nth(actualIndex);
                  const optionText = await option.textContent().catch(() => '');
                  const trimmedText = optionText.trim();
                  
                  // Only consider options that look like ICD codes
                  if (trimmedText && /^[A-Z]\d{2,3}\.?\d*/.test(trimmedText)) {
                    // Check if option starts with ICD prefix (e.g., "F33.4" starts with "F33")
                    if (trimmedText.toUpperCase().startsWith(icdPrefix.toUpperCase())) {
                      matchingOption = option;
                      matchingActualIndex = actualIndex;
                      console.log(`      ✅ Found matching visible option by prefix at index ${i} (actual ${actualIndex}): "${trimmedText.substring(0, 80)}..."`);
                      break;
                    }
                  }
                } catch (e) {
                  continue;
                }
              }
            }
            
            // If still no match, find first visible option that looks like an ICD code
            if (!matchingOption) {
              for (let i = 0; i < optionCount; i++) {
                try {
                  const actualIndex = allOptions.indices[i];
                  const option = optionsLocator.nth(actualIndex);
                  const optionText = await option.textContent().catch(() => '');
                  const trimmedText = optionText.trim();
                  
                  // Find first option that looks like an ICD code
                  if (trimmedText && /^[A-Z]\d{2,3}\.?\d*/.test(trimmedText)) {
                    matchingOption = option;
                    matchingActualIndex = actualIndex;
                    console.log(`      ⚠️  No exact match, using first visible ICD option at index ${i} (actual ${actualIndex}): "${trimmedText.substring(0, 80)}..."`);
                    break;
                  }
                } catch (e) {
                  continue;
                }
              }
            }
            
            // Final fallback: use first visible option if no ICD code found
            if (!matchingOption && optionCount > 0) {
              const actualIndex = allOptions.indices[0];
              matchingOption = optionsLocator.nth(actualIndex);
              matchingActualIndex = actualIndex;
              const optionText = await matchingOption.textContent().catch(() => '');
              console.log(`      ⚠️  No ICD code option found, using first visible option at index 0 (actual ${actualIndex}): "${optionText.trim().substring(0, 80)}..."`);
            }
            
            // Click the matching option
            if (matchingOption) {
              try {
                // Scroll option into view if needed
                await matchingOption.scrollIntoViewIfNeeded();
                await page.waitForTimeout(500);
                
                // Click the option
                await matchingOption.click({ timeout: 5000 });
                await page.waitForTimeout(1500); // Wait for selection to complete
                
                console.log(`      ✅ Selected dropdown option that starts with "${icdValue}"`);
              } catch (clickError) {
                console.log(`      ⚠️  Click failed: ${clickError.message}, trying force click...`);
                await matchingOption.click({ timeout: 5000, force: true });
                await page.waitForTimeout(1500);
                console.log(`      ✅ Selected dropdown option (force click)`);
              }
            } else {
              throw new Error('Could not find any dropdown option to select');
            }
          } catch (optionError) {
            console.log(`      ⚠️  Option selection failed: ${optionError.message}`);
            throw optionError;
          }
        } else {
          // Fallback: Search input not found
          throw new Error('ICD search input field not found - cannot enter ICD value');
        }
      } catch (error) {
        console.log(`      ⚠️  ICD entry failed: ${error.message}. Continuing...`);
        // Continue even if ICD entry fails
      }
    } else {
      console.log(`      ℹ️  No ICD provided in Excel row for this patient`);
    }
    } // End of if (!isCancelledAppointment) - DOS and ICD sections

    // Check if page and form are still open before setting billing status
    if (page.isClosed()) {
      throw new Error('Page closed before setting billing status');
    }
    
    // Check if form is still open by looking for Save button
    const formStillOpen = await page.getByRole('button', { name: 'Save' }).isVisible({ timeout: 2000 }).catch(() => false);
    if (!formStillOpen) {
      console.log(`      ⚠️  Form appears to be closed. Attempting to continue anyway...`);
    }
    
    // Set Billing Status based on Excel value (already checked earlier)
    // If Excel has "QHSLAB", mark as "Billed", otherwise mark as "Appt. Cancelled"
    console.log(`      🔍 Setting billing status from Excel...`);
    console.log(`      📊 Billing Status read from Excel: "${excelBillingStatus || '(empty)'}"`);
    console.log(`      📊 Billing Status (uppercase): "${excelBillingStatus.toUpperCase()}"`);
    console.log(`      📊 Is it QHSLAB? ${excelBillingStatus.toUpperCase() === 'QHSLAB'}`);
    
    let targetBillingStatus;
    let statusLabel; // For result status field
    
    if (excelBillingStatus.toUpperCase() === 'QHSLAB') {
      targetBillingStatus = 'Billed';
      statusLabel = 'Billed';
      console.log(`      ✅ Setting Billing Status to: "Billed" (Excel had "QHSLAB")`);
    } else if (excelBillingStatus) {
      targetBillingStatus = 'Appt. Cancelled';
      statusLabel = 'Appt. Cancelled';
      console.log(`      ⚠️  Setting Billing Status to: "Appt. Cancelled" (Excel had "${excelBillingStatus}", not "QHSLAB")`);
    } else {
      // If no billing status in Excel, default to "Appt. Cancelled" for safety
      targetBillingStatus = 'Appt. Cancelled';
      statusLabel = 'Appt. Cancelled';
      console.log(`      ℹ️  Setting Billing Status to: "Appt. Cancelled" (no value in Excel)`);
    }
    
    // Find Billing Status field using multiple strategies
    let billingStatusField = null;
    const billingStatusStrategies = [
      () => page.getByRole('textbox', { name: 'Billing Status' }),
      () => page.locator('label').filter({ hasText: /Billing Status/i }).locator('..').locator('input').first(),
      () => page.locator('input[placeholder*="Billing Status" i]').first(),
      () => page.locator('input').filter({ has: page.locator('label:has-text("Billing Status")') }).first(),
      () => page.locator('form input').filter({ has: page.locator('label:has-text("Billing")') }).first()
    ];
    
    for (const strategy of billingStatusStrategies) {
      try {
        const field = strategy();
        if (await field.isVisible({ timeout: 2000 }).catch(() => false)) {
          billingStatusField = field;
          console.log(`      ✅ Found Billing Status field using strategy ${billingStatusStrategies.indexOf(strategy) + 1}`);
          break;
        }
      } catch (e) {
        continue;
      }
    }
    
    if (!billingStatusField) {
      throw new Error('Could not find Billing Status field');
    }
    
    if (page.isClosed()) {
      throw new Error('Page closed before clicking Billing Status field');
    }
    
    await billingStatusField.click();
    await page.waitForTimeout(500);
    
    if (page.isClosed()) {
      throw new Error('Page closed after clicking Billing Status field');
    }
    
    let billingStatusSet = false;
    try {
      const menuItem = page.getByRole('menuitem', { name: targetBillingStatus });
      await menuItem.waitFor({ state: 'visible', timeout: 5000 });
      
      if (page.isClosed()) {
        throw new Error('Page closed before clicking billing status menu item');
      }
      
      await menuItem.click();
      billingStatusSet = true;
      console.log(`      ✅ Billing Status set to: "${targetBillingStatus}"`);
    } catch (error) {
      console.log(`      ⚠️  Billing Status "${targetBillingStatus}" not found in menu. Available options may have changed.`);
      // Try to find and click any visible billing status option as fallback
      try {
        const firstOption = page.getByRole('menuitem').first();
        if (await firstOption.isVisible({ timeout: 2000 }).catch(() => false)) {
          await firstOption.click();
          console.log(`      ⚠️  Selected first available billing status option as fallback`);
          billingStatusSet = true;
        }
      } catch (fallbackError) {
        console.log(`      ❌ Could not set billing status: ${fallbackError.message}`);
      }
    }
    
    // Update result with billing status information
    result.billingStatus = targetBillingStatus;

    // Save
    if (page.isClosed()) {
      throw new Error('Page closed before clicking Save button');
    }
    
    console.log(`      💾 Clicking Save button...`);
    try {
      await page.getByRole('button', { name: 'Save' }).click({ timeout: 10000 });
      await page.waitForTimeout(3000); // Wait longer for save to complete
    } catch (saveError) {
      if (page.isClosed()) {
        throw new Error('Page closed during save');
      }
      throw saveError;
    }
    
    // Wait for form/modal to close after save
    try {
      // Wait for the form/modal to disappear (indicating save was successful)
      const formStillOpen = await page.getByRole('button', { name: 'Save' }).isVisible({ timeout: 2000 }).catch(() => false);
      if (formStillOpen) {
        // Form still open, try to close it
        const closeBtn = page.getByRole('button', { name: /close|cancel/i }).first();
        if (await closeBtn.isVisible({ timeout: 1000 }).catch(() => false)) {
          await closeBtn.click();
          await page.waitForTimeout(1000);
        }
      } else {
        console.log(`      ✅ Form closed after save - ready for next patient`);
      }
    } catch (error) {
      // Form may have closed automatically
      console.log(`      ✅ Form closed automatically after save`);
    }
    
    // After save, we're automatically back on the Claims page - no navigation needed
    await page.waitForTimeout(2000);
    
    // Check if page is still open
    if (page.isClosed()) {
      throw new Error('Page closed after save');
    }
    
    // Wait for table to be visible and ready (we're already on Claims page)
    try {
      const table = page.locator('table').first();
      await table.waitFor({ state: 'visible', timeout: 10000 });
      console.log(`      ✅ Claims page is ready - table visible`);
    } catch (error) {
      console.log(`      ⚠️  Table not found: ${error.message}`);
    }
    
    // Note: For patient 2 onwards, filters (Custom ID, MRN, Appointment Date) will be updated
    // in processRow function - we just remove previous value and add new value
    // Entity Source and Assessment/Resource filters are never touched

    // Update status based on billing status that was set
    if (billingStatusSet) {
      if (result.billingStatus === 'Billed') {
        result.status = 'Billed';
        result.notes = `Patient ${result.patientIndex}/${totalPatients} processed successfully - Marked as Bill ed`;
      } else if (result.billingStatus === 'Appt. Cancelled') {
        result.status = 'Appt. Cancelled';
        result.notes = `Patient ${result.patientIndex}/${totalPatients} processed successfully - Marked as Appt. Cancelled`;
      } else {
        result.status = 'Success';
        result.notes = `Patient ${result.patientIndex}/${totalPatients} processed successfully`;
      }
    } else {
      result.status = 'Success';
      result.notes = `Patient ${result.patientIndex}/${totalPatients} processed successfully (billing status not set)`;
    }
    console.log(`      ✅ Patient ${result.patientIndex}/${totalPatients} processed successfully - Status: ${result.status}`);

  } catch (error) {
    result.status = 'Failed';
    result.errorMessage = error.message;
    console.error(`      ❌ Patient ${result.patientIndex}/${totalPatients} failed: ${error.message}`);
    
    // Try to close any open modals/forms
    try {
      const closeBtn = page.getByRole('button', { name: /close|cancel|×/i }).first();
      if (await closeBtn.isVisible({ timeout: 1000 }).catch(() => false)) {
        await closeBtn.click();
        await page.waitForTimeout(500);
      }
    } catch (closeError) {
      // Ignore close errors
    }
  } finally {
    result.processingTime = Date.now() - startTime;
  }

  return result;
}

/* =======================
   Process Single Row (Excel Row)
   Filters by MRN and processes all matching patients
======================= */
async function processRow(page, row, rowIndex) {
  const startTime = Date.now();
  const allResults = [];

  try {
    console.log(`\n📋 Processing Excel Row ${row._rowNumber || rowIndex + 1}: MRN=${row.mrn} (processing all patients with this MRN)`);

    if (CONFIG.DRY_RUN || !page) {
      console.log(`   [DRY RUN] Would process: MRN=${row.mrn}, DOS=${row.dosFormatted || row.dos}, CPT=${row.cptCodes?.join(', ') || 'N/A'}`);
      allResults.push({
        rowNumber: row._rowNumber || rowIndex + 1,
        mrn: row.mrn,
        customId: row.customId,
        dos: row.dos,
        appointmentDate: row.appointmentDate,
        cptCodes: row.cptCodes || [],
        patientIndex: 0,
        totalPatients: 0,
        status: 'Skipped (Dry Run)',
        errorMessage: '',
        generatedId: '',
        processingTime: 0,
        timestamp: new Date().toISOString(),
        notes: 'Dry run mode - no automation executed'
      });
      return allResults;
    }

    // Check if page is still open
    if (page.isClosed()) {
      throw new Error('Page was closed unexpectedly. Cannot continue processing.');
    }

    // Navigate to Accounts page only once (first time) - after that, we stay on Claims page
    if (!accountsPageNavigated) {
      console.log("🌐 Navigating to Accounts page (first time only)...");
      try {
      await page.goto(
        `${CONFIG.BASE_URL}/6oQ5FvCBDUC5CiIrutgARg/accounts`,
          { waitUntil: "domcontentloaded", timeout: 30000 }
        );
        await page.waitForTimeout(2000); // Wait for page to fully load
        accountsPageNavigated = true;
      } catch (error) {
        throw new Error(`Failed to navigate to Accounts page: ${error.message}`);
      }
    } else {
      // Already navigated - we're already on Claims page after save, no need to navigate
      console.log("ℹ️  Already on Claims page - no navigation needed");
    }

    // Click Operations and Claims only once per session
    if (!claimsClickedOnce) {
      try {
    await page.locator('div').filter({ hasText: /^Operations$/ }).nth(1).click();
        await page.waitForTimeout(500);
    await page.locator('div').filter({ hasText: /^Operations$/ }).first().click();
        await page.waitForTimeout(500);
    await page.locator('div').filter({ hasText: /^Operations$/ }).first().click();
        await page.waitForTimeout(500);
    await clickClaimsOnce(page);
        await page.waitForTimeout(1000);
      } catch (error) {
        console.log(`   ⚠️  Operations/Claims click issue: ${error.message}. Continuing (already on Claims page)...`);
        // We're already on Claims page after save, so just continue
      }
    } else {
      // Already clicked Claims, just wait a bit for page to be ready
      await page.waitForTimeout(1000);
    }

    // Apply Entity Source filter (IntellyChart) only once at the start
    // Note: For patient 2 onwards, we only update Custom ID, MRN, and Appointment Date filters
    // We don't clear or touch Entity Source and Assessment/Resource filters
    if (!entitySourceFilterApplied) {
      try {
        console.log(`   🔍 Applying Entity Source filter (IntellyChart) - one time only`);
        await page.getByRole('columnheader', { name: 'Entity Source Sort by Entity' }).getByLabel('All').click();
        await page.waitForTimeout(1000);
        await page.getByRole('option', { name: 'IntellyChart' }).getByRole('checkbox').check();
        await page.waitForTimeout(500);
        await page.getByRole('button', { name: 'Apply Filter' }).click();
        await page.waitForTimeout(2000); // Wait for filter to apply
        entitySourceFilterApplied = true;
        console.log(`   ✅ Entity Source filter applied: IntellyChart (one-time)`);
      } catch (error) {
        console.log(`   ⚠️  Entity Source filter failed: ${error.message}. Continuing...`);
        // Continue even if Entity Source filter fails
      }
    }

    // Apply Assessment/Resource filter only once at the start
    if (!assessmentResourceFilterApplied) {
      try {
        console.log(`   🔍 Applying Assessment/Resource filter - one time only`);
        await page.getByRole('columnheader', { name: 'Assessment/Resource Sort by' }).getByPlaceholder('filter').click();
        await page.waitForTimeout(1000);
        
        // Find and select "Health Assessment" from the available options
        let healthAssessmentFound = false;
        const strategies = [
          // Strategy 1: Find by role and text
          async () => {
            const option = page.getByRole('option', { name: /Health Assessment/i }).first();
            if (await option.isVisible({ timeout: 3000 }).catch(() => false)) {
              const checkbox = option.getByRole('checkbox').first();
              await checkbox.check();
              healthAssessmentFound = true;
              return true;
            }
            return false;
          },
          // Strategy 2: Find by text in list items
          async () => {
            const listItems = page.locator('li').filter({ hasText: /Health Assessment/i });
            const count = await listItems.count();
            if (count > 0) {
              const firstItem = listItems.first();
              // Try to find checkbox within the item
              const checkbox = firstItem.locator('input[type="checkbox"], [role="checkbox"]').first();
              if (await checkbox.isVisible({ timeout: 2000 }).catch(() => false)) {
                await checkbox.check();
                healthAssessmentFound = true;
                return true;
              }
              // If no checkbox, try clicking the item itself
              await firstItem.click();
              healthAssessmentFound = true;
              return true;
            }
            return false;
          },
          // Strategy 3: Find all checkboxes and match by nearby text
          async () => {
            const checkboxes = page.locator('input[type="checkbox"], [role="checkbox"]');
            const checkboxCount = await checkboxes.count();
            for (let i = 0; i < checkboxCount; i++) {
              const checkbox = checkboxes.nth(i);
              const isVisible = await checkbox.isVisible({ timeout: 500 }).catch(() => false);
              if (isVisible) {
                // Check if nearby text contains "Health Assessment"
                const parentText = await checkbox.evaluate((el) => {
                  const parent = el.closest('li, div, label');
                  return parent?.textContent || '';
                }).catch(() => '');
                if (/Health Assessment/i.test(parentText)) {
                  await checkbox.check();
                  healthAssessmentFound = true;
                  return true;
                }
              }
            }
            return false;
          }
        ];
        
        for (let i = 0; i < strategies.length; i++) {
          try {
            const success = await strategies[i]();
            if (success) {
              console.log(`   ✅ Found and selected "Health Assessment" using strategy ${i + 1}`);
              break;
            }
          } catch (e) {
            continue;
          }
        }
        
        if (!healthAssessmentFound) {
          throw new Error('Could not find "Health Assessment" option in the filter menu');
        }
        
        await page.waitForTimeout(500);
        await page.getByRole('button', { name: 'Apply' }).click();
        await page.waitForTimeout(2000); // Wait for filter to apply
        assessmentResourceFilterApplied = true;
        console.log(`   ✅ Assessment/Resource filter applied (one-time)`);
      } catch (error) {
        console.log(`   ⚠️  Assessment/Resource filter failed: ${error.message}. Continuing...`);
        // Continue even if Assessment/Resource filter fails
      }
    }

    // Update filters for this patient (patient 2 onwards: remove previous, add new)
    // Note: We only update Custom ID, MRN, and Appointment Date filters
    // Entity Source and Assessment/Resource filters are never touched
    
    // Step 1: Update Custom ID filter (remove previous, add new)
    if (row.customId && row.customId.toString().trim()) {
      try {
        console.log(`   🔍 Step 1: Updating Custom ID filter: ${row.customId} (removing previous, adding new)`);
        await clickCustomIdFilter(page);
        const customIdFilter = page.getByRole('columnheader', { name: /Custom ID/i }).getByPlaceholder(/filter/i).first();
        await customIdFilter.waitFor({ state: 'visible', timeout: 5000 });
        await customIdFilter.click();
        await customIdFilter.fill(''); // Remove previous value
        await page.waitForTimeout(500);
        await customIdFilter.fill(row.customId.toString().trim()); // Add new value
        await page.waitForTimeout(2000); // Wait for table to filter
        console.log(`   ✅ Custom ID filter updated: ${row.customId}`);
      } catch (error) {
        console.log(`   ⚠️  Custom ID filter failed: ${error.message}. Continuing with MRN filter...`);
        // Continue with MRN filter even if Custom ID fails
      }
    } else {
      console.log(`   ℹ️  No Custom ID provided in Excel row, skipping Custom ID filter`);
    }

    // Step 2: Update Patient MRN filter (remove previous, add new)
    console.log(`   🔍 Step 2: Updating MRN filter: ${row.mrn} (removing previous, adding new)`);
    try {
      const mrnFilter = page.getByRole('columnheader', { name: /Patient MRN/i }).getByPlaceholder(/filter/i).first();
      await mrnFilter.waitFor({ state: 'visible', timeout: 5000 });
    await mrnFilter.click();
      await mrnFilter.fill(''); // Remove previous value
      await page.waitForTimeout(500);
    await mrnFilter.fill(row.mrn || ''); // Add new value
      await page.waitForTimeout(2000); // Wait for table to filter and show all matching patients
      console.log(`   ✅ MRN filter updated: ${row.mrn}`);
    } catch (error) {
      throw new Error(`Failed to apply MRN filter: ${error.message}`);
    }

    // Step 3: Update Appointment Date filter (remove previous, add new) (if provided in Excel)
    // Check multiple possible sources for appointment date
    const appointmentDateValue = row.appointmentDateFormatted || 
                                 row['orig appt. date'] || 
                                 row['orig appt date'] || 
                                 row['original appointment date'] ||
                                 (row.appointmentDate ? formatDate(row.appointmentDate) : null);
    
    if (appointmentDateValue) {
      try {
        console.log(`   📅 Step 3: Updating Appointment Date filter: ${appointmentDateValue} (removing previous, adding new)`);
        
        // Wait a bit more after MRN filter to ensure page is stable
      await page.waitForTimeout(2000);
        
        // Try multiple strategies to find the appointment date filter button
        let appointmentDateButton = null;
        const strategies = [
          // Strategy 1: Direct selector
          () => page.locator('.MuiInputBase-input.jss1013'),
          // Strategy 2: Find by text "Appointment Date" and get the input NEXT to the one in that column
          async () => {
            const header = page.locator('th, [role="columnheader"]').filter({ hasText: /Appointment Date/i }).first();
            if (await header.isVisible({ timeout: 3000 }).catch(() => false)) {
              // Find the filter row and get input in same column, then get the NEXT one
              const table = page.locator('table').first();
              const headers = table.locator('th, [role="columnheader"]');
              const headerCount = await headers.count();
              let appointmentDateIndex = -1;
              
              for (let i = 0; i < headerCount; i++) {
                const headerText = await headers.nth(i).textContent();
                if (headerText && /Appointment Date|Orig Appt\.? Date|Appt\.? Date/i.test(headerText.trim())) {
                  appointmentDateIndex = i;
                  break;
                }
              }
              
              if (appointmentDateIndex !== -1) {
                const filterRow = table.locator('tr').filter({ has: page.locator('input, .MuiInputBase-input') }).first();
                const inputs = filterRow.locator('input, .MuiInputBase-input');
                // Get the input NEXT to the appointment date column (index + 1)
                const nextIndex = appointmentDateIndex + 1;
                if (nextIndex < await inputs.count()) {
                  return inputs.nth(nextIndex);
                }
                // If next doesn't exist, try the one before (index - 1)
                if (appointmentDateIndex > 0) {
                  return inputs.nth(appointmentDateIndex - 1);
                }
                // Fallback to the one in the appointment date column itself
                return inputs.nth(appointmentDateIndex);
              }
            }
            throw new Error('Appointment Date column not found');
          },
          // Strategy 3: Find input near "Appointment Date" text
          () => page.locator('div, td, th').filter({ hasText: /Appointment Date/i }).locator('input, .MuiInputBase-input').first(),
          // Strategy 4: Find all inputs and check which one is in appointment date column
          async () => {
            const inputs = page.locator('input[type="text"], .MuiInputBase-input');
            const count = await inputs.count();
            for (let i = 0; i < count; i++) {
              const input = inputs.nth(i);
              const isInAppointmentDateColumn = await input.evaluate((el) => {
                const cell = el.closest('td, th');
                if (!cell) return false;
                const row = cell.closest('tr');
                if (!row) return false;
                const table = row.closest('table');
                if (!table) return false;
                const headerRow = table.querySelector('tr');
                if (!headerRow) return false;
                const cells = Array.from(headerRow.querySelectorAll('th, td, [role="columnheader"]'));
                const cellIndex = Array.from(row.querySelectorAll('td, th')).indexOf(cell);
                if (cellIndex >= 0 && cellIndex < cells.length) {
                  const headerText = cells[cellIndex].textContent || '';
                  return /Appointment Date|Orig Appt\.? Date|Appt\.? Date/i.test(headerText);
                }
                return false;
              });
              if (isInAppointmentDateColumn) {
                return input;
              }
            }
            throw new Error('Appointment Date input not found');
          }
        ];
        
        // Try each strategy
        for (let i = 0; i < strategies.length; i++) {
          try {
            const button = await strategies[i]();
            if (await button.isVisible({ timeout: 3000 }).catch(() => false)) {
              appointmentDateButton = button;
              console.log(`   ✅ Found Appointment Date filter using strategy ${i + 1}`);
              break;
            }
          } catch (e) {
            continue;
          }
        }
        
        if (!appointmentDateButton) {
          throw new Error('Could not find Appointment Date filter button');
        }
        
        // Scroll the button into view (scroll horizontally if needed)
        await appointmentDateButton.scrollIntoViewIfNeeded();
      await page.waitForTimeout(500);

        // Try scrolling the container horizontally to the right
        try {
          const container = page.locator('.mantine-vg7fqm, table, [class*="table"]').first();
          await container.evaluate((el) => {
            el.scrollLeft = el.scrollWidth; // Scroll to the right
          });
          await page.waitForTimeout(500);
        } catch (e) {
          // If container scroll doesn't work, try page scroll
          await page.evaluate(() => {
            window.scrollBy(500, 0); // Scroll right
          });
          await page.waitForTimeout(500);
        }
        
        // Click the appointment date filter button to open the date picker
        await appointmentDateButton.click();
        await page.waitForTimeout(1500); // Wait for popover to open
        
        // Wait for date picker to be visible and then fill dates directly
        // Try to find date input fields - use simpler selectors
        await page.waitForTimeout(500);
        
        // Fill start date - try multiple approaches
        try {
          const startDateInput = page.getByRole('textbox', { name: 'MM/DD/YYYY' }).first();
          await startDateInput.waitFor({ state: 'visible', timeout: 5000 });
          await startDateInput.click();
          await page.waitForTimeout(300);
          await startDateInput.fill(appointmentDateValue);
        } catch (e) {
          // Try alternative: find input by placeholder or type
          const altInput = page.locator('input[type="text"], input[placeholder*="date" i], input[placeholder*="Date" i]').first();
          await altInput.click();
          await page.waitForTimeout(300);
          await altInput.fill(appointmentDateValue);
        }
        
        // Fill end date (use appointment date as end date)
        const endDate = row.appointmentDateFormatted || appointmentDateValue;
    await page.waitForTimeout(500);
        
        try {
          const endDateInput = page.getByRole('textbox', { name: 'MM/DD/YYYY' }).nth(1);
          await endDateInput.waitFor({ state: 'visible', timeout: 3000 });
          await endDateInput.click();
          await page.waitForTimeout(300);
          await endDateInput.fill(endDate);
        } catch (e) {
          // If second date field not found, use same date for both
          console.log(`   ℹ️  End date field not found, using same date for start and end`);
        }
        
        // Apply the filter
    await page.waitForTimeout(500);
        const applyBtn = page.getByRole('button', { name: 'Apply' });
        await applyBtn.waitFor({ state: 'visible', timeout: 5000 });
        await applyBtn.click();
        await page.waitForTimeout(2000); // Wait for filter to apply
        
        console.log(`   ✅ Appointment Date filter updated: ${appointmentDateValue} to ${endDate}`);
    } catch (error) {
        console.log(`   ⚠️  Appointment Date filter failed: ${error.message}. Continuing...`);
        // Continue even if Appointment Date fails
      }
    } else {
      console.log(`   ℹ️  No Appointment Date provided in Excel row, skipping Appointment Date filter`);
    }

    // Get count of matching patients/claims (rows containing IntellyChart)
    const table = page.locator('table').first();
    await table.waitFor({ state: 'visible', timeout: 5000 });
    const allRows = table.locator('tbody tr, tr[role="row"]')
      .filter({ hasNot: page.locator('input[placeholder*="filter" i]') })
      .filter({ has: page.locator('text=IntellyChart') });
    const patientCount = await allRows.count();
    
    console.log(`   📊 Found ${patientCount} patient(s)/claim(s) with MRN ${row.mrn}`);

    if (patientCount === 0) {
      console.log(`   ⚠️  No patients found with MRN ${row.mrn}`);
      allResults.push({
        rowNumber: row._rowNumber || rowIndex + 1,
        mrn: row.mrn,
        customId: row.customId,
        dos: row.dos,
        appointmentDate: row.appointmentDate,
        cptCodes: row.cptCodes || [],
        patientIndex: 0,
        totalPatients: 0,
        status: 'Skipped',
        errorMessage: 'No patients found with this MRN',
        generatedId: '',
        processingTime: Date.now() - startTime,
        timestamp: new Date().toISOString(),
        notes: 'No matching patients in table'
      });
      return allResults;
    }

    // Process each patient/claim
    for (let i = 0; i < patientCount; i++) {
      const result = await processPatientClaim(page, row, i, patientCount);
      allResults.push(result);
      
      // Small delay between patients (except for the last one)
      if (i < patientCount - 1) {
        await page.waitForTimeout(1000);
      }
    }

    console.log(`   ✅ Completed processing ${patientCount} patient(s)/claim(s) for MRN ${row.mrn}`);
    // Note: Filters are already cleared after the last patient in processPatientClaim

  } catch (error) {
    console.error(`   ❌ Failed to process MRN ${row.mrn}: ${error.message}`);
    allResults.push({
      rowNumber: row._rowNumber || rowIndex + 1,
      mrn: row.mrn,
      customId: row.customId,
      dos: row.dos,
      appointmentDate: row.appointmentDate,
      cptCodes: row.cptCodes || [],
      patientIndex: 0,
      totalPatients: 0,
      status: 'Failed',
      errorMessage: error.message,
      generatedId: '',
      processingTime: Date.now() - startTime,
      timestamp: new Date().toISOString(),
      notes: `Error processing MRN: ${error.message}`
    });
  }

  return allResults;
}

/* =======================
   Main Execution
======================= */
(async () => {
  let browser;
  const results = [];

  try {
    // Load Excel data
    console.log("📊 Loading Excel data...");
    const excelData = loadExcelData(CONFIG.EXCEL);
    
    console.log(`\n📈 Processing ${excelData.validRows.length} valid rows${CONFIG.DRY_RUN ? ' (DRY RUN MODE)' : ''}`);

    if (excelData.validRows.length === 0) {
      console.error("❌ No valid rows to process");
      return;
    }

    // Add skipped rows to results
    excelData.invalidRows.forEach(row => {
      results.push({
        rowNumber: row._rowNumber,
        mrn: row.mrn,
        customId: row.customId,
        dos: row.dos,
        appointmentDate: row.appointmentDate,
        cptCodes: row.cptCodes || [],
        status: 'Skipped',
        errorMessage: row._errors.join('; '),
        generatedId: '',
        processingTime: 0,
        timestamp: new Date().toISOString(),
        notes: 'Row validation failed'
      });
    });

    // Start browser session (only if not dry run)
    let page = null;
    if (!CONFIG.DRY_RUN) {
      const { browser: activeBrowser, page: activePage } = await startQHSLabSession();
      browser = activeBrowser;
      page = activePage;
    }

    // Process each valid row (each MRN from Excel)
    for (let i = 0; i < excelData.validRows.length; i++) {
      const row = excelData.validRows[i];
      
      try {
      if (CONFIG.DRY_RUN) {
          const rowResults = await processRow(null, row, i);
          // processRow now returns an array of results (one per patient)
          results.push(...rowResults);
      } else {
          // Check if page is still open before processing
          if (page && !page.isClosed()) {
            const rowResults = await processRow(page, row, i);
            // processRow now returns an array of results (one per patient)
            results.push(...rowResults);
      } else {
            console.error(`   ❌ Page is closed. Cannot process MRN ${row.mrn}. Skipping remaining rows.`);
            // Add error result for this row
            results.push({
              rowNumber: row._rowNumber || i + 1,
              mrn: row.mrn,
              customId: row.customId,
              dos: row.dos,
              appointmentDate: row.appointmentDate,
              cptCodes: row.cptCodes || [],
              patientIndex: 0,
              totalPatients: 0,
              status: 'Failed',
              errorMessage: 'Page closed unexpectedly',
              generatedId: '',
              processingTime: 0,
              timestamp: new Date().toISOString(),
              notes: 'Page closed - cannot continue processing'
            });
            break; // Stop processing if page is closed
          }
        
          // Small delay between MRNs (not between individual patients)
        if (i < excelData.validRows.length - 1) {
          await page.waitForTimeout(2000);
        }
      }
      } catch (error) {
        console.error(`   ❌ Error processing Excel row ${i + 1} (MRN: ${row.mrn}): ${error.message}`);
        // Add error result but continue with next row
        results.push({
          rowNumber: row._rowNumber || i + 1,
          mrn: row.mrn,
          customId: row.customId,
          dos: row.dos,
          appointmentDate: row.appointmentDate,
          cptCodes: row.cptCodes || [],
          patientIndex: 0,
          totalPatients: 0,
          status: 'Failed',
          errorMessage: error.message,
          generatedId: '',
          processingTime: 0,
          timestamp: new Date().toISOString(),
          notes: `Error in main loop: ${error.message}`
        });
        
        // If page is closed, stop processing
        if (page && page.isClosed()) {
          console.error(`   ❌ Page closed. Stopping processing.`);
          break;
        }
        
        // Continue with next row
        console.log(`   ⏭️  Continuing with next MRN...`);
      }
    }

    // Update original Excel file with status instead of creating new report
    console.log("\n📝 Updating original Excel file with status...");
    try {
      const updatedPath = updateOriginalExcel(
        excelData.filePath, 
        results, 
        CONFIG.EXCEL.sheetName || process.env.EXCEL_SHEET_NAME
      );
      console.log(`✅ Original Excel file updated: ${path.basename(updatedPath)}`);
    } catch (updateError) {
      console.error(`⚠️  Failed to update original Excel file: ${updateError.message}`);
      // Fallback to generating report if update fails
      console.log("📊 Generating backup report instead...");
    const reportPath = generateReport(results);
      console.log(`✅ Backup report saved to: ${reportPath}`);
    }

    // Print summary
    const summary = {
      total: results.length,
      success: results.filter(r => r.status === 'Success').length,
      failed: results.filter(r => r.status === 'Failed').length,
      skipped: results.filter(r => r.status === 'Skipped' || r.status.includes('Skipped')).length
    };

    console.log("\n📈 Processing Summary:");
    console.log(`   Total: ${summary.total}`);
    console.log(`   ✅ Success: ${summary.success}`);
    console.log(`   ❌ Failed: ${summary.failed}`);
    console.log(`   ⏭️  Skipped: ${summary.skipped}`);

    console.log("\n🎯 Automation completed");

  } catch (error) {
    console.error("❌ Script error:", error.message);
    console.error(error.stack);
    
    // Update original Excel file with status if we have any results
    if (results.length > 0 && excelData && excelData.filePath) {
      try {
        updateOriginalExcel(
          excelData.filePath, 
          results, 
          CONFIG.EXCEL.sheetName || process.env.EXCEL_SHEET_NAME
        );
        console.log(`✅ Original Excel file updated with status`);
      } catch (updateError) {
        console.error("❌ Failed to update original Excel file:", updateError.message);
        // Fallback to generating report
      try {
        generateReport(results);
      } catch (reportError) {
        console.error("❌ Failed to generate report:", reportError.message);
        }
      }
    }
  } finally {
    // Close browser after all processing is complete
    if (browser) {
      console.log("\n🔒 Closing browser...");
      await browser.close();
      console.log("✅ Browser closed");
    }
  }
})();
