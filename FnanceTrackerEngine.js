/**
 * ===========================
 * PERSONAL FINANCE TRACKER AUTOMATION SCRIPT
 * Version 3.0 - Complete Implementation with Full Documentation
 * ===========================
 *
 * PURPOSE:
 * This Google Apps Script automates the processing of recurring and single financial transactions
 * across monthly sheets, maintaining running balances for multiple accounts and providing
 * comprehensive financial tracking and projection capabilities.
 *
 * SYSTEM ARCHITECTURE:
 * - Processes recurring transactions (bi-weekly, monthly, etc.) from a configuration sheet
 * - Handles single transaction inputs from a dedicated input sheet
 * - Maintains daily running balances across 12 different accounts
 * - Automatically populates monthly sheets (Jan-Dec) with chronologically sorted transactions
 * - Provides starting balances from account registry and carries forward month-to-month
 *
 * SHEET DEPENDENCIES:
 * - Accounts: Master account registry with current balances
 * - Categories: Transaction category definitions
 * - Recurring Transactions: Automated recurring transaction setup
 * - Single Transactions: Manual transaction input hub
 * - Monthly Sheets (Jan-Dec): Generated transaction logs with running balances
 *
 * AUTHOR: Personal Finance Tracker System
 * CREATED: 2025
 * LAST MODIFIED: Phase 3 Implementation
 */

// ===========================
// CONFIGURATION SECTION
// ===========================

/**
 * Central configuration object containing all system settings and mappings.
 * This is the single source of truth for sheet names, account structures, and operational parameters.
 *
 * IMPORTANT: Account names must match EXACTLY across all sheets for proper functionality.
 * Any changes to account names must be updated in all three locations:
 * 1. Accounts sheet
 * 2. Monthly sheet column headers
 * 3. This accountColumns array
 */
const CONFIG = {
    // Sheet name mappings - must match exact sheet tab names in workbook
    sheets: {
        accounts: 'Accounts',                    // Master account registry
        recurring: 'Recurring Transactions',    // Recurring transaction definitions
        single: 'Single Transactions',          // Manual transaction input
        categories: 'Categories'                 // Transaction category definitions
    },

    /**
     * Account column mapping array - CRITICAL CONFIGURATION
     * This array defines:
     * 1. The order of account balance columns in monthly sheets (G through R)
     * 2. The exact account names that must match the Accounts sheet
     * 3. The accounts tracked in running balance calculations
     *
     * Column mapping:
     * G: Capital One Checking    H: Acorns Checking       I: Chase Checking
     * J: Savings Account         K: IRA                   L: Investment Account
     * M: CASH                    N: Destiny Card          O: Aspire Card
     * P: Indigo Card             Q: Capital One Quicksilver  R: Milestone Card
     */
    accountColumns: [
        'Capital One Checking',      // Primary checking - 70% paycheck
        'Acorns Checking',          // Secondary checking - 20% paycheck
        'Chase Checking',           // Tertiary checking - 10% paycheck
        'Savings Account',          // Emergency fund savings
        'IRA',                      // Retirement account (cash tracking only)
        'Investment Account',       // Investment account (cash tracking only)
        'CASH',                     // Physical cash tracking
        'Destiny Card',            // Credit card (negative balance)
        'Aspire Card',             // Credit card (negative balance)
        'Indigo Card',             // Credit card (negative balance)
        'Capital One Quicksilver', // Credit card (negative balance)
        'Milestone Card'           // Credit card (negative balance)
    ],

    // Monthly sheet names array - must match exact sheet tab names
    monthSheets: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

    // Current operating year
    year: 2025,

    // First month with transaction data (September = 9)
    // This determines where balance calculations begin
    startMonth: 9
};

// ===========================
// MAIN ORCHESTRATION FUNCTION
// ===========================

/**
 * Primary automation function that orchestrates the complete update process.
 * This is the main entry point called from the menu system.
 *
 * PROCESS FLOW:
 * 1. Clear all existing monthly sheet data
 * 2. Retrieve initial account balances from Accounts sheet
 * 3. Process each month sequentially from startMonth through December
 * 4. Carry forward ending balances to subsequent months
 * 5. Provide user feedback on completion or errors
 *
 * ERROR HANDLING:
 * - Comprehensive try-catch with user-friendly error messages
 * - Console logging for debugging purposes
 * - Graceful failure with specific error identification
 */
function updateAllMonthlySheets() {
    console.log('Starting update of all monthly sheets...');

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // Step 1: Clear all existing monthly data to ensure clean state
        clearAllMonthlySheets(ss);

        // Step 2: Get starting account balances from Accounts sheet
        const initialBalances = getAccountBalances(ss);

        // Step 3: Process each month sequentially
        // This ensures proper balance carry-forward between months
        let previousEndBalances = null;

        for (let monthIndex = CONFIG.startMonth - 1; monthIndex < 12; monthIndex++) {
            const monthName = CONFIG.monthSheets[monthIndex];
            const monthNumber = monthIndex + 1;

            console.log(`Processing ${monthName} ${CONFIG.year}...`);

            // Determine starting balances for this month:
            // - September (startMonth): Use initial balances from Accounts sheet
            // - Other months: Use ending balances from previous month
            const startingBalances = (monthNumber === CONFIG.startMonth)
                ? initialBalances.map(acc => acc.balance)
                : previousEndBalances;

            // Skip months without data (before startMonth or missing previous balances)
            if (!startingBalances && monthNumber > CONFIG.startMonth) {
                console.log(`Skipping ${monthName} - no previous balance data`);
                continue;
            }

            // Process the month and capture ending balances for next month
            const endBalances = processMonth(
                ss,
                monthName,
                monthNumber,
                CONFIG.year,
                startingBalances
            );

            // Store ending balances for next month's starting balances
            if (endBalances) {
                previousEndBalances = endBalances;
            }
        }

        // Success notification to user
        SpreadsheetApp.getUi().alert('✅ All monthly sheets updated successfully!');

    } catch (error) {
        // Error handling with detailed logging and user notification
        console.error('Error in updateAllMonthlySheets:', error);
        SpreadsheetApp.getUi().alert('❌ Error: ' + error.toString());
    }
}

// ===========================
// MONTH PROCESSING ENGINE
// ===========================

/**
 * Processes a single month's transactions and balance calculations.
 * This is the core engine that handles transaction processing for one month.
 *
 * @param {Spreadsheet} ss - The active spreadsheet object
 * @param {string} monthName - Name of the month sheet (e.g., 'Sep')
 * @param {number} monthNumber - Numeric month (1-12)
 * @param {number} year - Year being processed
 * @param {Array<number>} startingBalances - Array of starting account balances
 * @returns {Array<number>|null} - Ending account balances or null if no data
 *
 * PROCESS FLOW:
 * 1. Retrieve all recurring transactions for this month
 * 2. Retrieve all single transactions for this month
 * 3. Combine and sort transactions chronologically
 * 4. Build data array with starting balance and all transactions
 * 5. Calculate running balances after each transaction
 * 6. Write data to monthly sheet with proper formatting
 * 7. Return ending balances for next month
 */
function processMonth(ss, monthName, monthNumber, year, startingBalances) {
    const monthSheet = ss.getSheetByName(monthName);
    if (!monthSheet) {
        throw new Error(`Sheet ${monthName} not found`);
    }

    // Step 1: Get all transactions for this month
    const recurringTrans = getRecurringTransactionsForMonth(ss, monthNumber, year);
    const singleTrans = getSingleTransactionsForMonth(ss, monthNumber, year);

    // Step 2: Skip processing if no data available
    if (!startingBalances && recurringTrans.length === 0 && singleTrans.length === 0) {
        console.log(`Skipping ${monthName} - no data`);
        return null;
    }

    // Step 3: Combine and sort all transactions
    const allTransactions = [...recurringTrans, ...singleTrans].sort((a, b) => {
        // Primary sort: by date
        const dateCompare = a.date.getTime() - b.date.getTime();
        if (dateCompare !== 0) return dateCompare;

        // Secondary sort: Income before expenses on same day (cash flow optimization)
        if (a.category === 'Income' && b.category !== 'Income') return -1;
        if (b.category === 'Income' && a.category !== 'Income') return 1;

        return 0;
    });

    // Step 4: Build the data array for sheet output
    const data = [];
    let currentBalances = startingBalances ? [...startingBalances] : null;

    if (currentBalances) {
        // Add starting balance row (first day of month)
        data.push([
            new Date(year, monthNumber - 1, 1),  // Date
            'Starting Balance',                   // Description
            '',                                   // Category
            '',                                   // Account
            '',                                   // Amount
            'Initial',                           // Source
            ...currentBalances,                  // All account balances
            currentBalances.reduce((sum, bal) => sum + bal, 0)  // Net Worth
        ]);

        // Step 5: Process each transaction and update running balances
        allTransactions.forEach(trans => {
            // Update the balance for the transaction account
            const accountIndex = CONFIG.accountColumns.indexOf(trans.account);
            if (accountIndex !== -1) {
                currentBalances[accountIndex] += trans.amount;
            }

            // Handle transfer transactions (TODO: enhance with dedicated transfer column)
            if (trans.transferTo) {
                const transferIndex = CONFIG.accountColumns.indexOf(trans.transferTo);
                if (transferIndex !== -1) {
                    currentBalances[transferIndex] += Math.abs(trans.amount);
                }
            }

            // Add transaction row with updated balances
            data.push([
                trans.date,                                              // Date
                trans.description,                                       // Description
                trans.category,                                          // Category
                trans.account,                                           // Account
                trans.amount,                                            // Amount
                trans.source,                                            // Source
                ...currentBalances,                                      // All account balances
                currentBalances.reduce((sum, bal) => sum + bal, 0)      // Net Worth
            ]);
        });
    }

    // Step 6: Write data to sheet and format
    if (data.length > 0) {
        monthSheet.getRange(2, 1, data.length, data[0].length).setValues(data);

        // Apply formatting for better readability
        formatMonthSheet(monthSheet, data.length);

        // Return ending balances for next month's processing
        return currentBalances;
    }

    return null;
}

// ===========================
// RECURRING TRANSACTION PROCESSOR
// ===========================

/**
 * Extracts and calculates recurring transactions for a specific month.
 * Handles complex frequency calculations and date generation.
 *
 * @param {Spreadsheet} ss - The active spreadsheet object
 * @param {number} monthNumber - Target month (1-12)
 * @param {number} year - Target year
 * @returns {Array<Object>} - Array of transaction objects for the month
 *
 * SUPPORTED FREQUENCIES:
 * - Bi-weekly: Every 14 days from start date
 * - Monthly: Specific day of month
 * - Weekly: Every 7 days from start date
 * - Yearly: Annual occurrence on same date
 *
 * TRANSACTION OBJECT STRUCTURE:
 * {
 *   date: Date,           // Transaction date
 *   description: string,  // Transaction description
 *   category: string,     // Transaction category
 *   account: string,      // Account name
 *   amount: number,       // Transaction amount (signed)
 *   source: string,       // Always 'Recurring'
 *   transferTo: string    // Transfer destination (currently null)
 * }
 */
function getRecurringTransactionsForMonth(ss, monthNumber, year) {
    const sheet = ss.getSheetByName(CONFIG.sheets.recurring);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    // Get all data including headers (14 columns total)
    const allData = sheet.getRange(1, 1, lastRow, 14).getValues();

    const transactions = [];
    const monthStart = new Date(year, monthNumber - 1, 1);
    const monthEnd = new Date(year, monthNumber, 0);

    // Process each row, skipping headers and section dividers
    for (let i = 0; i < allData.length; i++) {
        const row = allData[i];

        // Skip header rows, section dividers, and inactive transactions
        if (!row[0] ||
            row[0].toString().includes('===') ||
            row[0] === 'Description' ||
            row[9] !== true) { // Active column (J) must be TRUE
            continue;
        }

        // Parse recurring transaction item
        const item = {
            description: row[0],                                    // Column A
            category: row[1],                                       // Column B
            amount: Number(row[2]),                                 // Column C
            account: row[3],                                        // Column D
            frequency: row[4],                                      // Column E
            startDate: row[5] ? new Date(row[5]) : null,           // Column F
            endDate: row[6] ? new Date(row[6]) : null,             // Column G
            dayOfMonth: row[7],                                     // Column H
            dayOfWeek: row[8]                                       // Column I
        };

        // Validate required fields
        if (!item.startDate) continue;
        if (item.startDate > monthEnd) continue;
        if (item.endDate && item.endDate < monthStart) continue;

        // Calculate occurrence dates for this item in this month
        const dates = calculateRecurringDates(item, monthStart, monthEnd);

        // Create transaction objects for each calculated date
        dates.forEach(date => {
            let amount = item.amount;

            // Apply correct sign based on category:
            // - Income: Always positive
            // - Transfer: Keep original sign (TODO: enhance transfer handling)
            // - Everything else: Always negative (expenses)
            if (item.category === 'Income') {
                amount = Math.abs(amount);
            } else if (item.category !== 'Transfer') {
                amount = -Math.abs(amount);
            }

            transactions.push({
                date: date,
                description: item.description,
                category: item.category,
                account: item.account,
                amount: amount,
                source: 'Recurring',
                transferTo: null  // TODO: Implement transfer account support
            });
        });
    }

    return transactions;
}

// ===========================
// RECURRING DATE CALCULATOR
// ===========================

/**
 * Calculates specific occurrence dates for recurring transactions within a month.
 * Handles complex frequency patterns and date arithmetic.
 *
 * @param {Object} item - Recurring transaction configuration object
 * @param {Date} monthStart - First day of target month
 * @param {Date} monthEnd - Last day of target month
 * @returns {Array<Date>} - Array of calculated occurrence dates
 *
 * FREQUENCY HANDLING:
 * - monthly: Uses dayOfMonth field, handles month-end edge cases
 * - bi-weekly/biweekly: 14-day intervals from start date
 * - weekly: 7-day intervals from start date
 * - yearly: Same date each year
 *
 * DATE VALIDATION:
 * - Respects start and end date boundaries
 * - Handles month boundaries correctly
 * - Removes duplicates and sorts chronologically
 */
function calculateRecurringDates(item, monthStart, monthEnd) {
    if (!item || !item.frequency) {
        console.log('Skipping undefined or invalid recurring item:', item);
        return [];
    }

    const freq = item.frequency.trim().toLowerCase();
    const dates = [];

    // Utility function to safely clone dates
    function cloneDate(d) {
        return new Date(d.getTime());
    }

    // Process different frequency types
    switch(freq) {
        case 'monthly':
            // Monthly transactions occur on a specific day of the month
            if (item.dayOfMonth && item.dayOfMonth > 0 && item.dayOfMonth <= 31) {
                const date = new Date(monthStart.getFullYear(), monthStart.getMonth(), item.dayOfMonth);

                // Validate date is within month and date boundaries
                if (date.getMonth() === monthStart.getMonth() &&
                    date >= item.startDate &&
                    date <= monthEnd &&
                    (!item.endDate || date <= item.endDate)) {
                    dates.push(date);
                }
            }
            break;

        case 'bi-weekly':
        case 'biweekly':
            // Bi-weekly transactions occur every 14 days from start date
        {
            let current = cloneDate(item.startDate);

            // Advance to first occurrence in target month
            while (current < monthStart) {
                current.setDate(current.getDate() + 14);
            }

            // Add all occurrences within the month
            while (current <= monthEnd) {
                if (current >= monthStart &&
                    current <= monthEnd &&
                    current >= item.startDate &&
                    (!item.endDate || current <= item.endDate)) {
                    dates.push(cloneDate(current));
                }
                current.setDate(current.getDate() + 14);
            }
        }
            break;

        case 'weekly':
            // Weekly transactions occur every 7 days from start date
        {
            let current = cloneDate(item.startDate);

            // Advance to first occurrence in target month
            while (current < monthStart) {
                current.setDate(current.getDate() + 7);
            }

            // Add all occurrences within the month
            while (current <= monthEnd) {
                if (current >= monthStart &&
                    current <= monthEnd &&
                    current >= item.startDate &&
                    (!item.endDate || current <= item.endDate)) {
                    dates.push(cloneDate(current));
                }
                current.setDate(current.getDate() + 7);
            }
        }
            break;

        case 'yearly':
            // Yearly transactions occur on the same date each year
        {
            const date = new Date(monthStart.getFullYear(), item.startDate.getMonth(), item.startDate.getDate());

            if (date >= monthStart &&
                date <= monthEnd &&
                date >= item.startDate &&
                (!item.endDate || date <= item.endDate)) {
                dates.push(date);
            }
        }
            break;

        default:
            console.log('Unknown frequency:', item.frequency);
    }

    // Remove duplicates and sort chronologically
    const uniqueDates = [...new Set(dates.map(d => d.getTime()))]
        .map(time => new Date(time))
        .sort((a, b) => a - b);

    return uniqueDates;
}

// ===========================
// SINGLE TRANSACTION PROCESSOR
// ===========================

/**
 * Extracts single (manual) transactions for a specific month.
 * Processes the Single Transactions sheet for one-time transaction entries.
 *
 * @param {Spreadsheet} ss - The active spreadsheet object
 * @param {number} monthNumber - Target month (1-12)
 * @param {number} year - Target year
 * @returns {Array<Object>} - Array of transaction objects for the month
 *
 * SINGLE TRANSACTIONS SHEET STRUCTURE:
 * Column A: Date
 * Column B: Description
 * Column C: Category
 * Column D: Account
 * Column E: Amount
 * Column F: Transfer To Account (optional)
 * Column G: Notes (optional)
 *
 * PROCESSING LOGIC:
 * - Filters transactions by date to match target month/year
 * - Applies correct amount signing based on category
 * - Handles transfer destinations (limited functionality)
 * - Skips empty or invalid rows
 */
function getSingleTransactionsForMonth(ss, monthNumber, year) {
    const sheet = ss.getSheetByName(CONFIG.sheets.single);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    // Get data starting from row 2 (skip headers), 7 columns
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const transactions = [];

    data.forEach(row => {
        if (!row[0]) return; // Skip empty rows

        const date = new Date(row[0]);

        // Filter to only include transactions from target month/year
        if (date.getMonth() + 1 === monthNumber && date.getFullYear() === year) {
            let amount = Number(row[4]) || 0;

            // Apply correct sign based on category
            if (row[2] === 'Income') {
                amount = Math.abs(amount);           // Income is always positive
            } else if (row[2] !== 'Transfer') {
                amount = -Math.abs(amount);          // Expenses are always negative
            }

            transactions.push({
                date: date,
                description: row[1] || '',           // Description
                category: row[2] || '',              // Category
                account: row[3] || '',               // Account
                amount: amount,                      // Signed amount
                source: 'Single',                    // Source identifier
                transferTo: row[5] || null           // Transfer destination
            });
        }
    });

    return transactions;
}

// ===========================
// ACCOUNT BALANCE MANAGER
// ===========================

/**
 * Retrieves current account balances from the Accounts sheet.
 * Maps account balances to the configured account order for proper column alignment.
 *
 * @param {Spreadsheet} ss - The active spreadsheet object
 * @returns {Array<Object>} - Array of account objects with names and balances
 *
 * RETURN FORMAT:
 * [
 *   { name: 'Capital One Checking', balance: 2500.00 },
 *   { name: 'Acorns Checking', balance: 800.00 },
 *   ...
 * ]
 *
 * MAPPING LOGIC:
 * - Reads account names and balances from Accounts sheet
 * - Maps to CONFIG.accountColumns order for consistent column placement
 * - Handles missing accounts by defaulting balance to 0
 * - Ensures exact name matching for proper account identification
 */
function getAccountBalances(ss) {
    const sheet = ss.getSheetByName(CONFIG.sheets.accounts);
    if (!sheet) throw new Error('Accounts sheet not found');

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error('No accounts found');

    // Get account data: columns A (name), B (type), C (balance)
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    // Map account data to configured account order
    return CONFIG.accountColumns.map(accountName => {
        const accountRow = data.find(row => row[0] === accountName);
        return {
            name: accountName,
            balance: accountRow ? Number(accountRow[2]) || 0 : 0
        };
    });
}

// ===========================
// UTILITY FUNCTIONS
// ===========================

/**
 * Clears all existing data from monthly sheets to ensure clean state.
 * Preserves headers (row 1) while clearing all transaction data.
 *
 * @param {Spreadsheet} ss - The active spreadsheet object
 *
 * CLEARING LOGIC:
 * - Processes each month sheet defined in CONFIG.monthSheets
 * - Preserves row 1 (headers) while clearing rows 2 and below
 * - Handles sheets with no data gracefully
 * - Ensures clean slate for fresh data population
 */
function clearAllMonthlySheets(ss) {
    CONFIG.monthSheets.forEach(monthName => {
        const sheet = ss.getSheetByName(monthName);
        if (sheet && sheet.getLastRow() > 1) {
            const lastRow = sheet.getLastRow();
            const lastCol = sheet.getLastColumn();
            sheet.getRange(2, 1, lastRow - 1, lastCol).clear();
        }
    });
}

/**
 * Applies comprehensive formatting to monthly sheets for better readability.
 * Includes number formatting, conditional formatting, and visual enhancements.
 *
 * @param {Sheet} sheet - The monthly sheet to format
 * @param {number} dataRows - Number of data rows to format
 *
 * FORMATTING APPLIED:
 * - Date column: M/d/yyyy format
 * - Amount column: Currency with red negatives
 * - Balance columns: Currency with red negatives
 * - Conditional formatting: Red text for negative balances
 * - Row banding: Alternating colors for readability
 */
function formatMonthSheet(sheet, dataRows) {
    if (dataRows === 0) return;

    // Format date column (Column A)
    sheet.getRange(2, 1, dataRows, 1).setNumberFormat('M/d/yyyy');

    // Format amount column (Column E) with red negatives
    const amountRange = sheet.getRange(2, 5, dataRows, 1);
    amountRange.setNumberFormat('$#,##0.00;[RED]-$#,##0.00');

    // Format all balance columns (Columns G through S)
    const balanceRange = sheet.getRange(2, 7, dataRows, 13);
    balanceRange.setNumberFormat('$#,##0.00;[RED]-$#,##0.00');

    // Add conditional formatting for negative balances
    const rules = [];
    const negativeRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setFontColor('#FF0000')
        .setRanges([balanceRange])
        .build();
    rules.push(negativeRule);
    sheet.setConditionalFormatRules(rules);

    // Add alternating row colors for improved readability
    const fullRange = sheet.getRange(2, 1, dataRows, sheet.getLastColumn());
    fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
}

// ===========================
// USER INTERFACE FUNCTIONS
// ===========================

/**
 * Creates custom menu system when spreadsheet opens.
 * Provides user-friendly access to all automation functions.
 *
 * MENU STRUCTURE:
 * - Update functions (main automation)
 * - Testing and debugging tools
 * - Configuration management
 * - Automation scheduling
 * - Help and documentation
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('💰 Finance Tracker')
        .addItem('📊 Update All Monthly Sheets', 'updateAllMonthlySheets')
        .addItem('📅 Update Current Month', 'updateCurrentMonthOnly')
        .addSeparator()
        .addItem('🧪 Test Recurring Calculations', 'testRecurringCalculations')
        .addItem('📋 View Configuration', 'showConfiguration')
        .addSeparator()
        .addItem('⚙️ Setup Daily Auto-Update', 'setupDailyTrigger')
        .addItem('🔧 Remove Auto-Update', 'removeTriggers')
        .addSeparator()
        .addItem('ℹ️ Help', 'showHelp')
        .show();
}

/**
 * Updates only the current month (simplified version).
 * Currently calls full update to maintain balance continuity.
 *
 * TODO: Implement optimized single-month update that preserves
 * cross-month balance relationships.
 */
function updateCurrentMonthOnly() {
    const today = new Date();
    const currentMonth = today.getMonth() + 1;

    if (currentMonth < CONFIG.startMonth) {
        SpreadsheetApp.getUi().alert('No data available before September 2025');
        return;
    }

    // For now, update all months to maintain balance continuity
    // TODO: Implement optimized current-month-only update
    updateAllMonthlySheets();
}

/**
 * Testing function to debug recurring transaction calculations.
 * Displays September recurring transactions in a user-friendly format.
 *
 * DEBUGGING OUTPUT:
 * - Groups transactions by date
 * - Shows description and amount for each transaction
 * - Helps verify recurring transaction logic is working correctly
 */
function testRecurringCalculations() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const recurring = getRecurringTransactionsForMonth(ss, 9, 2025); // Test September

    let message = 'September Recurring Transactions:\n\n';

    // Group transactions by date for better readability
    const byDate = {};
    recurring.forEach(trans => {
        const dateStr = Utilities.formatDate(trans.date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
        if (!byDate[dateStr]) byDate[dateStr] = [];
        byDate[dateStr].push(`${trans.description}: $${trans.amount.toFixed(2)}`);
    });

    // Build formatted output
    Object.keys(byDate).sort().forEach(date => {
        message += `${date}:\n`;
        byDate[date].forEach(trans => {
            message += `  • ${trans}\n`;
        });
        message += '\n';
    });

    SpreadsheetApp.getUi().alert('Recurring Transaction Test', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Displays current system configuration to user.
 * Helps verify setup and troubleshoot configuration issues.
 */
function showConfiguration() {
    const message = `Current Configuration:
  
Year: ${CONFIG.year}
Start Month: ${CONFIG.monthSheets[CONFIG.startMonth - 1]} (${CONFIG.startMonth})

Accounts Tracked:
${CONFIG.accountColumns.map((acc, i) => `${i + 1}. ${acc}`).join('\n')}

Monthly Sheets:
${CONFIG.monthSheets.join(', ')}`;

    SpreadsheetApp.getUi().alert('Configuration', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Sets up daily automated execution of the update function.
 * Schedules script to run automatically at 2 AM daily.
 */
function setupDailyTrigger() {
    removeTriggers();

    ScriptApp.newTrigger('updateAllMonthlySheets')
        .timeBased()
        .everyDays(1)
        .atHour(2)
        .create();

    SpreadsheetApp.getUi().alert('✅ Daily auto-update enabled (runs at 2 AM)');
}

/**
 * Removes all existing script triggers.
 * Used for disabling automation or before setting up new triggers.
 */
function removeTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Displays comprehensive help information to user.
 * Provides usage instructions and troubleshooting guidance.
 */
function showHelp() {
    const message = `Finance Tracker Help:

1. UPDATE SHEETS: Click "Update All Monthly Sheets" to refresh all data

2. RECURRING TRANSACTIONS: 
   - Managed in "Recurring Transactions" sheet
   - Changes there automatically flow to all months
   
3. SINGLE TRANSACTIONS:
   - Enter in "Single Transactions" sheet
   - Automatically appear in correct month
   
4. BALANCES:
   - September starts from Accounts sheet
   - Each month carries forward automatically
   
5. TROUBLESHOOTING:
   - Use "Test Recurring Calculations" to verify
   - Check dates match format: MM/DD/YYYY
   - Ensure Active = TRUE for recurring items

Need more help? Check your sheet structure matches the setup guide.`;

    SpreadsheetApp.getUi().alert('Help', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Centralized error handling function for consistent error management.
 *
 * @param {Error} error - The error object to handle
 * @param {string} context - Description of where the error occurred
 *
 * Provides both console logging for debugging and user-friendly alerts.
 */
function handleError(error, context) {
    console.error(`Error in ${context}:`, error);
    SpreadsheetApp.getUi().alert(
        `Error in ${context}`,
        error.toString(),
        SpreadsheetApp.getUi().ButtonSet.OK
    );
}

/**
 * ===========================
 * KNOWN LIMITATIONS AND FUTURE ENHANCEMENTS
 * ===========================
 *
 * CURRENT LIMITATIONS:
 *
 * 1. TRANSFER ACCOUNT HANDLING
 *    - Recurring Transactions sheet lacks "Transfer To Account" column
 *    - Credit card payments appear as expenses instead of balance transfers
 *    - Workaround: Manual reconciliation required
 *
 * 2. VARIABLE PAYMENT AMOUNTS
 *    - No streamlined way to handle amounts different from recurring setup
 *    - Current process: adjust start date, create single transaction, re-run script
 *    - Enhancement needed: Variable amount override system
 *
 * 3. YEAR-END ROLLOVER
 *    - Script hardcoded to start in September 2025
 *    - May not handle year transitions or fresh starts properly
 *    - Template usability limited for different start dates
 *
 * 4. DATA TRANSPORT TO NEW YEAR
 *    - No automated process for moving data to next year's workbook
 *    - Risk of manual transfer errors
 *    - Need: Year-end summary sheet and import process
 *
 * FUTURE ENHANCEMENT PRIORITIES:
 * - High: Implement transfer account functionality
 * - Medium: Variable payment amount system
 * - Medium: Dynamic start date and year handling
 * - Low: End-of-year data transport automation
 */