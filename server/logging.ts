// Log error disimpan di spreadsheet logging terpisah supaya tetap bisa dibaca
// setelah deployment production, termasuk saat web app diakses tanpa login.
const ERROR_LOG_MAX_CELL_LENGTH = 45_000;

type ErrorLogEntry = {
  functionName: string;
  message: string;
  stack?: string;
};

function truncateErrorLog_(
  value: string,
  maxLength = ERROR_LOG_MAX_CELL_LENGTH,
): string {
  if (value.length <= maxLength) return value;

  return `${value.slice(0, maxLength - 3)}...`;
}

// Cegah teks error diperlakukan sebagai formula oleh Google Sheets.
function formatErrorLogCell_(
  value: string,
  maxLength = ERROR_LOG_MAX_CELL_LENGTH,
): string {
  const truncated = truncateErrorLog_(value, maxLength);

  return /^[=+\-@]/.test(truncated) ? `'${truncated}` : truncated;
}

function writeErrorLog_(entry: ErrorLogEntry): void {
  try {
    const {
      logging: { spreadsheetId, sheetNames },
    } = SPREADSHEET_CONFIG;
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(sheetNames.errorLog);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetNames.errorLog);
      sheet.appendRow([
        'Timestamp',
        'Level',
        'Source',
        'Function',
        'Message',
        'Stack',
      ]);
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(1);
    }

    sheet.appendRow([
      new Date(),
      'ERROR',
      'server',
      formatErrorLogCell_(entry.functionName, 200),
      formatErrorLogCell_(entry.message),
      formatErrorLogCell_(entry.stack ?? ''),
    ]);
    SpreadsheetApp.flush();
  } catch (error) {
    // Jangan pernah melempar error dari logger, supaya error asli tetap sampai
    // ke client dan proses tidak berubah.
    console.error('[logger] gagal menulis log error:', error);
  }
}

function logServerError_(scope: string, error: unknown): void {
  console.error(`[server] ${scope}:`, error);

  writeErrorLog_({
    functionName: scope,
    message: errorMessage(error),
    stack: error instanceof Error ? (error.stack ?? '') : '',
  });
}

function withErrorLogging_<T>(scope: string, callback: () => T): T {
  try {
    return callback();
  } catch (error) {
    logServerError_(scope, error);
    throw error;
  }
}

/**
 * Jalankan manual dari editor Apps Script untuk memastikan sheet "Log Error"
 * bisa dibuat dan ditulis.
 */
function checkErrorLog(): void {
  writeErrorLog_({
    functionName: 'checkErrorLog',
    message: 'Tes penulisan log error dari editor Apps Script.',
  });

  console.log('Log error berhasil ditulis ke sheet "Log Error".');
}
