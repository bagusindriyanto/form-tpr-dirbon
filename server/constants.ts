// Script Properties memungkinkan project development dan production memakai
// spreadsheet berbeda tanpa mengubah kode. Nilai fallback di bawah ini hanya
// jaring pengaman; isi ketiga property di masing-masing project Apps Script
// (Project Settings > Script Properties) agar environment tidak ambigu.
const SPREADSHEET_ID_PROPERTIES = {
  database: 'DATABASE_SPREADSHEET_ID',
  lkh: 'LKH_SPREADSHEET_ID',
  tprDirbon: 'TPR_DIRBON_SPREADSHEET_ID',
  logging: 'LOGGING_SPREADSHEET_ID',
} as const;

type SpreadsheetArea = keyof typeof SPREADSHEET_ID_PROPERTIES;

function getSpreadsheetId(area: SpreadsheetArea, fallback: string): string {
  const value = PropertiesService.getScriptProperties()
    .getProperty(SPREADSHEET_ID_PROPERTIES[area])
    ?.trim();
  return value || fallback;
}

const SPREADSHEET_CONFIG = {
  database: {
    spreadsheetId: getSpreadsheetId(
      'database',
      '1AtRp3wCRilR-WYhpbL7MC2-thMHI2O-plKObyQB0p7A',
    ),
    sheetNames: {
      member: 'Operator',
      target: 'Target',
      output: 'Output',
      spoAndSize: 'SPO dan Size',
    },
  },
  lkh: {
    spreadsheetId: getSpreadsheetId(
      'lkh',
      '1Z5k7RuMQa-zlJr8dk2eUQ5NlVlpkDJCtMAk5HPoJeJE',
    ),
    sheetNames: {
      dirbon: '[Trial] OUTPUT DIRBON/JAM',
      tpr: '[Trial] Output TPR/Jam',
      spoAndPrc: 'Daftar SPOPRC',
    },
  },
  tprDirbon: {
    spreadsheetId: getSpreadsheetId(
      'tprDirbon',
      '1ky7rSxy20KQJ3zcZw0mVY4qE48G6CKAjAbb3zybEmKk',
    ),
    sheetNames: {
      kanban: 'Kanban',
    },
  },
  logging: {
    spreadsheetId: getSpreadsheetId(
      'logging',
      '1eugWu3PvkpLayisCDqC93j3cp74C33FGi75Iqw-zBOk',
    ),
    sheetNames: {
      errorLog: 'Log Error',
    },
  },
} as const satisfies Record<SpreadsheetArea, SpreadsheetConfig>;

const LKH_SHEETS = {
  dirbon: {
    startRow: 4,
    helperCol: 54,
    outputBase: 19,
    sizeCol: 18,
    createBlocks: (sheet, target, jobs, config) =>
      createJobBlocksDirbon(sheet, target, jobs, config),
  },
  tpr: {
    startRow: 3,
    helperCol: 42,
    outputBase: 18,
    sizeCol: 17,
    createBlocks: (sheet, target, jobs, config) =>
      createJobBlocksTPR(sheet, target, jobs, config),
  },
} as const satisfies Record<Division, LKHConfig>;

/**
 * Jalankan manual dari editor Apps Script untuk memastikan project yang aktif
 * membaca spreadsheet yang benar sebelum testing atau deploy.
 */
function checkSpreadsheetEnvironment(): void {
  const areas: SpreadsheetArea[] = Object.keys(
    SPREADSHEET_ID_PROPERTIES,
  ) as SpreadsheetArea[];
  const properties = PropertiesService.getScriptProperties();

  for (const area of areas) {
    const property = SPREADSHEET_ID_PROPERTIES[area];
    const source = properties.getProperty(property) ? property : 'fallback';
    console.log(
      `${area}: ${SPREADSHEET_CONFIG[area].spreadsheetId} (${source})`,
    );

    try {
      const spreadsheet = SpreadsheetApp.openById(
        SPREADSHEET_CONFIG[area].spreadsheetId,
      );
      const sheetNames: string[] = Object.values(
        SPREADSHEET_CONFIG[area].sheetNames,
      );
      const missingSheets = sheetNames.filter(
        (name) => !spreadsheet.getSheetByName(name),
      );

      console.log(`  nama file: ${spreadsheet.getName()}`);
      if (missingSheets.length)
        console.log(`  sheet tidak ditemukan: ${missingSheets.join(', ')}`);
    } catch (error) {
      console.error(`  gagal dibuka: ${errorMessage(error)}`);
    }
  }
}
