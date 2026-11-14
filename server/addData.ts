function addTarget(targetFormData: TargetFormData) {
  withScriptLock(() => {
    const {
      database: { spreadsheetId, sheetNames },
    } = SPREADSHEET_CONFIG;
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetNames.target);
    if (!sheet) throw new Error(`Sheet '${sheetNames.target}' tidak ditemukan`);

    const createdAt = new Date();
    const {
      line,
      shift,
      spo,
      style,
      prcs,
      parts,
      workingHour,
      leader,
      operators,
      target,
      sizes,
      moldings,
    } = targetFormData;
    const manPower = operators.length;

    const normalizedSizes = sizes.map((value) => value.trim().toUpperCase());

    if (!normalizedSizes.length) {
      throw new Error('Minimal satu size harus diisi');
    }

    if (normalizedSizes.length !== moldings.length) {
      throw new Error('Jumlah size dan molding tidak sama');
    }

    if (normalizedSizes.some((value) => !value)) {
      throw new Error('Size tidak boleh kosong');
    }

    if (new Set(normalizedSizes).size !== normalizedSizes.length) {
      throw new Error('Size tidak boleh duplikat');
    }

    if (moldings.some((value) => !Number.isInteger(value) || value < 1)) {
      throw new Error('Isi molding harus bilangan bulat minimal 1');
    }

    const lastRow = sheet.getLastRow();
    const nextRow = lastRow + 1;

    const recordDate = getShiftDate(shift);
    const targetId = [toDateSerial(recordDate), line, shift, spo].join('-');

    const payload = normalizedSizes.map<CellValue[]>((sizeValue, index) => {
      const rowNumber = nextRow + index;
      return [
        createdAt,
        `=IF(
          D${rowNumber}=2,
          INT(A${rowNumber})+IF(
            MOD(A${rowNumber},1)>=IF(
              WEEKDAY(INT(A${rowNumber}),2)=7,
              TIME(19,30,0),
              TIME(19,0,0)
            ),
            1,
            0
          ),
          INT(A${rowNumber})
        )`,
        line,
        shift,
        spo,
        style,
        prcs.join(', '),
        parts.join(', '),
        sizeValue,
        moldings[index],
        manPower,
        workingHour,
        leader,
        operators.join(', '),
        target,
        `=ROW()`,
        targetId,
      ];
    });

    sheet.getRange(nextRow, 1, payload.length, 17).setValues(payload);
    sheet.getRange(nextRow, 2, payload.length, 1).setNumberFormat('yyyy-MM-dd');
  });
}

function addOutput(outputFormData: OutputFormData) {
  withScriptLock(() => {
    const createdAt = new Date();
    const {
      line,
      shift,
      type,
      spo,
      style,
      prcs,
      parts,
      size,
      hour,
      actual,
      repair,
      reject,
    } = outputFormData;
    const recordDate = getShiftDate(shift);

    // 1. Menambahkan formData ke LKH
    if (line.startsWith('D')) {
      updateLKH('dirbon', outputFormData, recordDate);
    } else if (line.startsWith('T')) {
      updateLKH('tpr', outputFormData, recordDate);
    }

    // 2. Menambahkan formData ke Trial Form TPR Dirbon
    const {
      database: { spreadsheetId, sheetNames },
    } = SPREADSHEET_CONFIG;
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetNames.output);
    if (!sheet) throw new Error(`Sheet '${sheetNames.output}' tidak ditemukan`);
    const lastRow = sheet.getLastRow();
    const nextRow = lastRow + 1;

    sheet.appendRow([
      createdAt,
      `=IF(
        D${nextRow}=2,
        INT(A${nextRow})+IF(
          MOD(A${nextRow},1)>=IF(
            WEEKDAY(INT(A${nextRow}),2)=7,
            TIME(19,30,0),
            TIME(19,0,0)
          ),
          1,
          0
        ),
        INT(A${nextRow})
      )`,
      line,
      shift,
      type,
      spo,
      style,
      `=XLOOKUP(G${nextRow},'Database SPO'!$B$2:$B,'Database SPO'!$C$2:$C, XLOOKUP(F${nextRow},'Database SPO'!$A$2:$A,'Database SPO'!$C$2:$C, "NOT FOUND", 0), 0)`,
      prcs.join(', '),
      parts.join(', '),
      size,
      hour,
      actual,
      repair ?? 0,
      reject ?? 0,
    ]);

    sheet.getRange(`B${nextRow}`).setNumberFormat('yyyy-MM-dd');
  });
}

function addProblem(problemFormData: ProblemFormData) {
  withScriptLock(() => {
    const {
      tprDirbon: { spreadsheetId, sheetNames },
    } = SPREADSHEET_CONFIG;
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetNames.kanban);
    if (!sheet) throw new Error(`Sheet '${sheetNames.kanban}' tidak ditemukan`);

    const createdAt = new Date();
    const { line, problem, pic } = problemFormData;
    const shift = getShiftByTimestamp(createdAt);
    const recordDate = getShiftDate(shift);
    const week = getISOWeek(createdAt);
    const section = line.startsWith('T')
      ? 'TPR'
      : line.startsWith('D')
        ? 'Dirbon'
        : 'Unknown';

    sheet.appendRow([
      createdAt,
      line,
      problem,
      pic,
      null,
      null,
      recordDate,
      `=ROW()-1`,
      week,
      section,
      null,
      shift,
    ]);
  });
}
