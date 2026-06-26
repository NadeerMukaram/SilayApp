import * as XLSX from 'xlsx';

export type PatientRecord = {
  rowIndex: number;
  date: string;
  patientName: string;
  breastValue: string;
  isUr: boolean;
  fields: Record<string, string>;
};

export type ExcelParseResult = {
  sheetName: string;
  patients: PatientRecord[];
};

/** @deprecated Use ExcelParseResult and PatientRecord instead */
export type BreastCheckResult = PatientRecord & { sheetName: string };

const BREAST_KEYS = ['breast'];
const NAME_KEYS = ['name', 'patient name', 'patient'];
const DATE_KEYS = ['date'];
const ID_KEYS = ['id number', 'id no', 'id no.', 'id', 'hospital id'];

function normalizeHeader(value: unknown): string {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function normalizeCell(value: unknown): string {
  return String(value ?? '').trim();
}

function findColumnIndex(headers: string[], candidates: string[]): number {
  for (const candidate of candidates) {
    const index = headers.findIndex((header) => header === candidate);
    if (index >= 0) {
      return index;
    }
  }

  for (const candidate of candidates) {
    const index = headers.findIndex((header) => header.includes(candidate));
    if (index >= 0) {
      return index;
    }
  }

  return -1;
}

function getRowCell(row: (string | number | null)[], colIndex: number): unknown {
  if (colIndex < 0) {
    return null;
  }

  return row[colIndex] ?? null;
}

function rowHasData(row: (string | number | null)[]): boolean {
  return row.some((cell) => normalizeCell(cell) !== '');
}

function isPatientRow(
  row: (string | number | null)[],
  nameCol: number,
  dateCol: number,
  idCol: number,
): boolean {
  if (nameCol >= 0 && normalizeCell(getRowCell(row, nameCol))) {
    return true;
  }

  if (idCol >= 0 && normalizeCell(getRowCell(row, idCol))) {
    return true;
  }

  if (dateCol >= 0 && normalizeCell(getRowCell(row, dateCol))) {
    return true;
  }

  return rowHasData(row);
}

function formatDateValue(value: unknown): string {
  if (value instanceof Date) {
    return value.toLocaleDateString();
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const date = new Date(parsed.y, parsed.m - 1, parsed.d);
      return date.toLocaleDateString();
    }
  }

  const text = normalizeCell(value);
  return text || '(empty)';
}

function getSheetRange(sheet: XLSX.WorkSheet): XLSX.Range {
  const fallback = sheet['!ref'] ?? 'A1';
  const range = XLSX.utils.decode_range(fallback);

  for (const key of Object.keys(sheet)) {
    if (key[0] === '!') {
      continue;
    }

    const cell = XLSX.utils.decode_cell(key);
    if (cell.r < range.s.r) {
      range.s.r = cell.r;
    }
    if (cell.c < range.s.c) {
      range.s.c = cell.c;
    }
    if (cell.r > range.e.r) {
      range.e.r = cell.r;
    }
    if (cell.c > range.e.c) {
      range.e.c = cell.c;
    }
  }

  return range;
}

function sheetToMatrix(sheet: XLSX.WorkSheet): (string | number | null)[][] {
  const range = getSheetRange(sheet);
  const matrix: (string | number | null)[][] = [];

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row: (string | number | null)[] = [];

    for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      const cell = sheet[address];
      row.push(cell == null || cell.v == null ? null : cell.v);
    }

    matrix.push(row);
  }

  return matrix;
}

function findHeaderRowIndex(matrix: (string | number | null)[][]): number {
  const scanLimit = Math.min(matrix.length, 20);

  for (let rowIndex = 0; rowIndex < scanLimit; rowIndex += 1) {
    const headers = (matrix[rowIndex] ?? []).map(normalizeHeader);
    const breastCol = findColumnIndex(headers, BREAST_KEYS);
    const nameCol = findColumnIndex(headers, NAME_KEYS);
    const dateCol = findColumnIndex(headers, DATE_KEYS);

    if (breastCol >= 0 && (nameCol >= 0 || dateCol >= 0)) {
      return rowIndex;
    }
  }

  return 0;
}

function buildPatientRecord(
  row: (string | number | null)[],
  rowIndex: number,
  headers: string[],
  headerRow: (string | number | null)[],
  breastCol: number,
  nameCol: number,
  dateCol: number,
): PatientRecord {
  const breastValue = normalizeCell(getRowCell(row, breastCol));
  const isUr = breastValue.toUpperCase() === 'UR';
  const patientName =
    nameCol >= 0
      ? normalizeCell(getRowCell(row, nameCol)) || `Row ${rowIndex + 1}`
      : `Row ${rowIndex + 1}`;
  const date = dateCol >= 0 ? formatDateValue(getRowCell(row, dateCol)) : '(no date)';

  const fields: Record<string, string> = {};
  headers.forEach((header, colIndex) => {
    if (!header) {
      return;
    }

    const label = normalizeCell(headerRow[colIndex]) || header;
    const cellValue = getRowCell(row, colIndex);
    fields[label] =
      dateCol === colIndex ? formatDateValue(cellValue) : normalizeCell(cellValue) || '(empty)';
  });

  return {
    rowIndex: rowIndex + 1,
    date,
    patientName,
    breastValue: breastValue || '(empty)',
    isUr,
    fields,
  };
}

export function parseExcelBase64(base64: string): ExcelParseResult {
  const workbook = XLSX.read(base64, { type: 'base64' });
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error('The Excel file has no worksheets.');
  }

  const sheet = workbook.Sheets[sheetName];
  const matrix = sheetToMatrix(sheet);

  if (matrix.length < 2) {
    throw new Error('The Excel file has no data rows.');
  }

  const headerRowIndex = findHeaderRowIndex(matrix);
  const headerRow = matrix[headerRowIndex] ?? [];
  const headers = headerRow.map(normalizeHeader);

  const breastCol = findColumnIndex(headers, BREAST_KEYS);
  if (breastCol < 0) {
    throw new Error('Could not find a "Breast" column in the Excel file.');
  }

  const nameCol = findColumnIndex(headers, NAME_KEYS);
  const dateCol = findColumnIndex(headers, DATE_KEYS);
  const idCol = findColumnIndex(headers, ID_KEYS);

  const patients: PatientRecord[] = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    if (!isPatientRow(row, nameCol, dateCol, idCol)) {
      continue;
    }

    patients.push(
      buildPatientRecord(row, rowIndex, headers, headerRow, breastCol, nameCol, dateCol),
    );
  }

  if (patients.length === 0) {
    throw new Error('No data rows found below the header row.');
  }

  return {
    sheetName,
    patients,
  };
}
