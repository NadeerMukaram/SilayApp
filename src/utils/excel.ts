import * as XLSX from 'xlsx';

export type PatientRecord = {
  rowIndex: number;
  date: string;           // referral date (optional)
  screeningDate: string;  // screening date (required)
  patientName: string;
  breastValue: string;
  rightBreastValue: string;
  leftBreastValue: string;
  isUr: boolean;
  fields: Record<string, string>;
};

export type ExcelParseResult = {
  sheetName: string;
  patients: PatientRecord[];
};

/** @deprecated Use ExcelParseResult and PatientRecord instead */
export type BreastCheckResult = PatientRecord & { sheetName: string };

// ─── Column key groups ────────────────────────────────────────────────────────
const RIGHT_BREAST_KEYS = ['right breast'];
const LEFT_BREAST_KEYS  = ['left breast'];
const NAME_KEYS         = ['name', 'patient name', 'patient'];
const DATE_KEYS = ['referral date'];
const SCREENING_KEYS    = ['screening date', 'screening'];
const ID_KEYS           = ['id number', 'id no', 'id no.', 'id', 'hospital id'];

// Demographics
const DEMOGRAPHIC_KEYS  = ['age', 'address', 'contact no.', 'civil status', 'no. of children'];

// OB-GYNE HISTORY sub-groups
const MENSTRUAL_KEYS    = ['menarche', 'lmp', 'aog', 'menstrual bleeding pattern', 'no. of pads/day', 'no. of pads'];
const PREGNANCY_KEYS    = ['pregnancy history', 'gtpal', 'age at first full term pregnancy'];
const CONTRACEPTIVE_KEYS= ['oral contraceptives use', 'oral contraceptive use', 'duration of usage of oral contraceptives'];

// Sexual history
const SEXUAL_KEYS       = ['age of first intercourse', 'no. of sexual partner', 'no. of sexual partners', 'spouse/partner/s'];

// Family & social history
const FAMILY_SOCIAL_KEYS= ['family history of cancer', 'smoking'];

// Medical history
const MEDICAL_KEYS      = ['current medication', 'allergies', 'abdominal surgery',
                           'history of previous cervical cancer screening',
                           'history of abnormal vaginal discharge',
                           'history of abnormal vaginal bleeding'];

// Vital signs
const VITAL_KEYS        = ['bp', 'temp', 'hr', 'rr', 'temperature', 'blood pressure', 'pulse rate', 'respiratory rate'];

// Anthropometric
const ANTHROPOMETRIC_KEYS = ['height (cm)', 'height', 'weight (kg)', 'weight', 'bmi'];

// Physical examination (excluding breast)
const PHYSICAL_EXAM_KEYS  = ['skin', 'heent', 'chest and luns', 'chest and lungs', 'heart', 'abdomen'];

// ─── Helpers ──────────────────────────────────────────────────────────────────
function normalizeHeader(value: unknown): string {
  return String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeCell(value: unknown): string {
  return String(value ?? '').trim();
}

function findColumnIndex(headers: string[], candidates: string[]): number {
  // Exact match first
  for (const candidate of candidates) {
    const idx = headers.findIndex((h) => h === candidate);
    if (idx >= 0) return idx;
  }
  // Partial match — but require word boundary (space-separated words)
  for (const candidate of candidates) {
    const idx = headers.findIndex((h) => {
      // Only match if the candidate appears as a whole word or at start/end
      const words = h.split(/\s+/);
      return words.includes(candidate);
    });
    if (idx >= 0) return idx;
  }
  return -1;
}

function getRowCell(row: (string | number | null)[], colIndex: number): unknown {
  return colIndex < 0 ? null : row[colIndex] ?? null;
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
  if (nameCol >= 0 && normalizeCell(getRowCell(row, nameCol))) return true;
  if (idCol >= 0 && normalizeCell(getRowCell(row, idCol))) return true;
  if (dateCol >= 0 && normalizeCell(getRowCell(row, dateCol))) return true;
  return rowHasData(row);
}

function formatDateValue(value: unknown): string {
  if (value instanceof Date) return value.toLocaleDateString();
  if (typeof value === 'number' && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return new Date(parsed.y, parsed.m - 1, parsed.d).toLocaleDateString();
    }
  }
  const text = normalizeCell(value);
  return text || '(empty)';
}

function getSheetRange(sheet: XLSX.WorkSheet): XLSX.Range {
  const fallback = sheet['!ref'] ?? 'A1';
  const range = XLSX.utils.decode_range(fallback);
  for (const key of Object.keys(sheet)) {
    if (key[0] === '!') continue;
    const cell = XLSX.utils.decode_cell(key);
    if (cell.r < range.s.r) range.s.r = cell.r;
    if (cell.c < range.s.c) range.s.c = cell.c;
    if (cell.r > range.e.r) range.e.r = cell.r;
    if (cell.c > range.e.c) range.e.c = cell.c;
  }
  return range;
}

function sheetToMatrix(sheet: XLSX.WorkSheet): (string | number | null)[][] {
  const range = getSheetRange(sheet);
  const matrix: (string | number | null)[][] = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row: (string | number | null)[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = sheet[addr];
      row.push(cell == null || cell.v == null ? null : cell.v);
    }
    matrix.push(row);
  }
  return matrix;
}

function findHeaderRowIndex(matrix: (string | number | null)[][]): number {
  const scanLimit = Math.min(matrix.length, 20);
  for (let i = 0; i < scanLimit; i++) {
    const headers = (matrix[i] ?? []).map(normalizeHeader);
    const rightCol = findColumnIndex(headers, RIGHT_BREAST_KEYS);
    const leftCol  = findColumnIndex(headers, LEFT_BREAST_KEYS);
    const nameCol  = findColumnIndex(headers, NAME_KEYS);
    const dateCol  = findColumnIndex(headers, DATE_KEYS);
    if ((rightCol >= 0 || leftCol >= 0) && (nameCol >= 0 || dateCol >= 0)) return i;
  }
  return 0;
}

// ─── Main parser ──────────────────────────────────────────────────────────────
export function parseExcelBase64(base64: string): ExcelParseResult {
  const workbook = XLSX.read(base64, { type: 'base64' });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('The Excel file has no worksheets.');

  const sheet   = workbook.Sheets[sheetName];
  const matrix  = sheetToMatrix(sheet);
  if (matrix.length < 2) throw new Error('The Excel file has no data rows.');

  const headerRowIndex = findHeaderRowIndex(matrix);
  const headerRow      = matrix[headerRowIndex] ?? [];
  const headers        = headerRow.map(normalizeHeader);

  const rightBreastCol = findColumnIndex(headers, RIGHT_BREAST_KEYS);
  const leftBreastCol  = findColumnIndex(headers, LEFT_BREAST_KEYS);

  if (rightBreastCol < 0 && leftBreastCol < 0) {
    throw new Error('Could not find "Right Breast" or "Left Breast" columns in the Excel file.');
  }

  const nameCol      = findColumnIndex(headers, NAME_KEYS);
  const dateCol      = findColumnIndex(headers, DATE_KEYS);
  const screeningCol = findColumnIndex(headers, SCREENING_KEYS);
  const idCol        = findColumnIndex(headers, ID_KEYS);

  const patients: PatientRecord[] = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < matrix.length; rowIndex++) {
    const row = matrix[rowIndex] ?? [];
    if (!isPatientRow(row, nameCol, dateCol, idCol)) continue;

    const rightBreastValue = normalizeCell(getRowCell(row, rightBreastCol));
    const leftBreastValue  = normalizeCell(getRowCell(row, leftBreastCol));

    const rightIsUr = rightBreastValue.toUpperCase() === 'UR';
    const leftIsUr  = leftBreastValue.toUpperCase() === 'UR';
    const isUr      = rightIsUr && leftIsUr;

    const breastValue = isUr
      ? 'UR'
      : [
          !rightIsUr && rightBreastValue ? `R: ${rightBreastValue}` : '',
          !leftIsUr  && leftBreastValue  ? `L: ${leftBreastValue}`  : '',
        ]
          .filter(Boolean)
          .join(' | ') || '(empty)';

    const patientName =
      nameCol >= 0
        ? normalizeCell(getRowCell(row, nameCol)) || `Row ${rowIndex + 1}`
        : `Row ${rowIndex + 1}`;

    const date          = dateCol      >= 0 ? formatDateValue(getRowCell(row, dateCol))      : '(no date)';
    const screeningDate = screeningCol >= 0 ? formatDateValue(getRowCell(row, screeningCol)) : date;

    const fields: Record<string, string> = {};
    headers.forEach((header, colIndex) => {
      if (!header) return;
      
      // Skip generic Excel placeholder columns like "Column 42", "Column 43", etc.
      if (/^column\s*\d+$/.test(header)) return;
      
      const label = normalizeCell(headerRow[colIndex]) || header;
      const cellValue = getRowCell(row, colIndex);
      const normalizedValue = normalizeCell(cellValue);
      // Don't store "(empty)" placeholder for truly blank cells
      // This prevents clutter in the detail view
      fields[label] = dateCol === colIndex 
        ? formatDateValue(cellValue) 
        : (normalizedValue || '');
    });

    patients.push({
      rowIndex: rowIndex + 1,
      date,
      screeningDate,
      patientName,
      breastValue,
      rightBreastValue: rightBreastValue || '(empty)',
      leftBreastValue:  leftBreastValue  || '(empty)',
      isUr,
      fields,
    });
  }

  if (patients.length === 0) throw new Error('No data rows found below the header row.');

  return { sheetName, patients };
}

// ─── Field categorisation (used by App.tsx for display grouping) ──────────────
export type FieldCategory =
  | 'demographics'
  | 'menstrualHistory'
  | 'pregnancyHistory'
  | 'contraceptiveUse'
  | 'sexualHistory'
  | 'familySocialHistory'
  | 'medicalHistory'
  | 'vitalSigns'
  | 'anthropometric'
  | 'physicalExam'
  | 'breastRight'
  | 'breastLeft'
  | 'other';

  export function categorizeField(key: string): FieldCategory {
    const lower = key.toLowerCase().trim();
  
    // --- Check most specific groups first ---
    if (MENSTRUAL_KEYS.some((k)    => lower === k || lower.includes(k))) return 'menstrualHistory';
    if (PREGNANCY_KEYS.some((k)    => lower === k || lower.includes(k))) return 'pregnancyHistory';
    if (CONTRACEPTIVE_KEYS.some((k)=> lower === k || lower.includes(k))) return 'contraceptiveUse';
    if (SEXUAL_KEYS.some((k)       => lower === k || lower.includes(k))) return 'sexualHistory';
    if (FAMILY_SOCIAL_KEYS.some((k)=> lower === k || lower.includes(k))) return 'familySocialHistory';
    if (MEDICAL_KEYS.some((k)      => lower === k || lower.includes(k))) return 'medicalHistory';
  
    // --- Then broader groups ---
    if (DEMOGRAPHIC_KEYS.some((k) => lower === k || lower.includes(k))) return 'demographics';
    if (VITAL_KEYS.some((k)         => lower === k || lower.includes(k))) return 'vitalSigns';
    if (ANTHROPOMETRIC_KEYS.some((k)=> lower === k || lower.includes(k))) return 'anthropometric';
    if (PHYSICAL_EXAM_KEYS.some((k) => lower === k || lower.includes(k))) return 'physicalExam';
  
    // Breast fields are handled separately in the UI, but we keep them here
    if (RIGHT_BREAST_KEYS.some((k) => lower === k || lower.includes(k))) return 'breastRight';
    if (LEFT_BREAST_KEYS.some((k)  => lower === k || lower.includes(k))) return 'breastLeft';
  
    return 'other';
  }

export const CATEGORY_LABELS: Record<FieldCategory, string> = {
  demographics:        'Patient Demographics',
  menstrualHistory:    'Menstrual History',
  pregnancyHistory:    'Pregnancy History',
  contraceptiveUse:    'Contraceptive Use',
  sexualHistory:       'Sexual History',
  familySocialHistory: 'Family & Social History',
  medicalHistory:      'Medical History',
  vitalSigns:          'Vital Signs',
  anthropometric:      'Anthropometric',
  physicalExam:        'Physical Examination',
  breastRight:         'Right Breast',
  breastLeft:          'Left Breast',
  other:               'Other Details',
};

// OB-GYNE sub-categories (grouped under OB-GYNE history in the UI)
export const OB_GYNE_SUB_CATEGORIES: FieldCategory[] = [
  'menstrualHistory',
  'pregnancyHistory',
  'contraceptiveUse',
];

// Display order – "Other Details" first, then all others (excluding the separate breast fields)
export const DISPLAY_ORDER: FieldCategory[] = [
  'other',
  'demographics',
  'vitalSigns',
  'anthropometric',
  'menstrualHistory',
  'pregnancyHistory',
  'contraceptiveUse',
  'sexualHistory',
  'familySocialHistory',
  'medicalHistory',
  'physicalExam',
  'breastRight',
  'breastLeft',
];

/** Group fields by category. Skips name, date, and breast keys that are
 *  already handled separately. */
export function groupFieldsByCategory(
  fields: Record<string, string>,
): Record<FieldCategory, Record<string, string>> {
  const result = {} as Record<FieldCategory, Record<string, string>>;
  for (const cat of DISPLAY_ORDER) result[cat] = {};

  Object.entries(fields).forEach(([key, value]) => {
    const lower = key.toLowerCase();
    // Skip keys already surfaced at the top level
    if (
      lower.includes('name') ||
      lower.includes('screening date') ||
      lower.includes('referral date') ||
      lower === 'date'
    ) return;

    const cat = categorizeField(key);
    result[cat][key] = value;
  });

  return result;
}