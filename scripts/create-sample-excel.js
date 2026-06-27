const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const headers = [
  'Date',
  'ID No. (Hosp)',
  'Name',
  'Age',
  'Address',
  'Civil Status',
  'No. of Children / Last Child',
  'LMP',
  'EDD',
  'Menses (Blood, Pain)',
  'C/O Yesterday (TPM)',
  'D/C Contraceptive',
  'Hx of Abns',
  'Hx of Abn/Menses',
  'Age of Bleed/Menses',
  'What Menses',
  'Hx of Abn / Pelvic',
  'Spouse/Partner',
  'Family Hx of Cancer',
  'Smoking Hx',
  'Current Medication',
  'Allergies',
  'Abdominal Surgery',
  'BP',
  'Temp',
  'PR',
  'RR',
  'Height',
  'Weight',
  'BMI',
  'P/A',
  'HEENT',
  'Chest and Heart',
  'Breast',
  'Abdomen',
  'Pelvic',
];

const sampleRow = [
  '09/01/2021',
  'A0281',
  'Shana Mukaram',
  26,
  '',
  'Married',
  5,
  '14 Oct 2020',
  '',
  '5/30 Regular',
  '3/24 NAD-C',
  'No',
  '',
  'No',
  '',
  'No',
  '',
  '',
  'No',
  '',
  'Losartan 50mg, Amlodipine 5',
  'None',
  '',
  '120/110',
  '',
  119,
  20,
  1.57,
  '',
  32,
  'UR',
  'NAD',
  'NAD',
  'UR',
  'NAD',
  'UR',
];

const worksheet = XLSX.utils.aoa_to_sheet([headers, sampleRow]);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

const outputDir = path.join(__dirname, '..', 'sample-data');
fs.mkdirSync(outputDir, { recursive: true });

const outputPath = path.join(outputDir, 'sample-clinical.xlsx');
XLSX.writeFile(workbook, outputPath);

console.log(`Sample Excel created: ${outputPath}`);