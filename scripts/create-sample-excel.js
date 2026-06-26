const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const headers = [
  'date',
  'ID no. (Hosp)',
  'name',
  'age',
  'address',
  'civil status',
  'no. of children / last child',
  'LMP',
  'EDD',
  'menses (blood, pain)',
  'c/o yesterday (TPM)',
  'D/C contraceptive',
  'hx of abns',
  'hx of abn/menses',
  'age of bleed/menses',
  'What menses',
  'hx of abn / pelvic',
  'spouse/partner',
  'family hx of cancer',
  'Smoking hx',
  'Current medication',
  'Allergies',
  'Abdominal surgery',
  'BP',
  'Temp',
  'PR',
  'RR',
  'Height',
  'Weight',
  'BMI',
  'P/A',
  'HEENT',
  'Chest and heart',
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
