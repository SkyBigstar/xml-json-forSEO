import * as XLSX from 'xlsx';

const workbook: XLSX.WorkBook = XLSX.readFile('data/1.xlsx');
const worksheet: XLSX.WorkSheet = workbook.Sheets[workbook.SheetNames[0]];
const data: any[] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

console.log(data);
