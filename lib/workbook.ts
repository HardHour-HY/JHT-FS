import type { Batch, Product } from './shipping';
import { productCode } from './shipping';
export type SheetData = { name: string; rows: string[][] };
export function parseCsv(text: string): string[][] {
  const rows: string[][] = []; let row: string[] = [], value = '', quoted = false;
  text = text.replace(/^\uFEFF/, '');
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"') { if (quoted && text[i + 1] === '"') { value += '"'; i++; } else if (quoted || value === '') quoted = !quoted; else value += c; }
    else if (c === ',' && !quoted) { row.push(value); value = ''; }
    else if ((c === '\n' || c === '\r') && !quoted) { row.push(value); rows.push(row); row = []; value = ''; if (c === '\r' && text[i + 1] === '\n') i++; }
    else value += c;
  }
  if (quoted) throw new Error('CSV 引号未闭合，请检查文件格式');
  if (value || row.length) { row.push(value); rows.push(row); }
  return rows;
}
export async function readFile(file: File): Promise<SheetData[]> {
  if (file.size > 15 * 1024 * 1024) throw new Error('文件不能超过 15 MB');
  if (/\.csv$/i.test(file.name)) {
    const bytes = await file.arrayBuffer(); let decoded: string;
    try { decoded = new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
    catch { decoded = new TextDecoder('gb18030').decode(bytes); }
    return [{ name: 'CSV', rows: parseCsv(decoded) }];
  }
  if (!/\.xlsx$/i.test(file.name)) throw new Error('请选择 .xlsx 或 .csv 文件；旧版 .xls 请先另存为 .xlsx');
  const ExcelJS = (await import('exceljs')).default;
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(await file.arrayBuffer());
  return book.worksheets.map(sheet => {
    if (sheet.rowCount > 20001 || sheet.columnCount > 300) throw new Error('单个工作表最多支持 20000 行明细、300 列');
    const rows: string[][] = [];
    sheet.eachRow({ includeEmpty: true }, row => {
      const values: string[] = [];
      for (let i = 1; i <= Math.max(sheet.columnCount, row.cellCount); i++) {
        const cell = row.getCell(i);
        if (cell.type === ExcelJS.ValueType.Formula) { values.push('【公式单元格，请先转为值】'); continue; }
        if (typeof cell.value === 'number' && /^0+$/.test(cell.numFmt)) values.push(String(cell.value).padStart(cell.numFmt.length, '0'));
        else values.push(cell.text || '');
      }
      rows.push(values);
    });
    return { name: sheet.name, rows };
  });
}
export function detectColumns(rows: string[][]) {
  let headerRow = rows.slice(0, 30).findIndex(r => r.some(v => v.trim() === 'SKU货号'));
  if (headerRow < 0) headerRow = 0;
  const headers = rows[headerRow] || [];
  return { headerRow, skuColumn: headers.findIndex(v => ['SKU货号', 'SKU', 'sku'].includes(v.trim())), quantityColumn: headers.findIndex(v => ['发货数', '发货数量', '数量'].includes(v.trim())) };
}
export function download(content: BlobPart, name: string, type: string) {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const anchor = document.createElement('a'); anchor.href = url; anchor.download = name; anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
export async function exportSummary(batch: Batch, products: Product[]) {
  const ExcelJS = (await import('exceljs')).default;
  const book = new ExcelJS.Workbook(); const sheet = book.addWorksheet('发货汇总');
  sheet.columns = [{ header: 'SKU货号', key: 'sku', width: 40 }, { header: '数量', key: 'quantity', width: 12 }, { header: '产品货号', key: 'code', width: 32 }];
  const codes = new Set(products.map(p => p.code));
  for (const row of batch.rows) sheet.addRow({ sku: row.sku, quantity: row.quantity, code: codes.has(productCode(row.sku)) ? productCode(row.sku) : `请建立产品信息：${productCode(row.sku)}` });
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF276447' } };
  sheet.getRow(1).height = 28; sheet.views = [{ state: 'frozen', ySplit: 1 }];
  sheet.autoFilter = 'A1:C1'; sheet.getColumn(1).numFmt = '@'; sheet.getColumn(3).numFmt = '@';
  const buffer = await book.xlsx.writeBuffer();
  download(buffer as BlobPart, batch.name.replace(/\.[^.]+$/, '') + '-发货汇总.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}
