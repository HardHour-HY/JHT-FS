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
  // The universal build avoids downloading and starting a Worker for small
  // operational spreadsheets. Parsing remains asynchronous at the call site.
  const { default: readWorkbook } = await import('read-excel-file/universal');
  const sheets = await readWorkbook(await file.arrayBuffer());
  return sheets.map(({ sheet, data }) => {
    if (data.length > 20001 || data.some(row => row.length > 300)) throw new Error('单个工作表最多支持 20000 行明细、300 列');
    const rows = data.map(row => row.map(value => value instanceof Date ? value.toISOString().slice(0, 10) : value === null ? '' : String(value)));
    return { name: sheet, rows };
  });
}
export function detectColumns(rows: string[][]) {
  let headerRow = rows.slice(0, 30).findIndex(r => r.some(v => v.trim() === 'SKU货号'));
  if (headerRow < 0) headerRow = 0;
  const headers = rows[headerRow] || [];
  const find=(names:string[])=>headers.findIndex(v=>names.includes(v.trim()));
  return { headerRow, skuColumn: find(['SKU货号','SKU','sku']), quantityColumn: find(['发货数','发货数量','数量']), warehouseColumn:find(['收货仓库','仓库']),shopColumn:find(['店铺','店铺名称']),packageColumn:find(['包裹号','包裹编号']) };
}
export function download(content: BlobPart, name: string, type: string) {
  const url = URL.createObjectURL(new File([content], name, { type }));
  const anchor = document.createElement('a'); anchor.href = url; anchor.download = name; anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
export async function encodeGbkCsv(content: string): Promise<Uint8Array<ArrayBuffer>> {
  await import('@kayahr/text-encoding/encodings/gbk');
  const { TextEncoder: LegacyTextEncoder } = await import('@kayahr/text-encoding/no-encodings');
  const plain = content.replace(/^\uFEFF/, '');
  let encoded: Uint8Array;
  try { encoded = new LegacyTextEncoder('gbk').encode(plain); }
  catch { throw new Error('CSV 中包含 GBK 无法表示的特殊字符，请修改商品路径或 SKU 后再导出'); }
  const bytes = new Uint8Array(encoded.byteLength); bytes.set(encoded);
  const roundTrip = new TextDecoder('gbk').decode(bytes);
  if (roundTrip !== plain) throw new Error('CSV 中包含 GBK 无法表示的特殊字符，请修改商品路径或 SKU 后再导出');
  return bytes;
}
export async function downloadCsv(content: string, name: string) {
  if (!name.toLowerCase().endsWith('.csv')) name += '.csv';
  const bytes = await encodeGbkCsv(content);
  download(bytes.buffer, name, 'text/csv;charset=gbk');
}
export async function exportProductTemplate(categories: string[] = [], modes: string[] = []) {
  const ExcelJS = (await import('exceljs')).default;
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet('商品资料导入');
  const headers = ['产品图片','产品货号','产品模式','产品名称','排版分类','产品分类','源路径填写方式','源路径或剩余路径','源通用名称','目标路径填写方式','目标路径或剩余路径','目标通用名称','SKU处理','删除末尾段数','追加文字','备注'];
  sheet.addRow(headers);
  sheet.columns = headers.map((header, index) => ({ header, key: String(index), width: index === 0 ? 32 : index >= 6 && index <= 11 ? 22 : 16 }));
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF276447' } };
  sheet.getRow(1).height = 28; sheet.views = [{ state: 'frozen', ySplit: 1 }]; sheet.autoFilter = 'A1:P1';
  ['B','I','L'].forEach(column => sheet.getColumn(column).numFmt = '@');
  if(categories.length){const options=book.addWorksheet('产品分类选项',{state:'veryHidden'});categories.forEach(value=>options.addRow([value]));(sheet as any).dataValidations.add('F2:F10000',{type:'list',allowBlank:false,formulae:[`'产品分类选项'!$A$1:$A$${categories.length}`]});}
  if(modes.length){const options=book.addWorksheet('产品模式选项',{state:'veryHidden'});modes.forEach(value=>options.addRow([value]));(sheet as any).dataValidations.add('C2:C10000',{type:'list',allowBlank:false,formulae:[`'产品模式选项'!$A$1:$A$${modes.length}`]});}
  const guide = book.addWorksheet('填写说明');
  guide.addRows([
    ['填写说明','内容'],
    ['一件商品多组路径','每组路径填写一行，并重复填写相同的产品货号和商品资料。'],
    ['路径填写方式','填写“完整”或“通用”。完整：路径列填写完整路径；通用：路径列填写剩余路径，同时填写已建立的通用名称。'],
    ['每条路径处理SKU','SKU处理填写“是”时生效。删除末尾段数按“-”分段；例如 QFBKQS30-1000-1PC 删除1段得到 QFBKQS30-1000，再追加 -F 得到 QFBKQS30-1000-F。不处理可留空。'],
    ['已有产品货号','导入会更新商品资料，并用表格中的路径替换该商品现有的全部路径。'],
    ['必填字段','产品货号、产品模式、排版分类、产品分类。非设计排版商品还必须填写源路径和目标路径；产品名称可留空，填写后不能与其他商品重复。产品货号必须唯一。'],
    ['排版分类','只能填写“设计排版”或“非设计排版”，必须二选一。设计排版商品无需填写路径及 SKU 处理相关单元格。'],
    ['备注','最多60个中文字符。图片必须是 http 或 https 网络地址。'],
  ]);
  guide.columns = [{width:22},{width:95}]; guide.getRow(1).font = {bold:true,color:{argb:'FFFFFFFF'}}; guide.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF276447'}};
  guide.eachRow(row=>{row.alignment={vertical:'top',wrapText:true};});
  download(await book.xlsx.writeBuffer() as BlobPart, '商品资料批量导入模板.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
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
export async function exportModeSummary(batch:Batch,products:Product[]){
  if(!batch.details?.length)throw new Error('当前批次没有仓库、店铺和包裹号明细，请重新上传发货表格');
  const ExcelJS=(await import('exceljs')).default,book=new ExcelJS.Workbook(),sheet=book.addWorksheet('汇总'),lookup=new Map(products.map(p=>[p.code,p]));
  const warehouses=new Map<string,Map<string,number>>(),shops=new Map<string,Map<string,Set<string>>>();
  for(const row of batch.details){const product=lookup.get(productCode(row.sku));if(!product)throw new Error(`请先建立产品信息：${productCode(row.sku)}`);const mode=product.mode;let wm=warehouses.get(mode);if(!wm)warehouses.set(mode,wm=new Map());wm.set(row.warehouse,(wm.get(row.warehouse)||0)+row.quantity);let sm=shops.get(mode);if(!sm)shops.set(mode,sm=new Map());let packages=sm.get(row.shop);if(!packages)sm.set(row.shop,packages=new Set());packages.add(row.packageNo);}
  let start=1;for(const mode of warehouses.keys()){const wm=warehouses.get(mode)!,sm=shops.get(mode)||new Map<string,Set<string>>(),height=Math.max(wm.size,sm.size,1);sheet.getCell(start,1).value=`${mode}收货仓库`;sheet.getCell(start,2).value=`${mode}发货数`;sheet.getCell(start,4).value=`${mode}店铺`;sheet.getCell(start,5).value=`${mode}包裹数量`;for(const cell of [sheet.getCell(start,1),sheet.getCell(start,2),sheet.getCell(start,4),sheet.getCell(start,5)]){cell.font={bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF276447'}};}[...wm].forEach(([name,total],i)=>{sheet.getCell(start+1+i,1).value=name;sheet.getCell(start+1+i,2).value=total;});[...sm].forEach(([name,set],i)=>{sheet.getCell(start+1+i,4).value=name;sheet.getCell(start+1+i,5).value=set.size;});start+=height+3;}
  sheet.columns=[{width:28},{width:16},{width:4},{width:28},{width:16}];sheet.views=[{state:'frozen',ySplit:1}];download(await book.xlsx.writeBuffer() as BlobPart,`${batch.name.replace(/\.[^.]+$/,'')}-模式汇总.xlsx`,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}
