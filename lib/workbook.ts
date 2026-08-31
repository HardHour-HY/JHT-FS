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
  return { headerRow, skuColumn: find(['SKU货号','SKU','sku']), quantityColumn: find(['发货数','发货数量','数量']), warehouseColumn:find(['收货仓库','仓库']),shopColumn:find(['店铺','店铺名称']),packageColumn:find(['包裹号','包裹编号']),customIdColumn:find(['定制ID','定制 Id','定制id','Custom ID','custom_id']) };
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
  const headers = ['产品图片','产品货号','产品模式','产品名称','排版分类','产品分类','是否需要打印','源路径填写方式','源路径或剩余路径','源通用名称','目标路径填写方式','目标路径或剩余路径','目标通用名称','SKU处理','删除末尾段数','追加文字','备注'];
  sheet.addRow(headers);
  sheet.columns = headers.map((header, index) => ({ header, key: String(index), width: index === 0 ? 32 : index >= 7 && index <= 12 ? 22 : 16 }));
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF276447' } };
  sheet.getRow(1).height = 28; sheet.views = [{ state: 'frozen', ySplit: 1 }]; sheet.autoFilter = 'A1:Q1';
  ['B','J','M'].forEach(column => sheet.getColumn(column).numFmt = '@');
  (sheet as any).dataValidations.add('G2:G10000',{type:'list',allowBlank:false,formulae:['"是,否"']});
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
    ['是否需要打印','只能填写“是”或“否”；新商品不填写时按“否”处理。'],
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
  let start=1;
  for(const mode of warehouses.keys()){
    const wm=warehouses.get(mode)!,sm=shops.get(mode)||new Map<string,Set<string>>(),height=Math.max(wm.size,sm.size,1);
    sheet.getCell(start,1).value=`${mode}收货仓库`;sheet.getCell(start,2).value=`${mode}发货数`;sheet.getCell(start,4).value=`${mode}店铺`;sheet.getCell(start,5).value=`${mode}包裹数量`;
    for(const cell of [sheet.getCell(start,1),sheet.getCell(start,2),sheet.getCell(start,4),sheet.getCell(start,5)]){cell.font={bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF276447'}};}
    [...wm].forEach(([name,total],i)=>{sheet.getCell(start+1+i,1).value=name;sheet.getCell(start+1+i,2).value=total;});
    [...sm].forEach(([name,set],i)=>{sheet.getCell(start+1+i,4).value=name;sheet.getCell(start+1+i,5).value=set.size;});
    const warehouseTotalRow=start+1+wm.size,shopTotalRow=start+1+sm.size;
    sheet.getCell(warehouseTotalRow,1).value='合计';sheet.getCell(warehouseTotalRow,2).value=[...wm.values()].reduce((sum,value)=>sum+value,0);
    const allPackages=new Set<string>();for(const set of sm.values())for(const packageNo of set)allPackages.add(packageNo);
    sheet.getCell(shopTotalRow,4).value='合计';sheet.getCell(shopTotalRow,5).value=allPackages.size;
    for(const cell of [sheet.getCell(warehouseTotalRow,1),sheet.getCell(warehouseTotalRow,2),sheet.getCell(shopTotalRow,4),sheet.getCell(shopTotalRow,5)]){cell.font={bold:true};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE8F1EC'}};}
    start+=height+4;
  }
  sheet.columns=[{width:28},{width:16},{width:4},{width:28},{width:16}];sheet.views=[{state:'frozen',ySplit:1}];
  const paperSheet=book.addWorksheet('纸质产品'),paperDetails=batch.details.flatMap(row=>{const product=lookup.get(productCode(row.sku));return product?.category.includes('纸质')?[{row,product}]:[];}).sort((a,b)=>[a.product.category,a.row.warehouse,a.row.packageNo,a.row.sku].join('\u0000').localeCompare([b.product.category,b.row.warehouse,b.row.packageNo,b.row.sku].join('\u0000'),'zh-CN'));
  const table1Headers=['产品分类','收货仓库','包裹号','SKU货号','发货数'],table2Headers=['产品分类','产品货号','SKU货号','发货数'],table3Headers=['产品分类','收货仓库','发货数量','包裹数量'];
  const setHeaders=(column:number,headers:string[])=>headers.forEach((header,index)=>{const cell=paperSheet.getCell(1,column+index);cell.value=header;cell.font={bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF276447'}};});
  const setTotal=(row:number,column:number,values:(string|number)[])=>values.forEach((value,index)=>{const cell=paperSheet.getCell(row,column+index);cell.value=value;cell.font={bold:true};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE8F1EC'}};});
  const mergeRepeated=(rows:(string|number)[][],startColumn:number,dimensionColumns:number)=>{for(let dimension=0;dimension<dimensionColumns;dimension++){let from=0;while(from<rows.length){let to=from+1;while(to<rows.length&&rows[to].slice(0,dimension+1).every((value,index)=>value===rows[from][index]))to++;if(to-from>1){paperSheet.mergeCells(from+2,startColumn+dimension,to+1,startColumn+dimension);paperSheet.getCell(from+2,startColumn+dimension).alignment={vertical:'middle'};}from=to;}}};
  setHeaders(1,table1Headers);setHeaders(7,table2Headers);setHeaders(12,table3Headers);
  paperDetails.forEach(({row,product},index)=>{const values=[product.category,row.warehouse,row.packageNo,row.sku,row.quantity];values.forEach((value,column)=>paperSheet.getCell(index+2,column+1).value=value);});
  mergeRepeated(paperDetails.map(({row,product})=>[product.category,row.warehouse,row.packageNo,row.sku,row.quantity]),1,4);
  setTotal(paperDetails.length+2,1,['合计','','','',paperDetails.reduce((sum,item)=>sum+item.row.quantity,0)]);
  const skuTotals=new Map<string,{category:string;code:string;sku:string;quantity:number}>();
  for(const {row,product} of paperDetails){const key=`${product.category}\u0000${product.code}\u0000${row.sku}`,prior=skuTotals.get(key);if(prior)prior.quantity+=row.quantity;else skuTotals.set(key,{category:product.category,code:product.code,sku:row.sku,quantity:row.quantity});}
  const skuRows=[...skuTotals.values()].sort((a,b)=>[a.category,a.code,a.sku].join('\u0000').localeCompare([b.category,b.code,b.sku].join('\u0000'),'zh-CN')).map(item=>[item.category,item.code,item.sku,item.quantity] as (string|number)[]);
  skuRows.forEach((values,index)=>values.forEach((value,column)=>paperSheet.getCell(index+2,column+7).value=value));mergeRepeated(skuRows,7,3);
  setTotal(skuTotals.size+2,7,['合计','','',[...skuTotals.values()].reduce((sum,item)=>sum+item.quantity,0)]);
  const warehouseTotals=new Map<string,{category:string;warehouse:string;quantity:number;packages:Set<string>}>();
  for(const {row,product} of paperDetails){const key=`${product.category}\u0000${row.warehouse}`,prior=warehouseTotals.get(key);if(prior){prior.quantity+=row.quantity;prior.packages.add(row.packageNo);}else warehouseTotals.set(key,{category:product.category,warehouse:row.warehouse,quantity:row.quantity,packages:new Set([row.packageNo])});}
  const warehouseRows=[...warehouseTotals.values()].sort((a,b)=>[a.category,a.warehouse].join('\u0000').localeCompare([b.category,b.warehouse].join('\u0000'),'zh-CN')).map(item=>[item.category,item.warehouse,item.quantity,item.packages.size] as (string|number)[]);
  warehouseRows.forEach((values,index)=>values.forEach((value,column)=>paperSheet.getCell(index+2,column+12).value=value));mergeRepeated(warehouseRows,12,2);
  const allPaperPackages=new Set(paperDetails.map(item=>item.row.packageNo));
  setTotal(warehouseTotals.size+2,12,['合计','',[...warehouseTotals.values()].reduce((sum,item)=>sum+item.quantity,0),allPaperPackages.size]);
  paperSheet.columns=[{width:20},{width:22},{width:24},{width:32},{width:14},{width:4},{width:20},{width:20},{width:32},{width:14},{width:4},{width:20},{width:22},{width:14},{width:14}];paperSheet.views=[{state:'frozen',ySplit:1}];
  const calendarSheet=book.addWorksheet('定制挂历'),calendarDetails=batch.details.flatMap(row=>{const product=lookup.get(productCode(row.sku));return product?.category.includes('定制挂历')?[{row,product}]:[];}).sort((a,b)=>[a.product.category,a.row.warehouse,a.row.packageNo,a.row.sku,a.row.customId||''].join('\u0000').localeCompare([b.product.category,b.row.warehouse,b.row.packageNo,b.row.sku,b.row.customId||''].join('\u0000'),'zh-CN'));
  if(calendarDetails.some(item=>!item.row.customId?.trim()))throw new Error('定制挂历商品缺少定制ID，请重新上传并选择“定制ID”列');
  const calendarHeaders=(column:number,headers:string[])=>headers.forEach((header,index)=>{const cell=calendarSheet.getCell(1,column+index);cell.value=header;cell.font={bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF276447'}};});
  const calendarTotal=(row:number,column:number,values:(string|number)[])=>values.forEach((value,index)=>{const cell=calendarSheet.getCell(row,column+index);cell.value=value;cell.font={bold:true};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE8F1EC'}};});
  const mergeCalendar=(rows:(string|number)[][],startColumn:number,dimensionColumns:number)=>{for(let dimension=0;dimension<dimensionColumns;dimension++){let from=0;while(from<rows.length){let to=from+1;while(to<rows.length&&rows[to].slice(0,dimension+1).every((value,index)=>value===rows[from][index]))to++;if(to-from>1){calendarSheet.mergeCells(from+2,startColumn+dimension,to+1,startColumn+dimension);calendarSheet.getCell(from+2,startColumn+dimension).alignment={vertical:'middle'};}from=to;}}};
  calendarHeaders(1,['产品分类','收货仓库','包裹号','SKU货号','定制ID','发货数量']);calendarHeaders(8,['产品分类','收货仓库','发货数量','包裹数量']);
  const calendarRows=calendarDetails.map(({row,product})=>[product.category,row.warehouse,row.packageNo,row.sku,row.customId||'',row.quantity] as (string|number)[]);
  calendarRows.forEach((values,index)=>values.forEach((value,column)=>calendarSheet.getCell(index+2,column+1).value=value));mergeCalendar(calendarRows,1,5);
  calendarTotal(calendarRows.length+2,1,['合计','','','','',calendarDetails.reduce((sum,item)=>sum+item.row.quantity,0)]);
  const calendarWarehouseTotals=new Map<string,{category:string;warehouse:string;quantity:number;packages:Set<string>}>();
  for(const {row,product} of calendarDetails){const key=`${product.category}\u0000${row.warehouse}`,prior=calendarWarehouseTotals.get(key);if(prior){prior.quantity+=row.quantity;prior.packages.add(row.packageNo);}else calendarWarehouseTotals.set(key,{category:product.category,warehouse:row.warehouse,quantity:row.quantity,packages:new Set([row.packageNo])});}
  const calendarWarehouseRows=[...calendarWarehouseTotals.values()].sort((a,b)=>[a.category,a.warehouse].join('\u0000').localeCompare([b.category,b.warehouse].join('\u0000'),'zh-CN')).map(item=>[item.category,item.warehouse,item.quantity,item.packages.size] as (string|number)[]);
  calendarWarehouseRows.forEach((values,index)=>values.forEach((value,column)=>calendarSheet.getCell(index+2,column+8).value=value));mergeCalendar(calendarWarehouseRows,8,2);
  calendarTotal(calendarWarehouseRows.length+2,8,['合计','',calendarWarehouseRows.reduce((sum,row)=>sum+Number(row[2]),0),new Set(calendarDetails.map(item=>item.row.packageNo)).size]);
  calendarSheet.columns=[{width:22},{width:22},{width:24},{width:32},{width:24},{width:14},{width:4},{width:22},{width:22},{width:14},{width:14}];calendarSheet.views=[{state:'frozen',ySplit:1}];
  download(await book.xlsx.writeBuffer() as BlobPart,`${batch.name.replace(/\.[^.]+$/,'')}-模式汇总.xlsx`,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}
