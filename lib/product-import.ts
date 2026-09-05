import type { Alias, PathValue, Product, SkuRule } from './shipping';

export const productHeaders = ['产品图片','产品货号','产品模式','产品名称','排版分类','产品分类','源路径填写方式','源路径或剩余路径','源通用名称','目标路径填写方式','目标路径或剩余路径','目标通用名称','备注'] as const;
export type ProductImportResult = { products: Product[]; issues: string[]; sourceRows: number; newCount: number; updateCount: number };

function mode(value: string, line: number, label: string): 'manual'|'alias'|null {
  const v=value.trim().toLowerCase();
  if (['完整','手写','manual'].includes(v)) return 'manual';
  if (['通用','别名','alias'].includes(v)) return 'alias';
  return null;
}
export function detectProductHeader(rows: string[][]) {
  return rows.slice(0,30).findIndex(row => row.includes('产品货号') && row.includes('产品名称'));
}
export function parseProducts(rows: string[][], headerRow: number, existing: Product[], aliases: Alias[], idFactory=()=>crypto.randomUUID()): ProductImportResult {
  const header=rows[headerRow]||[];
  const index=new Map(header.map((name,i)=>[name.trim(),i]));
  const missing=productHeaders.filter(h=>!index.has(h));
  if(missing.length) return {products:[],issues:[`缺少表头：${missing.join('、')}`],sourceRows:0,newCount:0,updateCount:0};
  const get=(row:string[],name:typeof productHeaders[number])=>String(row[index.get(name)!]??'').trim();
  const aliasMap=new Map(aliases.map(a=>[a.name.trim(),a]));
  const existingMap=new Map(existing.map(p=>[p.code,p]));
  const grouped=new Map<string,{base:Omit<Product,'paths'>;paths:Product['paths'];firstLine:number}>();
  const issues:string[]=[]; let sourceRows=0;
  rows.slice(headerRow+1).forEach((row,offset)=>{
    if(row.every(v=>!String(v??'').trim()))return;
    const line=headerRow+offset+2; sourceRows++;
    const code=get(row,'产品货号'), image=get(row,'产品图片'), productMode=get(row,'产品模式'), name=get(row,'产品名称'), layoutType=get(row,'排版分类'), category=get(row,'产品分类'), note=get(row,'备注'),printText=index.has('是否需要打印')?String(row[index.get('是否需要打印')!]??'').trim():'';
    if(!code||code.includes('-'))issues.push(`第 ${line} 行：产品货号不能为空且不能包含 -`);
    if(!productMode||!category)issues.push(`第 ${line} 行：产品模式和产品分类均为必填`);
    if(!['设计排版','非设计排版'].includes(layoutType))issues.push(`第 ${line} 行：排版分类必须填写“设计排版”或“非设计排版”`);
    if(printText&&!['是','否'].includes(printText))issues.push(`第 ${line} 行：是否需要打印只能填写“是”或“否”`);
    if(Array.from(note).length>60)issues.push(`第 ${line} 行：备注超过60个字符`);
    if(image){try{const url=new URL(image);if(!['http:','https:'].includes(url.protocol))throw new Error();}catch{issues.push(`第 ${line} 行：产品图片必须是 http 或 https 地址`);}}
    const makePath=(kindName:'源路径填写方式'|'目标路径填写方式',valueName:'源路径或剩余路径'|'目标路径或剩余路径',aliasName:'源通用名称'|'目标通用名称'):PathValue|null=>{
      const kind=mode(get(row,kindName),line,kindName); if(!kind){issues.push(`第 ${line} 行：${kindName}只能填写“完整”或“通用”`);return null;}
      const value=get(row,valueName), aliasLabel=get(row,aliasName);
      if(kind==='manual'){if(!value)issues.push(`第 ${line} 行：${valueName}不能为空`);return {kind,value,aliasId:''};}
      const found=aliasMap.get(aliasLabel);if(!found)issues.push(`第 ${line} 行：找不到通用路径“${aliasLabel||'空'}”`);return {kind,value,aliasId:found?.id||''};
    };
    const printRequired=printText?printText==='是':existingMap.get(code)?.printRequired??false;
    const pathColumns=['源路径填写方式','源路径或剩余路径','源通用名称','目标路径填写方式','目标路径或剩余路径','目标通用名称'] as const;
    const hasPath=pathColumns.some(column=>get(row,column)!=='');
    const source=hasPath?makePath('源路径填写方式','源路径或剩余路径','源通用名称'):null,target=hasPath?makePath('目标路径填写方式','目标路径或剩余路径','目标通用名称'):null;
    const ruleEnabled=index.has('SKU处理')&&['是','开启','yes','1'].includes(String(row[index.get('SKU处理')!]??'').trim().toLowerCase());
    const dropRaw=index.has('删除末尾段数')?String(row[index.get('删除末尾段数')!]??'').trim():'';
    const dropSegments=dropRaw===''?0:Number(dropRaw); const append=index.has('追加文字')?String(row[index.get('追加文字')!]??'').trim():'';
    if(ruleEnabled&&(!Number.isSafeInteger(dropSegments)||dropSegments<0||dropSegments>20))issues.push(`第 ${line} 行：删除末尾段数须为0至20的整数`);
    if(ruleEnabled&&/[<>:"/\\|?*]/.test(append))issues.push(`第 ${line} 行：追加文字包含 Windows 文件名禁用字符`);
    const skuRule:SkuRule={enabled:ruleEnabled,dropSegments:Number.isSafeInteger(dropSegments)?dropSegments:0,append};
    if(!code||(hasPath&&(!source||!target)))return;
    const old=existingMap.get(code); const base={id:old?.id||idFactory(),code,name,mode:productMode,layoutType:layoutType as Product['layoutType'],category,printRequired,subcategory:old?.subcategory||'',image,note};
    const prior=grouped.get(code);
    if(prior){const same=['name','mode','layoutType','category','printRequired','subcategory','image','note'].every(k=>prior.base[k as keyof typeof prior.base]===base[k as keyof typeof base]);if(!same)issues.push(`第 ${line} 行：产品 ${code} 的资料与第 ${prior.firstLine} 行不一致`);if(hasPath)prior.paths.push({source:source!,target:target!,skuRule});}
    else grouped.set(code,{base,paths:hasPath?[{source:source!,target:target!,skuRule}]:[],firstLine:line});
  });
  if(!sourceRows)issues.push('没有找到商品资料行');
  const products=Array.from(grouped.values(),g=>({...g.base,paths:g.paths}));
  const owners=new Map(existing.filter(p=>p.name.trim()).map(p=>[p.name.trim(),p.code]));
  for(const product of products){if(!product.name.trim())continue;const owner=owners.get(product.name.trim());if(owner&&owner!==product.code)issues.push(`产品名称“${product.name}”已被货号 ${owner} 使用`);owners.set(product.name.trim(),product.code);}
  return {products,issues,sourceRows,newCount:products.filter(p=>!existingMap.has(p.code)).length,updateCount:products.filter(p=>existingMap.has(p.code)).length};
}
