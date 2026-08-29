export type PathValue = { kind: 'manual' | 'alias'; value: string; aliasId: string };
export type SkuRule = { enabled: boolean; dropSegments: number; append: string };
export type PathPair = { source: PathValue; target: PathValue; skuRule?: SkuRule };
export type Product = { id: string; code: string; name: string; mode: string; category: string; subcategory: string; image: string; note: string; paths: PathPair[] };
export type Alias = { id: string; name: string; path: string };
export type Shipment = { sku: string; quantity: number };
export type Batch = { id: string; name: string; createdAt: string; sourceRows: number; rows: Shipment[] };
export type AppState = { products: Product[]; aliases: Alias[]; batches: Batch[] };
export const emptyState: AppState = { products: [], aliases: [], batches: [] };
export const blankPath = (): PathValue => ({ kind: 'manual', value: '', aliasId: '' });
export const blankSkuRule = (): SkuRule => ({ enabled: false, dropSegments: 0, append: '' });
export const blankProduct = (code = ''): Product => ({ id: crypto.randomUUID(), code, name: '', mode: '', category: '', subcategory: '', image: '', note: '', paths: [{ source: blankPath(), target: blankPath(), skuRule: blankSkuRule() }] });
export function productCode(sku: string) { return sku.split('-')[0]; }
export function transformSku(sku: string, rule?: SkuRule) {
  if (!rule?.enabled) return sku;
  const parts = sku.split('-');
  const base = rule.dropSegments ? parts.slice(0, -rule.dropSegments).join('-') : sku;
  if (!base) throw new Error(`SKU ${sku} 删除的末尾分段过多`);
  return base + rule.append;
}
export function applyPathsToProducts(products: Product[], selectedIds: Set<string>, paths: PathPair[]) {
  return products.map(product => selectedIds.has(product.id) ? { ...product, paths: structuredClone(paths) } : product);
}
export function resolvePath(value: PathValue, aliases: Alias[]): string {
  if (value.kind === 'manual') return value.value.trim();
  const alias = aliases.find(a => a.id === value.aliasId);
  if (!alias) throw new Error('通用路径不存在，请重新选择');
  const base = alias.path.trim();
  const tail = value.value.trim();
  if (!tail) return base;
  const separator = base.includes('\\') ? '\\' : '/';
  return base.replace(/[\\/]+$/, '') + separator + tail.replace(/^[\\/]+/, '').replace(/[\\/]/g, separator);
}
export function aggregate(rows: unknown[][], skuColumn: number, quantityColumn: number, headerRow: number) {
  if (skuColumn < 0 || quantityColumn < 0 || skuColumn === quantityColumn) throw new Error('请选择不同的 SKU 货号列和数量列');
  const map = new Map<string, number>();
  const issues: string[] = [];
  let sourceRows = 0;
  rows.slice(headerRow + 1).forEach((row, index) => {
    if (row.every(v => v === null || v === undefined || String(v).trim() === '')) return;
    const sku = String(row[skuColumn] ?? '').trim();
    const raw = String(row[quantityColumn] ?? '').trim();
    const quantity = Number(raw);
    const line = headerRow + index + 2;
    if (!sku || !productCode(sku)) { issues.push(`第 ${line} 行：SKU 货号为空或格式不完整`); return; }
    if (!/^\d+(?:\.0+)?$/.test(raw) || !Number.isSafeInteger(quantity) || quantity <= 0) { issues.push(`第 ${line} 行：发货数必须是正整数（当前：${raw || '空'}）`); return; }
    const sum = (map.get(sku) ?? 0) + quantity;
    if (!Number.isSafeInteger(sum)) { issues.push(`第 ${line} 行：数量超出安全范围`); return; }
    sourceRows++;
    map.set(sku, sum);
  });
  if (!sourceRows && !issues.length) issues.push('没有找到有效的发货明细');
  const result = Array.from(map, ([sku, quantity]) => ({ sku, quantity }));
  if (!Number.isSafeInteger(result.reduce((n, r) => n + r.quantity, 0))) issues.push('总数量超出安全范围');
  return { rows: result, issues, sourceRows };
}
export function summarize(rows: Shipment[], products: Product[]) {
  const lookup = new Map(products.map(p => [p.code, p]));
  const modes = new Map<string, number>();
  let total = 0, unmatched = 0;
  const missing = new Set<string>();
  for (const row of rows) {
    total += row.quantity;
    const code = productCode(row.sku), product = lookup.get(code);
    if (!product) { unmatched += row.quantity; missing.add(code); }
    else modes.set(product.mode, (modes.get(product.mode) ?? 0) + row.quantity);
  }
  return { total, modes: Array.from(modes), unmatched, missing: Array.from(missing) };
}
export function copyRows(rows: Shipment[], state: AppState): (string | number)[][] {
  const lookup = new Map(state.products.map(p => [p.code, p]));
  const result: (string | number)[][] = [];
  for (const row of rows) {
    const product = lookup.get(productCode(row.sku));
    if (!product) throw new Error(`请先建立产品信息：${productCode(row.sku)}`);
    if (!product.paths.length) throw new Error(`${product.code} 尚未配置路径`);
    for (const pair of product.paths) {
      const source = resolvePath(pair.source, state.aliases), target = resolvePath(pair.target, state.aliases);
      if (!source || !target) throw new Error(`${product.code} 的路径尚未填写完整`);
      result.push([source, transformSku(row.sku, pair.skuRule), target, row.quantity]);
    }
  }
  return result;
}
export function toCsv(rows: (string | number)[][]): string {
  return '\uFEFF' + rows.map(row => row.map(value => {
    const text = String(value);
    // Reject formula-like text rather than silently changing file names or paths.
    if (typeof value === 'string' && /^[\s\u0000-\u001f]*[=+@-]/.test(text)) throw new Error(`CSV 中存在可能被表格软件执行的内容，请先修改：${text.slice(0, 35)}`);
    return '"' + text.replace(/"/g, '""') + '"';
  }).join(',')).join('\r\n') + '\r\n';
}
function textField(v: unknown, name: string, max = 500, required = true): asserts v is string {
  if (typeof v !== 'string' || (required && !v.trim()) || Array.from(v).length > max || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(v)) throw new Error(`${name} 格式不正确或超出长度限制`);
}
export function validateState(state: AppState): void {
  if (!state || !Array.isArray(state.products) || !Array.isArray(state.aliases) || !Array.isArray(state.batches)) throw new Error('数据格式无效');
  if (state.products.length > 10000 || state.aliases.length > 500 || state.batches.length > 30) throw new Error('超过保存上限（商品10000、路径500、批次30）');
  const unique = (values: string[], name: string) => { if (new Set(values).size !== values.length) throw new Error(`${name}不能重复`); };
  unique(state.aliases.map(a => a.id), '路径ID'); unique(state.aliases.map(a => a.name.trim()), '路径名称');
  unique(state.products.map(p => p.id), '商品ID'); unique(state.products.map(p => p.code.trim()), '产品货号');
  unique(state.batches.map(b => b.id), '批次ID');
  for (const a of state.aliases) { textField(a.id, '路径ID', 100); textField(a.name, '路径名称', 60); textField(a.path, '实际路径', 2000); }
  for (const p of state.products) {
    textField(p.id, '商品ID', 100); textField(p.code, '产品货号', 150); textField(p.name, '产品名称', 200);
    if (p.code.includes('-') || p.code !== p.code.trim()) throw new Error('产品货号应为 SKU 第一个 - 前的内容，不能包含 - 或首尾空格');
    textField(p.mode, '产品模式', 60); textField(p.category, '一级分类', 60); textField(p.subcategory, '二级分类', 60); textField(p.note, '备注', 60, false); textField(p.image, '图片链接', 2000, false);
    if (p.image) { let u: URL; try { u = new URL(p.image); } catch { throw new Error('图片链接格式不正确'); } if (!['https:', 'http:'].includes(u.protocol)) throw new Error('图片须使用 http 或 https 链接'); }
    if (!Array.isArray(p.paths) || !p.paths.length || p.paths.length > 50) throw new Error('每个商品须有1至50组对应路径');
    for (const pair of p.paths) {
      for (const path of [pair.source, pair.target]) {
        if (!path || !['manual', 'alias'].includes(path.kind)) throw new Error('路径格式无效');
        textField(path.value, '路径', 2000, false); textField(path.aliasId, '路径引用', 100, false);
        if (!resolvePath(path, state.aliases)) throw new Error('源文件夹和目标文件夹均不能为空');
      }
      if (pair.skuRule) {
        if (typeof pair.skuRule.enabled !== 'boolean' || !Number.isSafeInteger(pair.skuRule.dropSegments) || pair.skuRule.dropSegments < 0 || pair.skuRule.dropSegments > 20) throw new Error('SKU删除末尾段数须为0至20的整数');
        textField(pair.skuRule.append, 'SKU追加文字', 100, false);
        if (/[<>:"/\\|?*]/.test(pair.skuRule.append)) throw new Error('SKU追加文字不能包含 Windows 文件名禁用字符');
      }
    }
  }
  for (const b of state.batches) {
    textField(b.id, '批次ID', 100); textField(b.name, '批次名称', 300); textField(b.createdAt, '批次时间', 100);
    if (!Number.isSafeInteger(b.sourceRows) || b.sourceRows < 1 || !Array.isArray(b.rows) || !b.rows.length || b.rows.length > 20000) throw new Error('批次数据无效');
    unique(b.rows.map(r => r.sku), '批次SKU');
    for (const r of b.rows) { textField(r.sku, 'SKU', 500); if (!productCode(r.sku) || !Number.isSafeInteger(r.quantity) || r.quantity <= 0) throw new Error('发货数量或SKU无效'); }
    if (!Number.isSafeInteger(b.rows.reduce((n, r) => n + r.quantity, 0))) throw new Error('批次总数量超出安全范围');
  }
}
