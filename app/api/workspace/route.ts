import { database } from '@/lib/database';
import { requireUser } from '@/lib/auth';
import { normalizeState, validateState, type AppState } from '@/lib/shipping';
export const dynamic = 'force-dynamic';
const reply = (value: unknown, status = 200) => Response.json(value, { status, headers: { 'Cache-Control': 'no-store' } });
export async function GET(request: Request) {
  try {
    await requireUser(request);
    const db = await database();
    const row = await db.prepare('SELECT version,payload,updated_at,updated_by FROM workspace WHERE id=1').first<{version:number;payload:string;updated_at:string;updated_by:string}>();
    if (!row) throw new Error('工作空间读取失败');
    return reply({ version: row.version, state: normalizeState(JSON.parse(row.payload)), updatedAt: row.updated_at });
  } catch (e) { const message=e instanceof Error?e.message:'读取失败';return reply({ error: message==='UNAUTHORIZED'?'请先登录':message }, message==='UNAUTHORIZED'?401:503); }
}
export async function PUT(request: Request) {
  try {
    const user = await requireUser(request);
    const origin = request.headers.get('origin');
    if (origin && origin !== new URL(request.url).origin) return reply({ error: '不允许跨站修改' }, 403);
    if (!request.headers.get('content-type')?.includes('application/json')) return reply({ error: '需要 JSON 数据' }, 415);
    const body = await request.text();
    if (body.length > 3_000_000) return reply({ error: '数据超过保存容量，请删除不需要的历史批次' }, 413);
    const parsed = JSON.parse(body) as { state: AppState; version: number };const state=normalizeState(parsed.state),version=parsed.version;
    if (!Number.isSafeInteger(version) || version < 0) return reply({ error: '版本号无效' }, 400);
    validateState(state);
    const db = await database();
    const updatedAt = new Date().toISOString();
    const result = await db.prepare('UPDATE workspace SET payload=?,version=version+1,updated_at=?,updated_by=? WHERE id=1 AND version=?').bind(JSON.stringify(state), updatedAt, user.username, version).run();
    if (result.meta.changes !== 1) return reply({ error: '资料已被其他窗口修改。请先刷新资料，再重新保存；你的表单内容仍保留。' }, 409);
    return reply({ version: version + 1, updatedAt });
  } catch (e) { const message=e instanceof Error?e.message:'保存失败，请重试';return reply({ error: message==='UNAUTHORIZED'?'请先登录':message },message==='UNAUTHORIZED'?401:400); }
}
