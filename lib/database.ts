import { env } from 'cloudflare:workers';
import { emptyState } from './shipping';
export async function database() {
  const db = (env as unknown as { DB: D1Database }).DB;
  if (!db) throw new Error('数据库尚未连接，请稍后重试');
  await db.prepare('CREATE TABLE IF NOT EXISTS workspace (id INTEGER PRIMARY KEY, version INTEGER NOT NULL DEFAULT 0, payload TEXT NOT NULL, updated_at TEXT NOT NULL, updated_by TEXT NOT NULL)').run();
  await db.prepare('INSERT OR IGNORE INTO workspace (id,version,payload,updated_at,updated_by) VALUES (1,0,?,?,?)').bind(JSON.stringify(emptyState), new Date().toISOString(), 'system').run();
  return db;
}
