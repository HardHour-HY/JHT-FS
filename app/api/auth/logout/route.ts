import { clearSessionCookie,ensureAuthTables,tokenHashFromRequest } from '@/lib/auth';
export async function POST(request:Request){const hash=await tokenHashFromRequest(request);if(hash){const db=await ensureAuthTables();await db.prepare('DELETE FROM app_sessions WHERE token_hash=?').bind(hash).run();}return Response.json({ok:true},{headers:{'Set-Cookie':clearSessionCookie}});}
