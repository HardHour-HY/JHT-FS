import { currentUser } from '@/lib/auth';
export const dynamic='force-dynamic';
export async function GET(request:Request){const user=await currentUser(request);return user?Response.json({user},{headers:{'Cache-Control':'no-store'}}):Response.json({error:'请先登录'},{status:401,headers:{'Cache-Control':'no-store'}});}
