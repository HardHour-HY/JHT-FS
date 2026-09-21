import { requireUser } from '@/lib/auth';
import { filesBucket } from '@/lib/files';

export const dynamic='force-dynamic';
export async function GET(request:Request,{params}:{params:Promise<{id:string}>}){
  try{await requireUser(request);const {id}=await params;if(!/^[A-Za-z0-9_-]+$/.test(id))return new Response('图片不存在',{status:404});const object=await filesBucket().get(`products/${id}`);if(!object)return new Response('图片不存在',{status:404});const headers=new Headers({'Cache-Control':'private, max-age=3600','X-Content-Type-Options':'nosniff'});object.writeHttpMetadata(headers);return new Response(object.body,{headers});}
  catch(e){return new Response(e instanceof Error&&e.message==='UNAUTHORIZED'?'请先登录':'图片读取失败',{status:e instanceof Error&&e.message==='UNAUTHORIZED'?401:500});}
}
