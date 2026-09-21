import { requireUser } from '@/lib/auth';
import { filesBucket } from '@/lib/files';

export const dynamic='force-dynamic';
const allowed=new Set(['image/jpeg','image/png','image/webp']);
const reply=(value:unknown,status=200)=>Response.json(value,{status,headers:{'Cache-Control':'no-store'}});

export async function POST(request:Request){
  try{
    await requireUser(request);const origin=request.headers.get('origin');if(origin&&origin!==new URL(request.url).origin)return reply({error:'不允许跨站上传'},403);
    const form=await request.formData(),file=form.get('file');if(!(file instanceof File))return reply({error:'请选择商品图片'},400);
    if(!allowed.has(file.type))return reply({error:'商品图片仅支持 JPG、PNG 或 WebP'},415);if(file.size<1||file.size>5_000_000)return reply({error:'商品图片须小于5 MB'},413);
    const imageKey=crypto.randomUUID();await filesBucket().put(`products/${imageKey}`,await file.arrayBuffer(),{httpMetadata:{contentType:file.type},customMetadata:{originalName:file.name.slice(0,200)}});
    return reply({imageKey,url:`/api/images/${imageKey}`},201);
  }catch(e){const message=e instanceof Error?e.message:'图片上传失败';return reply({error:message==='UNAUTHORIZED'?'请先登录':message},message==='UNAUTHORIZED'?401:500);}
}
