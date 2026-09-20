import { env } from 'cloudflare:workers';

export function filesBucket() {
  const bucket=(env as unknown as {FILES?:R2Bucket}).FILES;
  if(!bucket)throw new Error('图片存储尚未连接，请稍后重试');
  return bucket;
}
