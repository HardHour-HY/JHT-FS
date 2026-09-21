'use client';

import {useEffect,useState} from 'react';
import {ExternalLink,Link2,LoaderCircle,PackageCheck} from 'lucide-react';
import {Button} from '@/components/ui/button';
import {Input} from '@/components/ui/input';

declare global { interface Window { chrome?: any } }

type ExtensionReply={ok:boolean;data?:{version?:string;job?:{status?:string;message?:string}|null};error?:string};

export default function DianxiaomiBridge(){
  const [extensionId,setExtensionId]=useState(''),[state,setState]=useState('尚未连接'),[busy,setBusy]=useState(false);
  useEffect(()=>setExtensionId(localStorage.getItem('dxm-extension-id')||''),[]);
  async function send(type:string):Promise<ExtensionReply>{
    const id=extensionId.trim();
    if(!id)throw new Error('请先填写 Chrome 扩展 ID。');
    const runtime=window.chrome?.runtime;
    if(!runtime?.sendMessage)throw new Error('请使用 Chrome 打开发货工作台。');
    return new Promise((resolve,reject)=>runtime.sendMessage(id,{type},(reply:ExtensionReply)=>{
      const error=runtime.lastError;if(error)return reject(new Error('未连接到插件：'+error.message));
      resolve(reply);
    }));
  }
  async function connect(){
    setBusy(true);try{
      localStorage.setItem('dxm-extension-id',extensionId.trim());
      const reply=await send('SHIPMENT_STUDIO_STATUS');if(!reply?.ok)throw new Error(reply?.error||'插件未返回状态。');
      const job=reply.data?.job;setState(job?.message?`已连接 · ${job.message}`:`已连接 · 插件 ${reply.data?.version||''}`);
    }catch(error){setState(error instanceof Error?error.message:'连接失败');}finally{setBusy(false);}
  }
  async function open(){
    setBusy(true);try{const reply=await send('SHIPMENT_STUDIO_OPEN');if(!reply?.ok)throw new Error(reply?.error||'无法打开店小秘。');setState('已打开店小秘插件小窗口。');}
    catch(error){setState(error instanceof Error?error.message:'打开失败');}finally{setBusy(false);}
  }
  return <section className="panel pluginbridge" aria-label="店小秘插件连接">
    <div className="pluginbridge-icon"><PackageCheck size={21}/></div>
    <div className="pluginbridge-copy"><h2>店小秘插件</h2><p>连接后可从此工作台打开店小秘的普单处理和 Y2 非纸质处理小窗口。</p><small>{state}</small></div>
    <div className="pluginbridge-controls"><Input value={extensionId} onChange={e=>setExtensionId(e.target.value)} placeholder="粘贴 Chrome 扩展 ID" aria-label="Chrome 扩展 ID"/><Button variant="outline" disabled={busy} onClick={()=>void connect()}><Link2 size={16}/>{busy?'正在连接':'连接插件'}</Button><Button disabled={busy} onClick={()=>void open()}><ExternalLink size={16}/>打开店小秘</Button></div>
    <p className="pluginbridge-help">首次使用：在 Chrome 扩展管理页打开“开发者模式”，复制“店小秘 TEMU · SKU 与加工文件”的 ID 并粘贴到这里。</p>
  </section>;
}
