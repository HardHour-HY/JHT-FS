import type { Metadata } from 'next';
import './globals.css';
export const metadata: Metadata={
  title:'发货工作台',
  description:'商品资料、SKU汇总与文件复制清单，一站式发货辅助。',
  metadataBase:new URL('https://shipment-studio-0828.chatgpt-team.site'),
  icons:{icon:'/favicon.svg'},
  openGraph:{title:'发货工作台',description:'商品资料 · SKU 汇总 · 复制清单',images:[{url:'/og.png',width:1536,height:1024,alt:'发货工作台'}],locale:'zh_CN',type:'website'},
  twitter:{card:'summary_large_image',title:'发货工作台',description:'商品资料 · SKU 汇总 · 复制清单',images:['/og.png']},
};
export default function RootLayout({children}:Readonly<{children:React.ReactNode}>){return <html lang="zh-CN"><body>{children}</body></html>}
