export type DriveItem = {
  id:string; name:string; size?:number; webUrl?:string;
  file?:{mimeType?:string}; folder?:{childCount?:number};
  parentReference?:{path?:string};
  ['@microsoft.graph.downloadUrl']?:string;
};

const ROOT_PATH = process.env.NEXT_PUBLIC_ONEDRIVE_ROOT_PATH || '';
const VIDEO = /\.(mp4|mkv|webm|mov|m4v|avi|ts|m2ts)$/i;
const AUDIO = /\.(mp3|m4a|flac|wav|aac|ogg)$/i;
const IMAGE = /\.(jpg|jpeg|png|webp)$/i;

async function graph<T>(token:string, url:string):Promise<T>{
  const r=await fetch(url,{headers:{Authorization:`Bearer ${token}`}});
  if(!r.ok) throw new Error(`Graph ${r.status}: ${await r.text()}`);
  return r.json();
}

export async function listChildren(token:string, itemId?:string){
  const base='https://graph.microsoft.com/v1.0/me/drive';
  const first=itemId ? `${base}/items/${encodeURIComponent(itemId)}/children` : `${base}/root/children`;
  const all:DriveItem[]=[]; let url=first;
  while(url){const data=await graph<{value:DriveItem[];['@odata.nextLink']?:string}>(token,url);all.push(...data.value);url=data['@odata.nextLink']||'';}
  return all;
}

export async function resolveRoot(token:string){
  if(!ROOT_PATH.trim()) return {id:undefined,name:'OneDrive'} as {id:string|undefined,name:string};
  const u=`https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(ROOT_PATH.trim()).replace(/%2F/g,'/')}`;
  return graph<{id:string;name:string}>(token,u);
}

export async function getDownloadUrl(token:string,id:string){
  const url=`https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(id)}?$select=id,name,@microsoft.graph.downloadUrl,file,size`;
  const x=await graph<DriveItem>(token,url); return x['@microsoft.graph.downloadUrl'];
}

export function kind(name:string){ if(VIDEO.test(name)) return 'video'; if(AUDIO.test(name)) return 'audio'; if(IMAGE.test(name)) return 'image'; return 'other'; }
export function mediaKind(name:string){return kind(name)!=='other'}
