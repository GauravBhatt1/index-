'use client';
import {useEffect,useMemo,useState} from 'react';
import {getMsal,graphScopes} from '../lib/msal';
import {getDownloadUrl,listChildren,mediaKind,resolveRoot,type DriveItem} from '../lib/graph';

type Media=DriveItem & {path:string; kind:string; category:string};
const categories=['All','Movies','TV Shows','Anime','Music','Other'];

function catFor(path:string,name:string){const p=(path+'/'+name).toLowerCase(); if(p.includes('/movies/'))return 'Movies'; if(p.includes('/tv shows/')||p.includes('/tv/'))return 'TV Shows'; if(p.includes('/anime/'))return 'Anime'; if(/\.(mp3|m4a|flac|wav|aac|ogg)$/i.test(name))return 'Music'; return 'Other';}

export default function App(){
 const [account,setAccount]=useState<any>(); const [items,setItems]=useState<Media[]>([]); const [q,setQ]=useState(''); const [tab,setTab]=useState('All'); const [busy,setBusy]=useState(false); const [err,setErr]=useState(''); const [playing,setPlaying]=useState<{item:Media,url:string}|null>(null); const [loaded,setLoaded]=useState(false);
 const msal=getMsal();
 useEffect(()=>{ (async()=>{try{await msal.initialize(); const a=msal.getAllAccounts()[0]; if(a){setAccount(a); await scan(a)} }catch(e:any){setErr(e.message||String(e))}finally{setLoaded(true)}})()},[]);
 async function token(){const a=msal.getAllAccounts()[0]; if(!a) throw new Error('Please sign in first.'); try{return (await msal.acquireTokenSilent({account:a,scopes:graphScopes})).accessToken}catch{return (await msal.acquireTokenPopup({account:a,scopes:graphScopes})).accessToken}}
 async function scan(a:any){setBusy(true);setErr(''); try{const t=await msal.acquireTokenSilent({account:a,scopes:graphScopes}); const root=await resolveRoot(t.accessToken); const rootItems=await listChildren(t.accessToken,root.id); const out:Media[]=[];
   async function walk(list:DriveItem[],path:string,depth:number){ if(depth>8)return; for(const it of list){const next=path+'/'+it.name; if(it.folder){const children=await listChildren(t.accessToken,it.id); await walk(children,next,depth+1)} else if(mediaKind(it.name)){out.push({...it,path,kind:it.file?.mimeType?.startsWith('video')?'video':it.file?.mimeType?.startsWith('audio')?'audio':'image',category:catFor(path,it.name)})}}}
   await walk(rootItems,root.name,0); setItems(out); }catch(e:any){setErr(e.message||String(e))}finally{setBusy(false)}}
 async function login(){setErr('');try{await msal.initialize(); const r=await msal.loginPopup({scopes:graphScopes});setAccount(r.account);await scan(r.account)}catch(e:any){setErr(e.message||String(e))}}
 async function logout(){await msal.logoutPopup({postLogoutRedirectUri:window.location.origin});setAccount(undefined);setItems([])}
 const filtered=useMemo(()=>items.filter(x=>(tab==='All'||x.category===tab)&&x.name.toLowerCase().includes(q.toLowerCase())),[items,tab,q]);
 async function play(item:Media){try{const t=await token(); const url=await getDownloadUrl(t,item.id); if(!url)throw new Error('OneDrive did not return a playback URL.'); setPlaying({item,url})}catch(e:any){setErr(e.message||String(e))}}
 return <div className="shell"><header className="topbar"><div className="brand">OneDrive <span>Media</span></div><div className="actions">{account&&<span className="sub">{account.username}</span>}{account?<button className="btn" onClick={logout}>Sign out</button>:<button className="btn primary" onClick={login}>Sign in with Microsoft</button>}</div></header>
 <main className="container"><section className="hero"><h1>Your media, organized.</h1><p>OneDrive stays the storage. This app only builds a Jellyfin-style library and sends playback requests directly to OneDrive.</p></section>
 {!account?<div className="empty">Sign in to scan your OneDrive media library.<div className="hint">Required permission: Files.Read</div></div>:<><div className="toolbar"><input className="search" placeholder="Search movies, shows, anime…" value={q} onChange={e=>setQ(e.target.value)}/><div className="tabs">{categories.map(c=><button key={c} className={'tab '+(tab===c?'active':'')} onClick={()=>setTab(c)}>{c}</button>)}<button className="tab" onClick={()=>scan(account)} disabled={busy}>{busy?'Scanning…':'↻ Rescan'}</button></div></div>
 {err&&<div className="error">{err}</div>}{busy&&<div className="loading">Scanning OneDrive folders…</div>}{!busy&&<div className="grid">{filtered.map(item=><div className="card" key={item.id} onClick={()=>item.kind==='video'||item.kind==='audio'?play(item):null}><div className="poster"><span className="badge">{item.kind.toUpperCase()}</span></div><div className="cardbody"><div className="title" title={item.name}>{item.name}</div><div className="sub">{item.category} · {item.size?formatBytes(item.size):'—'}</div></div></div>)}</div>}{!busy&&!filtered.length&&<div className="empty">No matching media found.</div>}</>}
 </main>
 {playing&&<div className="player"><div className="playerbar"><div><strong>{playing.item.name}</strong><div className="sub">Streaming from OneDrive</div></div><button className="btn" onClick={()=>setPlaying(null)}>Close</button></div>{playing.item.kind==='video'?<video className="video" controls autoPlay playsInline src={playing.url}/>:<audio className="video" controls autoPlay src={playing.url}/>}</div>}</div>
}
function formatBytes(n:number){const u=['B','KB','MB','GB','TB'];let i=0;let x=n;while(x>=1024&&i<u.length-1){x/=1024;i++}return `${x.toFixed(i?1:0)} ${u[i]}`}
