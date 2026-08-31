'use client';
import { useEffect, useMemo, useState } from 'react';
import { getMsal, graphScopes } from '../lib/msal';
import { getDownloadUrl, listChildren, mediaKind, resolveRoot, type DriveItem } from '../lib/graph';

type Media = DriveItem & { path: string; kind: string; category: string };
type SeasonGroup = { key: string; number: number; episodes: Media[] };
type ShowGroup = { key: string; name: string; seasons: SeasonGroup[]; episodes: Media[] };
const categories = ['All', 'Movies', 'TV Shows', 'Anime', 'Music', 'Other'];

function catFor(path: string, name: string) { const p = (path + '/' + name).toLowerCase(); if (p.includes('/movies/')) return 'Movies'; if (p.includes('/anime/')) return 'Anime'; if (/\.(mp3|m4a|flac|wav|aac|ogg)$/i.test(name)) return 'Music'; if (/\bS\d{1,3}E\d{1,4}\b/i.test(name)) return 'TV Shows'; if (p.includes('/tv shows/') || p.includes('/tv/')) return 'TV Shows'; return 'Other'; }
function episodeInfo(name: string) { const base = name.replace(/\.[^.]+$/, ''); const m = base.match(/^(.*?)(?:\s*[-_. ]\s*)?S(\d{1,3})E(\d{1,4})\b/i); if (!m) return null; return { show: m[1].trim() || 'Unknown Show', season: Number(m[2]), episode: Number(m[3]) }; }
function showName(name: string) { return episodeInfo(name)?.show || name.replace(/\.[^.]+$/, ''); }
function episodeSort(a: Media, b: Media) { const x = episodeInfo(a.name), y = episodeInfo(b.name); if (x && y) return x.season - y.season || x.episode - y.episode; return a.name.localeCompare(b.name, undefined, { numeric: true }); }

async function thumbnailUrl(token: string, id: string) { try { const r = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(id)}/thumbnails`, { headers: { Authorization: `Bearer ${token}` }, cache: 'no-store' }); if (!r.ok) return ''; const x = await r.json(); return x?.value?.[0]?.large?.url || x?.value?.[0]?.medium?.url || x?.value?.[0]?.small?.url || ''; } catch { return ''; } }
async function tvMazeArtwork(name: string, season?: number) { try { const r = await fetch(`https://api.tvmaze.com/singlesearch/shows?q=${encodeURIComponent(name)}&embed=seasons`, { cache: 'no-store' }); if (!r.ok) return ''; const show = await r.json(); if (season && Array.isArray(show?._embedded?.seasons)) { const s = show._embedded.seasons.find((x: any) => Number(x.number) === Number(season)); if (s?.image?.original || s?.image?.medium) return s.image.original || s.image.medium; } return show?.image?.original || show?.image?.medium || ''; } catch { return ''; } }

function Poster({ token, item, label, show, season }: { token: () => Promise<string>; item?: Media; label: string; show?: string; season?: number }) {
  const [src, setSrc] = useState('');
  useEffect(() => {
    let live = true;
    (async () => {
      try {
        if (show) {
          const art = await tvMazeArtwork(show, season);
          if (live && art) { setSrc(art); return; }
        }
        if (item) {
          const t = await token();
          const thumb = await thumbnailUrl(t, item.id);
          if (live) setSrc(thumb);
        }
      } catch {}
    })();
    return () => { live = false; };
  }, [item?.id, show, season]);
  return <div className="poster" style={{ position: 'relative', overflow: 'hidden' }}>{src ? <img src={src} alt="" onError={() => setSrc('')} style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'cover' }} /> : null}<span className="badge" style={{ position: 'relative', zIndex: 1 }}>{label}</span></div>;
}

export default function App() {
 const [account,setAccount]=useState<any>(); const [items,setItems]=useState<Media[]>([]); const [q,setQ]=useState(''); const [tab,setTab]=useState('All'); const [busy,setBusy]=useState(false); const [err,setErr]=useState(''); const [playing,setPlaying]=useState<{item:Media;url:string}|null>(null); const [selectedShow,setSelectedShow]=useState<ShowGroup|null>(null); const [selectedSeason,setSelectedSeason]=useState<SeasonGroup|null>(null); const [loaded,setLoaded]=useState(false); const msal=getMsal();
 useEffect(()=>{(async()=>{try{await msal.initialize();const a=msal.getAllAccounts()[0];if(a){setAccount(a);await scan(a)}}catch(e:any){setErr(e.message||String(e))}finally{setLoaded(true)}})()},[]);
 async function token(){const a=msal.getAllAccounts()[0];if(!a)throw new Error('Please sign in first.');try{return(await msal.acquireTokenSilent({account:a,scopes:graphScopes})).accessToken}catch{return(await msal.acquireTokenPopup({account:a,scopes:graphScopes})).accessToken}}
 async function scan(a:any){setBusy(true);setErr('');setSelectedShow(null);setSelectedSeason(null);try{const t=await msal.acquireTokenSilent({account:a,scopes:graphScopes});const root=await resolveRoot(t.accessToken);const rootItems=await listChildren(t.accessToken,root.id);const out:Media[]=[];async function walk(list:DriveItem[],path:string,depth:number){if(depth>8)return;for(const it of list){const next=path+'/'+it.name;if(it.folder)await walk(await listChildren(t.accessToken,it.id),next,depth+1);else if(mediaKind(it.name))out.push({...it,path,kind:it.file?.mimeType?.startsWith('video')?'video':it.file?.mimeType?.startsWith('audio')?'audio':'image',category:catFor(path,it.name)})}}await walk(rootItems,root.name,0);setItems(out)}catch(e:any){setErr(e.message||String(e))}finally{setBusy(false)}}
 async function login(){setErr('');try{await msal.initialize();const r=await msal.loginPopup({scopes:graphScopes});setAccount(r.account);await scan(r.account)}catch(e:any){setErr(e.message||String(e))}}
 async function logout(){await msal.logoutPopup({postLogoutRedirectUri:window.location.origin});setAccount(undefined);setItems([]);setSelectedShow(null);setSelectedSeason(null)}
 const filtered=useMemo(()=>items.filter(x=>(tab==='All'||x.category===tab)&&x.name.toLowerCase().includes(q.toLowerCase())),[items,tab,q]);
 const showGroups=useMemo<ShowGroup[]>(()=>{const map=new Map<string,ShowGroup>();items.filter(x=>x.category==='TV Shows').forEach(ep=>{const name=showName(ep.name),key=name.toLowerCase();if(!map.has(key))map.set(key,{key,name,seasons:[],episodes:[]});const show=map.get(key)!;show.episodes.push(ep);const info=episodeInfo(ep.name);const sn=info?.season??1;let season=show.seasons.find(s=>s.number===sn);if(!season){season={key:`${key}-s${sn}`,number:sn,episodes:[]};show.seasons.push(season)}season.episodes.push(ep)});for(const s of map.values())for(const season of s.seasons)season.episodes.sort(episodeSort);return [...map.values()].map(s=>({...s,seasons:s.seasons.sort((a,b)=>a.number-b.number)})).filter(s=>!q.trim()||s.name.toLowerCase().includes(q.toLowerCase())||s.episodes.some(e=>e.name.toLowerCase().includes(q.toLowerCase()))).sort((a,b)=>a.name.localeCompare(b.name))},[items,q]);
 async function play(item:Media){try{const t=await token();const url=await getDownloadUrl(t,item.id);setPlaying({item,url})}catch(e:any){setErr(e.message||String(e))}}
 function mediaCard(item:Media){return <div className="card" key={item.id} onClick={()=>item.kind!=='image'&&play(item)}><Poster token={token} item={item} label={item.kind.toUpperCase()}/><div className="cardbody"><div className="title" title={item.name}>{item.name}</div><div className="sub">{item.category} · {item.size?formatBytes(item.size):'—'}</div></div></div>}
 function groupCard(name:string,sub:string,item:Media|undefined,onClick:()=>void,label:string,show?:string,season?:number){return <div className="card" onClick={onClick}><Poster token={token} item={item} label={label} show={show} season={season}/><div className="cardbody"><div className="title">{name}</div><div className="sub">{sub}</div></div></div>}
 if(!loaded)return <div className="shell"/>;
 return <div className="shell"><header className="topbar"><div className="brand">OneDrive <span>Media</span></div><div className="actions">{account&&<span className="sub">{account.username}</span>}{account?<button className="btn" onClick={logout}>Sign out</button>:<button className="btn primary" onClick={login}>Sign in with Microsoft</button>}</div></header><main className="container"><section className="hero"><h1>Your media, organized.</h1><p>OneDrive stays the storage. This app builds a Jellyfin-style library and streams directly from OneDrive.</p></section>{!account?<div className="empty">Sign in to scan your OneDrive media library.<div className="hint">Required permission: Files.Read</div></div>:<><div className="toolbar"><input className="search" placeholder="Search movies, shows, anime…" value={q} onChange={e=>setQ(e.target.value)}/><div className="tabs">{categories.map(c=><button key={c} className={'tab '+(tab===c?'active':'')} onClick={()=>{setTab(c);setSelectedShow(null);setSelectedSeason(null)}}>{c}</button>)}<button className="tab" onClick={()=>scan(account)} disabled={busy}>{busy?'Scanning…':'↻ Rescan'}</button></div></div>{err&&<div className="error">{err}</div>}{busy&&<div className="loading">Scanning OneDrive folders…</div>}{!busy&&tab==='TV Shows'?(selectedShow?(selectedSeason?<section className="show-detail"><div className="section-head"><div><button className="backlink" onClick={()=>setSelectedSeason(null)}>← {selectedShow.name}</button><h2>Season {selectedSeason.number}</h2><div className="sub">{selectedSeason.episodes.length} {selectedSeason.episodes.length===1?'episode':'episodes'}</div></div><button className="btn" onClick={()=>setSelectedSeason(null)}>← Back to seasons</button></div><div className="grid">{selectedSeason.episodes.map(mediaCard)}</div></section>:<section className="show-detail"><div className="section-head"><div><button className="backlink" onClick={()=>setSelectedShow(null)}>← TV Shows</button><h2>{selectedShow.name}</h2><div className="sub">{selectedShow.seasons.length} {selectedShow.seasons.length===1?'season':'seasons'} · {selectedShow.episodes.length} episodes</div></div></div><div className="grid">{selectedShow.seasons.map(s=>groupCard(s.number===0?'Specials':`Season ${s.number}`,`${s.episodes.length} ${s.episodes.length===1?'episode':'episodes'}`,s.episodes[0],()=>setSelectedSeason(s),'SEASON',selectedShow.name,s.number))}</div></section>):<div className="grid">{showGroups.map(s=>groupCard(s.name,`TV Shows · ${s.seasons.length} ${s.seasons.length===1?'season':'seasons'}`,s.episodes[0],()=>{setSelectedShow(s);setSelectedSeason(null)},'TV SHOW',s.name))}</div>):!busy?<div className="grid">{filtered.map(mediaCard)}</div>:null}{!busy&&((tab==='TV Shows'&&!selectedShow&&!showGroups.length)||(tab!=='TV Shows'&&!filtered.length))&&<div className="empty">No matching media found.</div>}</>}</main>{playing&&<div className="player"><div className="playerbar"><div><strong>{playing.item.name}</strong><div className="sub">Streaming from OneDrive</div></div><button className="btn" onClick={()=>setPlaying(null)}>Close</button></div>{playing.item.kind==='video'?<video className="video" controls autoPlay playsInline src={playing.url}/>:<audio className="video" controls autoPlay src={playing.url}/>}</div>}</div>;
}
function formatBytes(n:number){const u=['B','KB','MB','GB','TB'];let i=0,x=n;while(x>=1024&&i<u.length-1){x/=1024;i++}return `${x.toFixed(i?1:0)} ${u[i]}`;}
