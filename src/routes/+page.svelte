<script lang="ts">
  import { onMount } from 'svelte';
  import { getMsal, graphScopes } from '$lib/msal';
  import { getDownloadUrl, listChildren, mediaKind, resolveRoot, type DriveItem } from '$lib/graph';

  type Media = DriveItem & { path: string; kind: string; category: string };
  type SeasonGroup = { key: string; number: number; episodes: Media[] };
  type ShowGroup = { key: string; name: string; seasons: SeasonGroup[]; episodes: Media[] };
  const categories = ['All', 'Movies', 'TV Shows', 'Anime', 'Music', 'Other'];
  const msal = getMsal();

  let account: any = undefined;
  let items: Media[] = [];
  let q = '';
  let tab = 'All';
  let busy = false;
  let err = '';
  let loaded = false;
  let playing: { item: Media; url: string } | null = null;
  let selectedShow: ShowGroup | null = null;
  let selectedSeason: SeasonGroup | null = null;
  const posterCache = new Map<string, Promise<string>>();

  function catFor(path: string, name: string) {
    const p = (path + '/' + name).toLowerCase();
    if (p.includes('/movies/')) return 'Movies';
    if (p.includes('/anime/')) return 'Anime';
    if (/\.(mp3|m4a|flac|wav|aac|ogg)$/i.test(name)) return 'Music';
    if (/\bS\d{1,3}E\d{1,4}\b/i.test(name)) return 'TV Shows';
    if (p.includes('/tv shows/') || p.includes('/tv/')) return 'TV Shows';
    return 'Other';
  }

  function episodeInfo(name: string) {
    const base = name.replace(/\.[^.]+$/, '');
    const m = base.match(/^(.*?)(?:\s*[-_. ]\s*)?S(\d{1,3})E(\d{1,4})\b/i);
    return m ? { show: m[1].trim() || 'Unknown Show', season: Number(m[2]), episode: Number(m[3]) } : null;
  }

  function showName(name: string) { return episodeInfo(name)?.show || name.replace(/\.[^.]+$/, ''); }
  function episodeSort(a: Media, b: Media) {
    const x = episodeInfo(a.name), y = episodeInfo(b.name);
    return x && y ? x.season - y.season || x.episode - y.episode : a.name.localeCompare(b.name, undefined, { numeric: true });
  }

  async function thumbnailUrl(token: string, id: string) {
    try {
      const r = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(id)}/thumbnails`, { headers: { Authorization: `Bearer ${token}` } });
      if (!r.ok) return '';
      const x = await r.json();
      return x?.value?.[0]?.large?.url || x?.value?.[0]?.medium?.url || x?.value?.[0]?.small?.url || '';
    } catch { return ''; }
  }

  async function tvMazeArtwork(name: string, season?: number) {
    try {
      const r = await fetch(`https://api.tvmaze.com/singlesearch/shows?q=${encodeURIComponent(name)}&embed=seasons`);
      if (!r.ok) return '';
      const show = await r.json();
      if (season !== undefined && Array.isArray(show?._embedded?.seasons)) {
        const s = show._embedded.seasons.find((x: any) => Number(x.number) === season);
        if (s?.image?.original || s?.image?.medium) return s.image.original || s.image.medium;
      }
      return show?.image?.original || show?.image?.medium || '';
    } catch { return ''; }
  }

  function posterKey(item?: Media, show?: string, season?: number) { return `${item?.id || ''}|${show || ''}|${season ?? ''}`; }
  function poster(item?: Media, show?: string, season?: number) {
    const key = posterKey(item, show, season);
    if (!posterCache.has(key)) {
      posterCache.set(key, (async () => {
        if (show) {
          const art = await tvMazeArtwork(show, season);
          if (art) return art;
        }
        if (item) {
          const a = msal.getAllAccounts()[0];
          if (a) {
            try {
              const t = (await msal.acquireTokenSilent({ account: a, scopes: graphScopes })).accessToken;
              return await thumbnailUrl(t, item.id);
            } catch {}
          }
        }
        return '';
      })());
    }
    return posterCache.get(key)!;
  }

  async function token() {
    const a = msal.getAllAccounts()[0];
    if (!a) throw new Error('Please sign in first.');
    try { return (await msal.acquireTokenSilent({ account: a, scopes: graphScopes })).accessToken; }
    catch { return (await msal.acquireTokenPopup({ account: a, scopes: graphScopes })).accessToken; }
  }

  async function scan(a: any) {
    busy = true; err = ''; selectedShow = null; selectedSeason = null;
    try {
      const t = await msal.acquireTokenSilent({ account: a, scopes: graphScopes });
      const root = await resolveRoot(t.accessToken);
      const rootItems = await listChildren(t.accessToken, root.id);
      const out: Media[] = [];
      async function walk(list: DriveItem[], path: string, depth: number) {
        if (depth > 8) return;
        for (const it of list) {
          const next = path + '/' + it.name;
          if (it.folder) await walk(await listChildren(t.accessToken, it.id), next, depth + 1);
          else if (mediaKind(it.name)) out.push({ ...it, path, kind: it.file?.mimeType?.startsWith('video') ? 'video' : it.file?.mimeType?.startsWith('audio') ? 'audio' : 'image', category: catFor(path, it.name) });
        }
      }
      await walk(rootItems, root.name, 0);
      items = out;
      posterCache.clear();
    } catch (e: any) { err = e.message || String(e); }
    finally { busy = false; }
  }

  async function login() {
    err = '';
    try { await msal.initialize(); const r = await msal.loginPopup({ scopes: graphScopes }); account = r.account; await scan(r.account); }
    catch (e: any) { err = e.message || String(e); }
  }

  async function logout() {
    await msal.logoutPopup({ postLogoutRedirectUri: window.location.origin });
    account = undefined; items = []; selectedShow = null; selectedSeason = null;
  }

  async function play(item: Media) {
    try { const url = await getDownloadUrl(await token(), item.id); playing = { item, url }; }
    catch (e: any) { err = e.message || String(e); }
  }

  function formatBytes(n: number) {
    const u = ['B', 'KB', 'MB', 'GB', 'TB']; let i = 0, x = n;
    while (x >= 1024 && i < u.length - 1) { x /= 1024; i++; }
    return `${x.toFixed(i ? 1 : 0)} ${u[i]}`;
  }

  $: filtered = items.filter((x) => (tab === 'All' || x.category === tab) && x.name.toLowerCase().includes(q.toLowerCase()));
  $: showGroups = (() => {
    const map = new Map<string, ShowGroup>();
    for (const ep of items.filter((x) => x.category === 'TV Shows')) {
      const name = showName(ep.name), key = name.toLowerCase();
      if (!map.has(key)) map.set(key, { key, name, seasons: [], episodes: [] });
      const show = map.get(key)!; show.episodes.push(ep);
      const sn = episodeInfo(ep.name)?.season ?? 1;
      let season = show.seasons.find((s) => s.number === sn);
      if (!season) { season = { key: `${key}-s${sn}`, number: sn, episodes: [] }; show.seasons.push(season); }
      season.episodes.push(ep);
    }
    for (const show of map.values()) for (const season of show.seasons) season.episodes.sort(episodeSort);
    return [...map.values()].map((s) => ({ ...s, seasons: s.seasons.sort((a, b) => a.number - b.number) }))
      .filter((s) => !q.trim() || s.name.toLowerCase().includes(q.toLowerCase()) || s.episodes.some((e) => e.name.toLowerCase().includes(q.toLowerCase())))
      .sort((a, b) => a.name.localeCompare(b.name));
  })();

  onMount(async () => {
    try {
      await msal.initialize();
      const a = msal.getAllAccounts()[0];
      if (a) { account = a; await scan(a); }
    } catch (e: any) { err = e.message || String(e); }
    finally { loaded = true; }
  });
</script>

<svelte:head><title>OneDrive Media</title><meta name="description" content="Lightweight OneDrive media library" /></svelte:head>

{#if !loaded}
  <div class="shell"></div>
{:else}
  <div class="shell">
    <header class="topbar">
      <div class="brand">OneDrive <span>Media</span></div>
      <div class="actions">
        {#if account}<span class="sub">{account.username}</span><button class="btn" onclick={logout}>Sign out</button>{:else}<button class="btn primary" onclick={login}>Sign in with Microsoft</button>{/if}
      </div>
    </header>

    <main class="container">
      <section class="hero"><h1>Your media, organized.</h1><p>OneDrive stays the storage. This app builds a Jellyfin-style library and streams directly from OneDrive.</p></section>

      {#if !account}
        <div class="empty">Sign in to scan your OneDrive media library.<div class="hint">Required permission: Files.Read</div></div>
      {:else}
        <div class="toolbar">
          <input class="search" placeholder="Search movies, shows, anime…" bind:value={q} />
          <div class="tabs">
            {#each categories as c}<button class:active={tab === c} class="tab" onclick={() => { tab = c; selectedShow = null; selectedSeason = null; }}>{c}</button>{/each}
            <button class="tab" disabled={busy} onclick={() => scan(account)}>{busy ? 'Scanning…' : '↻ Rescan'}</button>
          </div>
        </div>

        {#if err}<div class="error">{err}</div>{/if}
        {#if busy}<div class="loading">Scanning OneDrive folders…</div>{/if}

        {#if !busy && tab === 'TV Shows'}
          {#if selectedShow}
            {#if selectedSeason}
              <section class="show-detail"><div class="section-head"><div><button class="backlink" onclick={() => selectedSeason = null}>← {selectedShow.name}</button><h2>Season {selectedSeason.number === 0 ? 'Specials' : selectedSeason.number}</h2><div class="sub">{selectedSeason.episodes.length} {selectedSeason.episodes.length === 1 ? 'episode' : 'episodes'}</div></div></div><div class="grid">{#each selectedSeason.episodes as item (item.id)}<div class="card" onclick={() => play(item)}>{#await poster(item) then src}<div class="poster">{#if src}<img src={src} alt="" />{/if}<span class="badge">{item.kind.toUpperCase()}</span></div>{/await}<div class="cardbody"><div class="title">{item.name}</div><div class="sub">{formatBytes(item.size || 0)}</div></div></div>{/each}</div></section>
            {:else}
              <section class="show-detail"><div class="section-head"><div><button class="backlink" onclick={() => selectedShow = null}>← TV Shows</button><h2>{selectedShow.name}</h2><div class="sub">{selectedShow.seasons.length} {selectedShow.seasons.length === 1 ? 'season' : 'seasons'} · {selectedShow.episodes.length} episodes</div></div></div><div class="grid">{#each selectedShow.seasons as season (season.key)}<div class="card" onclick={() => selectedSeason = season}>{#await poster(season.episodes[0], selectedShow.name, season.number) then src}<div class="poster">{#if src}<img src={src} alt="" />{/if}<span class="badge">SEASON</span></div>{/await}<div class="cardbody"><div class="title">{season.number === 0 ? 'Specials' : `Season ${season.number}`}</div><div class="sub">{season.episodes.length} {season.episodes.length === 1 ? 'episode' : 'episodes'}</div></div></div>{/each}</div></section>
            {/if}
          {:else}
            <div class="grid">{#each showGroups as show (show.key)}<div class="card" onclick={() => { selectedShow = show; selectedSeason = null; }}>{#await poster(show.episodes[0], show.name) then src}<div class="poster">{#if src}<img src={src} alt="" />{/if}<span class="badge">TV SHOW</span></div>{/await}<div class="cardbody"><div class="title">{show.name}</div><div class="sub">{show.seasons.length} {show.seasons.length === 1 ? 'season' : 'seasons'} · {show.episodes.length} episodes</div></div></div>{/each}</div>
          {/if}
          {#if !selectedShow && !showGroups.length}<div class="empty">No matching media found.</div>{/if}
        {:else if !busy}
          <div class="grid">{#each filtered as item (item.id)}<div class="card" onclick={() => item.kind !== 'image' && play(item)}>{#await poster(item) then src}<div class="poster">{#if src}<img src={src} alt="" />{/if}<span class="badge">{item.kind.toUpperCase()}</span></div>{/await}<div class="cardbody"><div class="title" title={item.name}>{item.name}</div><div class="sub">{item.category} · {item.size ? formatBytes(item.size) : '—'}</div></div></div>{/each}</div>
          {#if !filtered.length}<div class="empty">No matching media found.</div>{/if}
        {/if}
      {/if}
    </main>

    {#if playing}<div class="player"><div class="playerbar"><div><strong>{playing.item.name}</strong><div class="sub">Streaming from OneDrive</div></div><button class="btn" onclick={() => playing = null}>Close</button></div>{#if playing.item.kind === 'video'}<video class="video" controls autoplay playsinline src={playing.url}></video>{:else}<audio class="video" controls autoplay src={playing.url}></audio>{/if}</div>{/if}
  </div>
{/if}

<style>
  :global(*){box-sizing:border-box}:global(html),:global(body){margin:0;background:#0b0d10;color:#f4f6f8;font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}:global(button),:global(input){font:inherit}
  .shell{min-height:100vh}.topbar{height:64px;border-bottom:1px solid #252b33;display:flex;align-items:center;justify-content:space-between;padding:0 24px;position:sticky;top:0;background:rgba(11,13,16,.94);backdrop-filter:blur(14px);z-index:10}.brand{font-weight:800}.brand span{color:#6ea8fe}.actions{display:flex;gap:10px;align-items:center}.btn{border:1px solid #252b33;background:#1a1f26;color:#f4f6f8;padding:9px 14px;border-radius:10px;cursor:pointer}.btn.primary{background:#eaf2ff;color:#07101c;border-color:#eaf2ff}.btn:disabled{opacity:.55}.container{max-width:1400px;margin:auto;padding:28px 24px 60px}.hero{padding:20px 0 8px}.hero h1{font-size:42px;line-height:1.05;margin:0 0 10px;letter-spacing:-.04em}.hero p,.sub{color:#9aa5b1}.hero p{margin:0;max-width:760px}.toolbar{display:flex;gap:12px;flex-wrap:wrap;margin:28px 0}.search{flex:1;min-width:240px;background:#11151a;border:1px solid #252b33;color:#f4f6f8;padding:12px 14px;border-radius:12px;outline:none}.tabs{display:flex;gap:8px;flex-wrap:wrap}.tab{background:#11151a;color:#9aa5b1;border:1px solid #252b33;padding:10px 13px;border-radius:999px;cursor:pointer}.tab.active{color:#f4f6f8;background:#1b222a}.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:18px}.card{border:1px solid #252b33;background:#14181d;border-radius:14px;overflow:hidden;cursor:pointer;transition:transform .15s,border-color .15s}.card:hover{transform:translateY(-2px);border-color:#3b4653}.poster{aspect-ratio:2/3;background:linear-gradient(135deg,#222a34,#11151a);display:flex;align-items:flex-end;padding:14px;position:relative;overflow:hidden}.poster img{position:absolute;inset:0;width:100%;height:100%;object-fit:cover}.badge{position:relative;z-index:1;background:rgba(0,0,0,.55);border:1px solid rgba(255,255,255,.12);padding:6px 8px;border-radius:8px;font-size:12px}.cardbody{padding:13px}.title{font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.empty,.loading{padding:50px 0;color:#9aa5b1;text-align:center}.hint{font-size:13px;margin-top:8px}.error{background:#241418;color:#ffb7c0;border:1px solid #5d2931;padding:12px;border-radius:10px;margin:14px 0}.section-head{display:flex;justify-content:space-between;align-items:flex-start;margin:20px 0}.section-head h2{margin:10px 0 4px}.backlink{border:0;background:none;color:#6ea8fe;cursor:pointer;padding:0}.player{position:fixed;inset:0;background:rgba(0,0,0,.94);z-index:50;display:flex;flex-direction:column}.playerbar{height:60px;display:flex;align-items:center;justify-content:space-between;padding:0 18px}.video{width:100%;height:calc(100vh - 60px);background:#000;object-fit:contain}
  @media(max-width:700px){.topbar{padding:0 14px}.container{padding:18px 14px 40px}.hero h1{font-size:34px}.grid{grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.search{min-width:100%}.sub{font-size:12px}}
</style>
