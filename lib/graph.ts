export type DriveItem = {
  id: string;
  name: string;
  size?: number;
  webUrl?: string;
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  parentReference?: { path?: string };
  ['@microsoft.graph.downloadUrl']?: string;
};

const ROOT_PATH = process.env.NEXT_PUBLIC_ONEDRIVE_ROOT_PATH || '';
const VIDEO = /\.(mp4|mkv|webm|mov|m4v|avi|ts|m2ts)$/i;
const AUDIO = /\.(mp3|m4a|flac|wav|aac|ogg)$/i;
const IMAGE = /\.(jpg|jpeg|png|webp|gif)$/i;

async function graph<T>(token: string, url: string): Promise<T> {
  const r = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    cache: 'no-store',
  });

  if (!r.ok) {
    throw new Error(`Graph ${r.status}: ${await r.text()}`);
  }

  return r.json();
}

export async function listChildren(token: string, itemId?: string) {
  const base = 'https://graph.microsoft.com/v1.0/me/drive';
  const first = itemId
    ? `${base}/items/${encodeURIComponent(itemId)}/children`
    : `${base}/root/children`;

  const all: DriveItem[] = [];
  let url = first;

  while (url) {
    const data = await graph<{
      value: DriveItem[];
      ['@odata.nextLink']?: string;
    }>(token, url);

    all.push(...data.value);
    url = data['@odata.nextLink'] || '';
  }

  return all;
}

export async function resolveRoot(token: string) {
  if (!ROOT_PATH.trim()) {
    return { id: undefined, name: 'OneDrive' } as {
      id: string | undefined;
      name: string;
    };
  }

  const path = encodeURIComponent(ROOT_PATH.trim()).replace(/%2F/g, '/');
  const u = `https://graph.microsoft.com/v1.0/me/drive/root:/${path}`;
  return graph<{ id: string; name: string }>(token, u);
}

export async function getDownloadUrl(token: string, id: string) {
  // Microsoft Graph documents @microsoft.graph.downloadUrl as the
  // browser-friendly pre-authenticated URL for JavaScript playback.
  const params = new URLSearchParams({
    '$select': 'id,@microsoft.graph.downloadUrl',
  });

  const url =
    `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(id)}` +
    `?${params.toString()}`;

  const x = await graph<DriveItem>(token, url);
  const downloadUrl = x['@microsoft.graph.downloadUrl'];

  if (!downloadUrl) {
    throw new Error(
      'Microsoft Graph returned the file, but no playback URL. ' +
      'Make sure the signed-in account has access to this file and Files.Read is granted.'
    );
  }

  return downloadUrl;
}

export function kind(name: string) {
  if (VIDEO.test(name)) return 'video';
  if (AUDIO.test(name)) return 'audio';
  if (IMAGE.test(name)) return 'image';
  return 'other';
}

export function mediaKind(name: string) {
  return kind(name) !== 'other';
}
