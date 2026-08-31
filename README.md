# OneDrive Media Library

A lightweight Jellyfin-style media browser that keeps your actual media in OneDrive and uses Vercel only for the web app. Browser playback requests the Microsoft Graph `@microsoft.graph.downloadUrl` and plays the file directly from OneDrive, rather than proxying video bytes through Vercel.

## 1. Create the Microsoft Entra app

1. Open Microsoft Entra admin center → App registrations → New registration.
2. Choose a normal web/public app registration suitable for your Microsoft account type. For a personal OneDrive account, enable personal Microsoft accounts; for a Microsoft 365 work/school account, use the organization option.
3. Add a **Single-page application (SPA)** redirect URI equal to your deployed Vercel URL, e.g. `https://your-app.vercel.app` (also add `http://localhost:3000` for local development).
4. API permissions → Microsoft Graph → Delegated permissions → add **Files.Read**. No write permission is required.
5. Copy the Application (client) ID.

The app uses MSAL's authorization-code/PKCE browser flow and delegated `Files.Read`. Microsoft documents PKCE for SPA authorization-code flow. See https://learn.microsoft.com/en-us/entra/identity-platform/msal-authentication-flows

## 2. Configure Vercel

Set these environment variables:

```text
NEXT_PUBLIC_AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
NEXT_PUBLIC_AZURE_TENANT_ID=common
NEXT_PUBLIC_ONEDRIVE_ROOT_PATH=Media
```

`NEXT_PUBLIC_ONEDRIVE_ROOT_PATH=Media` means the scanner starts inside a OneDrive folder named `Media`. You can leave it blank to scan the OneDrive root.

Recommended folder layout:

```text
Media/
  Movies/
  TV Shows/
  Anime/
  Music/
```

## 3. Run locally

```bash
npm install
cp .env.example .env.local
# edit .env.local
npm run dev
```

Open http://localhost:3000.

## 4. Deploy

Push this folder to GitHub and import it into Vercel, or use the Vercel CLI. Add the same environment variables in Vercel.

## Notes

- Graph's `children` API is used to recursively enumerate folders.
- Playback obtains the `@microsoft.graph.downloadUrl` for the selected file. Microsoft notes this is the browser-friendly way to retrieve the preauthenticated file URL for JavaScript apps.
- The MVP does not transcode/remux video. Browser/device codec support therefore matters (MP4/H.264/AAC is the safest target).
- This MVP intentionally does not rename/move/delete your OneDrive files.
