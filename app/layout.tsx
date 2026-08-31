import './globals.css';
import type { Metadata } from 'next';

export const metadata: Metadata = { title: 'OneDrive Media Library', description: 'Private Jellyfin-style media browser for OneDrive' };
export default function RootLayout({children}:{children:React.ReactNode}){return <html lang="en"><body>{children}</body></html>}
