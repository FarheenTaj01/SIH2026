import { Analytics } from '@vercel/analytics/next'
import type { Metadata, Viewport } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'Nexus Material Intelligence | SIH 2026',
  description: 'Explainable material harmonization prototype for SIH 2026 Problem Statement 26099.',
  generator: 'SIH 2026 Prototype',
}

export const viewport: Viewport = {
  colorScheme: 'light',
  themeColor: '#f7fafb',
}

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="en" className="bg-background"><body className="antialiased">{children}{process.env.NODE_ENV === 'production' && <Analytics />}</body></html>
}
