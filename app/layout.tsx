import { Analytics } from '@vercel/analytics/next'
import type { Metadata, Viewport } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'Nexus Material Intelligence',
  description: 'Explainable material harmonization workspace for enterprise material data.',
  generator: 'Nexus Material Intelligence',
}

export const viewport: Viewport = {
  colorScheme: 'light',
  themeColor: '#f7fafb',
}

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="en" className="bg-background"><body className="antialiased">{children}{process.env.NODE_ENV === 'production' && <Analytics />}</body></html>
}
