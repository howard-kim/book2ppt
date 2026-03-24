import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'IDML → PPT 변환기',
  description: 'InDesign IDML 파일을 PowerPoint로 변환합니다.',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  )
}
