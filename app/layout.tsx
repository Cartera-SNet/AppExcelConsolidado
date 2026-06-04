import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Consolidados de Cartera',
  description: 'Genera el consolidado de cartera en formato Excel a partir de archivos IPS',
  icons: { icon: '/favicon.svg' },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="es">
      <body>{children}</body>
    </html>
  );
}
