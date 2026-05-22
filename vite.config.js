import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { VitePWA } from 'vite-plugin-pwa'

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: 'autoUpdate',
      includeAssets: ['favicon.svg', 'icons/*.png'],
      manifest: {
        name: 'RH-Paie Pro — Gestion Paie & RH',
        short_name: 'RH-Paie Pro',
        description: 'Gestion de la paie, CNSS, ITS et ressources humaines',
        theme_color: '#1a3a6b',
        background_color: '#f0f4ff',
        display: 'standalone',
        orientation: 'portrait-primary',
        scope: '/',
        start_url: '/',
        lang: 'fr',
        icons: [
          {
            src: '/icons/icon-192.png',
            sizes: '192x192',
            type: 'image/png',
            purpose: 'any maskable'
          },
          {
            src: '/icons/icon-512.png',
            sizes: '512x512',
            type: 'image/png',
            purpose: 'any maskable'
          }
        ],
        shortcuts: [
          {
            name: 'Employés',
            short_name: 'Employés',
            url: '/?page=employees',
            icons: [{ src: '/icons/icon-96.png', sizes: '96x96' }]
          },
          {
            name: 'Fiches de paie',
            short_name: 'Paie',
            url: '/?page=payroll',
            icons: [{ src: '/icons/icon-96.png', sizes: '96x96' }]
          },
          {
            name: 'Déclarations',
            short_name: 'Décl.',
            url: '/?page=declarations',
            icons: [{ src: '/icons/icon-96.png', sizes: '96x96' }]
          }
        ],
        categories: ['business', 'finance', 'productivity'],
      },
      workbox: {
        globPatterns: ['**/*.{js,css,html,ico,png,svg,woff2}'],
        runtimeCaching: [
          {
            urlPattern: /^https:\/\/rducxnyxgqrmtyvqwfzw\.supabase\.co\/.*/i,
            handler: 'NetworkFirst',
            options: {
              cacheName: 'supabase-rh-cache',
              expiration: {
                maxEntries: 50,
                maxAgeSeconds: 60 * 60 * 24
              },
              networkTimeoutSeconds: 10,
            },
          },
        ],
      },
      devOptions: { enabled: true },
    }),
  ],
  optimizeDeps: {
    include: ['exceljs', 'jszip']
  }
})

// PWA enabled
