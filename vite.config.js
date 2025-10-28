import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react'; // or vue / svelte depending on your stack
import { VitePWA } from 'vite-plugin-pwa';

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: 'autoUpdate',  // auto updates service worker
      includeAssets: ['favicon.ico', 'robots.txt', 'apple-touch-icon.png'],
      manifest: {
        name: 'My Awesome Web App',
        short_name: 'MyApp',
        description: 'An installable web app built with Vite 🚀',
        theme_color: '#0d6efd',
        background_color: '#ffffff',
        display: 'standalone',
        start_url: '/',
        icons: [
          {
            src: '/icons/icon-192x192.png',
            sizes: '192x192',
            type: 'image/png',
          },
          {
            src: '/icons/icon-512x512.png',
            sizes: '512x512',
            type: 'image/png',
          },
        ],
      },
      workbox: {
        // Custom caching rules (optional)
        runtimeCaching: [
          {
            urlPattern: /^https:\/\/your-api-domain\.com\/.*$/,
            handler: 'NetworkFirst',
            options: {
              cacheName: 'api-cache',
              expiration: { maxEntries: 20, maxAgeSeconds: 86400 },
            },
          },
        ],
      },
    }),
  ],
});
