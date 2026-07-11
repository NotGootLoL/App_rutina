import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

export default defineConfig({
  base: '/App_rutina/',
  plugins: [
    react(), 
    tailwindcss(),
    // 🛠️ INYECTOR FORZADO DE SCRIPT PARA GITHUB PAGES
    {
      name: 'force-html-script',
      transformIndexHtml(html) {
        // Borramos cualquier etiqueta script rebelde que use .tsx
        const cleanHtml = html.replace(/<script.*?>.*?<\/script>/gi, '');
        // Inyectamos al final del body el script de producción real con la ruta base correcta
        return cleanHtml.replace(
          '</body>',
          '<script type="module" src="/App_rutina/assets/main.js"></script></body>'
        );
      }
    }
  ],
  server: {
    hmr: process.env.DISABLE_HMR !== 'true',
    watch: process.env.DISABLE_HMR === 'true' ? null : {},
  },
  build: {
    outDir: 'dist',
    assetsDir: 'assets',
    sourcemap: false,
    minify: 'esbuild',
    rollupOptions: {
      output: {
        entryFileNames: 'assets/main.js',
        chunkFileNames: 'assets/[name].js',
        assetFileNames: 'assets/[name].[ext]'
      }
    }
  }
});
