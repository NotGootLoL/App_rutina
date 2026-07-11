import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

export default defineConfig(() => {
  return {
    base: '/App_rutina/',
    plugins: [react(), tailwindcss()],
    server: {
      hmr: process.env.DISABLE_HMR !== 'true',
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
    // 🛠️ FORZAMOS LAS RUTAS RELATIVAS DIRECTAS DESDE LA RAÍZ DEL PROYECTO
    build: {
      outDir: 'dist',
      rollupOptions: {
        input: 'index.html', // Ruta directa como string, sin usar 'path' ni '__dirname'
        output: {
          entryFileNames: 'assets/[name].js',
          chunkFileNames: 'assets/[name].js',
          assetFileNames: 'assets/[name].[ext]'
        }
      }
    }
  };
});
