import { defineConfig, loadEnv } from 'vite';
import { process } from 'node:process';

export default defineConfig(({ mode }) => {
  // Vercel will set process.env variables during the build process.
  const env = loadEnv(mode, process.cwd(), '');

  return {
    // Vite config
    define: {
      // This replaces `process.env.API_KEY` in the source code with the actual
      // environment variable value at build time.
      'process.env.API_KEY': JSON.stringify(env.API_KEY)
    },
    build: {
      outDir: 'dist'
    }
  }
})