import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import basicSsl from "@vitejs/plugin-basic-ssl";

// GitHub Pages serves project sites below /<repository>/; development remains at /.
const isPagesBuild = process.env.GITHUB_ACTIONS === "true";

export default defineConfig({
  base: isPagesBuild ? "/word-docXray/" : "/",
  plugins: [react(), basicSsl()],
  server: { port: 3000 },
  preview: { port: 3000 }
});
