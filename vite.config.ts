import { defineConfig } from "vite";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const keyPath = path.join(certDir, "localhost.key");
const certPath = path.join(certDir, "localhost.crt");
const https =
  fs.existsSync(keyPath) && fs.existsSync(certPath)
    ? { key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) }
    : undefined;

export default defineConfig(({ command }) => ({
  root: "src/taskpane",
  base: command === "build" ? "/skillsEEpowerPoint/" : "/",
  publicDir: path.resolve(__dirname, "public"),
  server: { port: 3000, strictPort: true, https },
  build: { outDir: path.resolve(__dirname, "dist"), emptyOutDir: true },
}));
