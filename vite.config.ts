import { defineConfig } from "vite";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const keyPath = path.join(certDir, "localhost.key");
const certPath = path.join(certDir, "localhost.crt");
const hasCerts = fs.existsSync(keyPath) && fs.existsSync(certPath);

export default defineConfig(({ command }) => {
  if (command === "serve" && !hasCerts) {
    throw new Error("Faltan los certificados de desarrollo. Ejecuta: npm run certs");
  }

  const https = hasCerts
    ? { key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) }
    : undefined;

  return {
    root: path.resolve(import.meta.dirname, "src/taskpane"),
    base: command === "build" ? "/skillsEEpowerPoint/" : "/",
    publicDir: path.resolve(import.meta.dirname, "public"),
    server: { port: 3000, strictPort: true, https },
    build: { outDir: path.resolve(import.meta.dirname, "dist"), emptyOutDir: true },
  };
});
