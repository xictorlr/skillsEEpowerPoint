# Fase 1: Andamiaje + Alignment + Size — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Dejar un add-in de PowerPoint (Office.js) instalable en Mac y Windows, desplegado en GitHub Pages, con el panel lateral y las secciones Alignment y Size funcionando sobre la selección.

**Architecture:** Task pane en TypeScript sin framework. `src/core/` contiene geometría pura sobre modelos planos (`ShapeModel`, `Patch`) con tests Vitest. `src/office/` traduce selección ↔ `PowerPoint.Shape` y ejecuta un único `context.sync()` por comando. `src/commands/` registra comandos declarativos que el panel renderiza.

**Tech Stack:** Node 26, TypeScript 5, Vite 8, Vitest 5, `@types/office-js`, `office-addin-manifest`, `office-addin-dev-certs`, GitHub Actions + GitHub Pages.

**Spec:** `docs/superpowers/specs/2026-10-14-slide-tools-design.md`

**Nombre del producto:** "Slide Tools" (proveedor: Victor Lopez). URL de producción: `https://xictorlr.github.io/skillsEEpowerPoint/`.

---

## Mapa de archivos

| Archivo | Responsabilidad |
|---|---|
| `package.json`, `tsconfig.json`, `vite.config.ts`, `vitest.config.ts` | Toolchain |
| `public/assets/icon-{16,32,64,80}.png` | Iconos del manifest (generados por script) |
| `scripts/make-icons.py` | Genera los PNG con stdlib de Python |
| `manifest.xml` | Manifest de producción (GitHub Pages) |
| `scripts/make-dev-manifest.sh` | Genera `manifest.dev.xml` (localhost) desde `manifest.xml` |
| `scripts/sideload-mac.sh` | Copia un manifest a la carpeta `wef` de PowerPoint Mac |
| `src/core/types.ts` | `Rect`, `ShapeModel`, `SlideSize`, `SelectionContext`, `Patch` |
| `src/core/geometry.ts` | `alignShapes`, `alignToSlide`, `distribute`, `swapPositions`, `matchSize`, `stretchToEdge`, `arrangeMatrix` |
| `src/core/png.ts` | `pngDimensions(base64)` y `slideSizeFromAspect` |
| `src/office/runtime.ts` | `runCommand`: envuelve `PowerPoint.run`, normaliza errores |
| `src/office/selection.ts` | `readSelection(context)` → modelo + proxies |
| `src/office/slideSize.ts` | `getSlideSize(context)` con caché y fallback 960×540 |
| `src/office/apply.ts` | `applyPatches(proxies, patches)` |
| `src/commands/types.ts` | `Command`, `CommandOptions`, `SectionId`, `CommandResult` |
| `src/commands/alignment.ts`, `src/commands/size.ts` | Definiciones de comandos de cada sección |
| `src/commands/registry.ts` | Lista ordenada de secciones y comandos |
| `src/commands/execute.ts` | `executeCommand(cmd, options)`: leer → calcular → aplicar → sync |
| `src/taskpane/index.html`, `styles.css` | Panel |
| `src/taskpane/icons.ts` | SVG inline por id de icono |
| `src/taskpane/settings.ts` | Opciones persistidas en `localStorage` (matriz) |
| `src/taskpane/panel.ts` | Renderiza secciones/botones, estado habilitado, avisos |
| `src/taskpane/main.ts` | `Office.onReady`, evento de cambio de selección |
| `tests/core/*.test.ts`, `tests/office/*.test.ts`, `tests/commands/*.test.ts` | Tests |
| `.github/workflows/deploy.yml` | Build + test + deploy a Pages |
| `README.md` | Instalación, desarrollo, comandos |

---

### Task 1: Toolchain y panel "hola mundo"

**Files:**
- Create: `package.json`, `tsconfig.json`, `vite.config.ts`, `vitest.config.ts`
- Create: `src/taskpane/index.html`, `src/taskpane/main.ts`, `src/taskpane/styles.css`
- Create: `tests/smoke.test.ts`

- [ ] **Step 1: Crear `package.json`**

```json
{
  "name": "slide-tools",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "author": "Victor Lopez",
  "description": "Panel de herramientas de alineación, tamaño, formato y texto para PowerPoint (Office.js).",
  "scripts": {
    "dev": "vite",
    "build": "tsc --noEmit && vite build",
    "preview": "vite preview",
    "test": "vitest run",
    "test:watch": "vitest",
    "typecheck": "tsc --noEmit",
    "validate": "office-addin-manifest validate manifest.xml",
    "certs": "office-addin-dev-certs install",
    "manifest:dev": "bash scripts/make-dev-manifest.sh",
    "sideload:mac": "bash scripts/sideload-mac.sh manifest.dev.xml",
    "icons": "python3 scripts/make-icons.py"
  },
  "devDependencies": {
    "@types/office-js": "^1.0.612",
    "office-addin-dev-certs": "^2.0.0",
    "office-addin-manifest": "^3.0.1",
    "typescript": "^5.6.0",
    "vite": "^8.3.4",
    "vitest": "^5.0.4"
  }
}
```

- [ ] **Step 2: Crear `tsconfig.json`**

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "module": "ESNext",
    "moduleResolution": "bundler",
    "lib": ["ES2020", "DOM", "DOM.Iterable"],
    "types": ["office-js", "vite/client"],
    "strict": true,
    "noUncheckedIndexedAccess": true,
    "noImplicitOverride": true,
    "skipLibCheck": true,
    "isolatedModules": true,
    "noEmit": true
  },
  "include": ["src", "tests", "vite.config.ts", "vitest.config.ts"]
}
```

- [ ] **Step 3: Crear `vite.config.ts`**

```ts
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
```

- [ ] **Step 4: Crear `vitest.config.ts`**

```ts
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["tests/**/*.test.ts"],
    environment: "node",
    coverage: { provider: "v8", include: ["src/core/**"], thresholds: { lines: 90 } },
  },
});
```

- [ ] **Step 5: Crear `src/taskpane/index.html`**

```html
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Slide Tools</title>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <link rel="stylesheet" href="./styles.css" />
</head>
<body>
  <header class="st-header">
    <span class="st-title">Slide Tools</span>
    <span class="st-author">Victor Lopez</span>
  </header>
  <div id="notice" class="st-notice" hidden></div>
  <main id="panel"></main>
  <script type="module" src="./main.ts"></script>
</body>
</html>
```

- [ ] **Step 6: Crear `src/taskpane/styles.css`**

```css
:root {
  --st-bg: #ffffff;
  --st-fg: #1f1f1f;
  --st-muted: #6b6b6b;
  --st-border: #e1e1e1;
  --st-accent: #1f6feb;
  --st-hover: #f0f4ff;
  --st-danger: #b42318;
  font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  font-size: 12px;
  color: var(--st-fg);
}
@media (prefers-color-scheme: dark) {
  :root { --st-bg: #1e1e1e; --st-fg: #eaeaea; --st-muted: #a0a0a0; --st-border: #3a3a3a; --st-hover: #26324a; }
}
body { margin: 0; background: var(--st-bg); }
.st-header { display: flex; justify-content: space-between; align-items: baseline; padding: 8px 10px 4px; }
.st-title { font-weight: 600; font-size: 14px; }
.st-author { color: var(--st-muted); font-size: 11px; }
.st-notice { margin: 4px 10px; padding: 6px 8px; border-radius: 4px; background: var(--st-hover); border-left: 3px solid var(--st-accent); }
.st-notice.is-error { border-left-color: var(--st-danger); }
.st-section { padding: 6px 10px 2px; }
.st-section-title { color: var(--st-muted); font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; margin: 4px 0; }
.st-row { display: flex; flex-wrap: wrap; gap: 2px; }
.st-btn {
  width: 28px; height: 28px; display: inline-flex; align-items: center; justify-content: center;
  border: 1px solid transparent; border-radius: 4px; background: transparent; color: var(--st-fg); cursor: pointer; padding: 0;
}
.st-btn:hover:not(:disabled) { background: var(--st-hover); border-color: var(--st-border); }
.st-btn:disabled { opacity: 0.35; cursor: default; }
.st-btn svg { width: 18px; height: 18px; }
.st-divider { border: 0; border-top: 1px solid var(--st-border); margin: 6px 10px; }
.st-field { display: inline-flex; align-items: center; gap: 4px; margin-right: 8px; color: var(--st-muted); }
.st-field input { width: 36px; font: inherit; padding: 2px 4px; border: 1px solid var(--st-border); border-radius: 3px; background: var(--st-bg); color: var(--st-fg); }
```

- [ ] **Step 7: Crear `src/taskpane/main.ts` mínimo**

```ts
Office.onReady((info) => {
  const panel = document.getElementById("panel");
  if (!panel) return;
  panel.textContent =
    info.host === Office.HostType.PowerPoint ? "Listo en PowerPoint." : "Este panel requiere PowerPoint.";
});
```

- [ ] **Step 8: Crear `tests/smoke.test.ts`**

```ts
import { describe, expect, it } from "vitest";

describe("toolchain", () => {
  it("runs", () => {
    expect(1 + 1).toBe(2);
  });
});
```

- [ ] **Step 9: Instalar y verificar**

Run: `npm install && npm test && npm run build`
Expected: `npm test` → `1 passed`; `npm run build` → `dist/index.html` y `dist/assets/*.js` creados sin errores de `tsc`.

- [ ] **Step 10: Commit**

```bash
git add package.json package-lock.json tsconfig.json vite.config.ts vitest.config.ts src tests
git commit -m "build: toolchain Vite + Vitest + TypeScript y panel mínimo"
```

---

### Task 2: Iconos PNG del manifest

**Files:**
- Create: `scripts/make-icons.py`
- Create: `public/assets/icon-16.png`, `icon-32.png`, `icon-64.png`, `icon-80.png` (generados)

- [ ] **Step 1: Crear `scripts/make-icons.py`**

```python
"""Genera iconos PNG sólidos para el manifest usando solo la stdlib."""
import pathlib
import struct
import zlib

COLOR = (0x1F, 0x6F, 0xEB)  # azul acento del panel
SIZES = (16, 32, 64, 80)


def chunk(tag: bytes, data: bytes) -> bytes:
    return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)


def png(size: int, rgb: tuple[int, int, int]) -> bytes:
    row = b"\x00" + bytes(rgb) * size
    raw = row * size
    ihdr = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def main() -> None:
    out = pathlib.Path(__file__).resolve().parent.parent / "public" / "assets"
    out.mkdir(parents=True, exist_ok=True)
    for size in SIZES:
        (out / f"icon-{size}.png").write_bytes(png(size, COLOR))
        print(f"icon-{size}.png")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Generar y verificar**

Run: `npm run icons && file public/assets/icon-32.png`
Expected: `PNG image data, 32 x 32, 8-bit/color RGB`.

- [ ] **Step 3: Commit**

```bash
git add scripts/make-icons.py public/assets
git commit -m "build: iconos PNG del manifest generados con stdlib"
```

---

### Task 3: Manifest de producción, manifest de desarrollo y sideload Mac

**Files:**
- Create: `manifest.xml`
- Create: `scripts/make-dev-manifest.sh`
- Create: `scripts/sideload-mac.sh`
- Modify: `.gitignore` (añadir `manifest.dev.xml`)

- [ ] **Step 1: Crear `manifest.xml`**

```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>7b2c1f3e-5a84-4d6b-9c1e-2f8a4b6d9e01</Id>
  <Version>0.1.0.0</Version>
  <ProviderName>Victor Lopez</ProviderName>
  <DefaultLocale>es-ES</DefaultLocale>
  <DisplayName DefaultValue="Slide Tools"/>
  <Description DefaultValue="Alineación, tamaño, formato y texto para PowerPoint desde un panel lateral."/>
  <IconUrl DefaultValue="https://xictorlr.github.io/skillsEEpowerPoint/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://xictorlr.github.io/skillsEEpowerPoint/assets/icon-80.png"/>
  <SupportUrl DefaultValue="https://github.com/xictorlr/skillsEEpowerPoint"/>
  <AppDomains>
    <AppDomain>https://xictorlr.github.io</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.8"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://xictorlr.github.io/skillsEEpowerPoint/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="SlideTools.Group">
                <Label resid="Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="SlideTools.OpenPane">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="Button.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>SlideTools.Pane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://xictorlr.github.io/skillsEEpowerPoint/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://xictorlr.github.io/skillsEEpowerPoint/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://xictorlr.github.io/skillsEEpowerPoint/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://github.com/xictorlr/skillsEEpowerPoint"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://xictorlr.github.io/skillsEEpowerPoint/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Slide Tools instalado"/>
        <bt:String id="Group.Label" DefaultValue="Slide Tools"/>
        <bt:String id="Button.Label" DefaultValue="Abrir panel"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Abre el panel desde la pestaña Inicio para alinear, dimensionar y formatear formas."/>
        <bt:String id="Button.Tooltip" DefaultValue="Muestra el panel lateral de Slide Tools."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

- [ ] **Step 2: Validar el manifest**

Run: `npm run validate`
Expected: `The manifest is valid.`

- [ ] **Step 3: Crear `scripts/make-dev-manifest.sh`**

```bash
#!/usr/bin/env bash
# Genera manifest.dev.xml apuntando a https://localhost:3000 con un Id distinto.
set -euo pipefail
cd "$(dirname "$0")/.."
sed \
  -e 's#https://xictorlr.github.io/skillsEEpowerPoint/#https://localhost:3000/#g' \
  -e 's#<AppDomain>https://xictorlr.github.io</AppDomain>#<AppDomain>https://localhost:3000</AppDomain>#' \
  -e 's#<Id>7b2c1f3e-5a84-4d6b-9c1e-2f8a4b6d9e01</Id>#<Id>7b2c1f3e-5a84-4d6b-9c1e-2f8a4b6d9e02</Id>#' \
  -e 's#DisplayName DefaultValue="Slide Tools"#DisplayName DefaultValue="Slide Tools (dev)"#' \
  manifest.xml > manifest.dev.xml
echo "manifest.dev.xml generado"
```

- [ ] **Step 4: Crear `scripts/sideload-mac.sh`**

```bash
#!/usr/bin/env bash
# Copia un manifest a la carpeta de add-ins sideloaded de PowerPoint para Mac.
set -euo pipefail
MANIFEST="${1:-manifest.xml}"
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF"
cp "$MANIFEST" "$WEF/slide-tools-$(basename "$MANIFEST")"
echo "Copiado a: $WEF"
echo "Reinicia PowerPoint y abre Insertar > Complementos > Mis complementos > Complementos para desarrolladores."
```

- [ ] **Step 5: Generar el manifest de desarrollo y validarlo**

Run: `chmod +x scripts/*.sh && npm run manifest:dev && npx office-addin-manifest validate manifest.dev.xml && printf 'manifest.dev.xml\n' >> .gitignore`
Expected: ambas validaciones `The manifest is valid.`

- [ ] **Step 6: Commit**

```bash
git add manifest.xml scripts/make-dev-manifest.sh scripts/sideload-mac.sh .gitignore
git commit -m "feat: manifest del add-in, generador de manifest dev y sideload Mac"
```

---

### Task 4: Tipos del dominio y `alignShapes`

**Files:**
- Create: `src/core/types.ts`
- Create: `src/core/geometry.ts`
- Create: `tests/core/fixtures.ts`
- Create: `tests/core/align.test.ts`

- [ ] **Step 1: Crear `src/core/types.ts`**

```ts
export interface Rect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface ShapeModel extends Rect {
  id: string;
  name: string;
  type: string;
  rotation: number;
  zOrder: number;
  isGroupChild: boolean;
}

export interface SlideSize {
  width: number;
  height: number;
}

export interface SelectionContext {
  /** En el orden devuelto por getSelectedShapes(): el primero es la referencia. */
  shapes: ShapeModel[];
  slide: SlideSize;
}

/** Cambio parcial de geometría para una forma. Solo se incluyen los campos modificados. */
export type Patch = { id: string } & Partial<Rect>;
```

- [ ] **Step 2: Crear `tests/core/fixtures.ts`**

```ts
import type { ShapeModel, SlideSize } from "../../src/core/types";

let counter = 0;

export function shape(partial: Partial<ShapeModel> & Pick<ShapeModel, "left" | "top" | "width" | "height">): ShapeModel {
  counter += 1;
  return {
    id: partial.id ?? `s${counter}`,
    name: partial.name ?? `Shape ${counter}`,
    type: partial.type ?? "GeometricShape",
    rotation: partial.rotation ?? 0,
    zOrder: partial.zOrder ?? counter,
    isGroupChild: partial.isGroupChild ?? false,
    left: partial.left,
    top: partial.top,
    width: partial.width,
    height: partial.height,
  };
}

export const slide16x9: SlideSize = { width: 960, height: 540 };
```

- [ ] **Step 3: Escribir el test que falla `tests/core/align.test.ts`**

```ts
import { describe, expect, it } from "vitest";
import { alignShapes } from "../../src/core/geometry";
import { shape } from "./fixtures";

describe("alignShapes", () => {
  const ref = shape({ id: "ref", left: 100, top: 100, width: 200, height: 100 });
  const a = shape({ id: "a", left: 400, top: 300, width: 50, height: 40 });
  const b = shape({ id: "b", left: 10, top: 20, width: 80, height: 60 });

  it("returns [] with fewer than 2 shapes", () => {
    expect(alignShapes([ref], "left")).toEqual([]);
  });

  it("aligns left edges to the reference (first shape)", () => {
    expect(alignShapes([ref, a, b], "left")).toEqual([
      { id: "a", left: 100 },
      { id: "b", left: 100 },
    ]);
  });

  it("aligns right edges", () => {
    expect(alignShapes([ref, a, b], "right")).toEqual([
      { id: "a", left: 250 },
      { id: "b", left: 220 },
    ]);
  });

  it("aligns horizontal centers", () => {
    expect(alignShapes([ref, a, b], "centerH")).toEqual([
      { id: "a", left: 175 },
      { id: "b", left: 160 },
    ]);
  });

  it("aligns top, bottom and vertical centers", () => {
    expect(alignShapes([ref, a], "top")).toEqual([{ id: "a", top: 100 }]);
    expect(alignShapes([ref, a], "bottom")).toEqual([{ id: "a", top: 160 }]);
    expect(alignShapes([ref, a], "centerV")).toEqual([{ id: "a", top: 130 }]);
  });

  it("omits shapes already aligned", () => {
    const same = shape({ id: "same", left: 100, top: 0, width: 10, height: 10 });
    expect(alignShapes([ref, same], "left")).toEqual([]);
  });
});
```

- [ ] **Step 4: Ejecutar y ver que falla**

Run: `npx vitest run tests/core/align.test.ts`
Expected: FAIL, `Failed to resolve import "../../src/core/geometry"`.

- [ ] **Step 5: Implementar `src/core/geometry.ts` (solo alignShapes y helpers)**

```ts
import type { Patch, Rect, ShapeModel, SlideSize } from "./types";

export type AlignEdge = "left" | "centerH" | "right" | "top" | "centerV" | "bottom";

const EPS = 0.01;

function near(a: number, b: number): boolean {
  return Math.abs(a - b) < EPS;
}

function targetLeft(shape: Rect, ref: Rect, edge: "left" | "centerH" | "right"): number {
  if (edge === "left") return ref.left;
  if (edge === "right") return ref.left + ref.width - shape.width;
  return ref.left + ref.width / 2 - shape.width / 2;
}

function targetTop(shape: Rect, ref: Rect, edge: "top" | "centerV" | "bottom"): number {
  if (edge === "top") return ref.top;
  if (edge === "bottom") return ref.top + ref.height - shape.height;
  return ref.top + ref.height / 2 - shape.height / 2;
}

/** Alinea cada forma al rectángulo `ref` por el borde indicado. */
function alignToRect(shapes: ShapeModel[], ref: Rect, edge: AlignEdge): Patch[] {
  const patches: Patch[] = [];
  for (const s of shapes) {
    if (edge === "left" || edge === "centerH" || edge === "right") {
      const left = targetLeft(s, ref, edge);
      if (!near(left, s.left)) patches.push({ id: s.id, left });
    } else {
      const top = targetTop(s, ref, edge);
      if (!near(top, s.top)) patches.push({ id: s.id, top });
    }
  }
  return patches;
}

/** Alinea las formas 2..n a la primera (referencia). */
export function alignShapes(shapes: ShapeModel[], edge: AlignEdge): Patch[] {
  if (shapes.length < 2) return [];
  const [ref, ...rest] = shapes;
  if (!ref) return [];
  return alignToRect(rest, ref, edge);
}

/** Alinea todas las formas a los bordes/centro de la diapositiva. */
export function alignToSlide(shapes: ShapeModel[], slide: SlideSize, edge: AlignEdge): Patch[] {
  if (shapes.length === 0) return [];
  return alignToRect(shapes, { left: 0, top: 0, width: slide.width, height: slide.height }, edge);
}
```

- [ ] **Step 6: Ejecutar y ver que pasa**

Run: `npx vitest run tests/core/align.test.ts`
Expected: `6 passed`.

- [ ] **Step 7: Añadir tests de `alignToSlide` al mismo archivo**

```ts
import { alignToSlide } from "../../src/core/geometry";
import { slide16x9 } from "./fixtures";

describe("alignToSlide", () => {
  const a = shape({ id: "a", left: 10, top: 10, width: 100, height: 50 });

  it("centers horizontally and vertically on the slide", () => {
    expect(alignToSlide([a], slide16x9, "centerH")).toEqual([{ id: "a", left: 430 }]);
    expect(alignToSlide([a], slide16x9, "centerV")).toEqual([{ id: "a", top: 245 }]);
  });

  it("aligns to slide edges", () => {
    expect(alignToSlide([a], slide16x9, "right")).toEqual([{ id: "a", left: 860 }]);
    expect(alignToSlide([a], slide16x9, "bottom")).toEqual([{ id: "a", top: 490 }]);
  });

  it("works with a single shape and returns [] for none", () => {
    expect(alignToSlide([], slide16x9, "left")).toEqual([]);
  });
});
```

Run: `npx vitest run tests/core/align.test.ts`
Expected: `9 passed`.

- [ ] **Step 8: Commit**

```bash
git add src/core/types.ts src/core/geometry.ts tests/core/fixtures.ts tests/core/align.test.ts
git commit -m "feat(core): tipos del dominio y alineación a referencia o diapositiva"
```

---

### Task 5: `distribute` y `swapPositions`

**Files:**
- Modify: `src/core/geometry.ts`
- Create: `tests/core/distribute.test.ts`

- [ ] **Step 1: Escribir el test que falla**

```ts
import { describe, expect, it } from "vitest";
import { distribute, swapPositions } from "../../src/core/geometry";
import { shape } from "./fixtures";

describe("distribute", () => {
  it("returns [] with fewer than 3 shapes", () => {
    const a = shape({ left: 0, top: 0, width: 10, height: 10 });
    const b = shape({ left: 100, top: 0, width: 10, height: 10 });
    expect(distribute([a, b], "h")).toEqual([]);
  });

  it("spaces shapes evenly between the outermost ones (horizontal)", () => {
    const a = shape({ id: "a", left: 0, top: 0, width: 100, height: 10 });
    const b = shape({ id: "b", left: 120, top: 0, width: 50, height: 10 });
    const c = shape({ id: "c", left: 400, top: 0, width: 100, height: 10 });
    // span 0..500, widths 250, gaps = (500-250)/2 = 125 → b.left = 225
    expect(distribute([c, a, b], "h")).toEqual([{ id: "b", left: 225 }]);
  });

  it("spaces shapes evenly (vertical) regardless of input order", () => {
    const a = shape({ id: "a", left: 0, top: 0, width: 10, height: 40 });
    const b = shape({ id: "b", left: 0, top: 50, width: 10, height: 20 });
    const c = shape({ id: "c", left: 0, top: 200, width: 10, height: 40 });
    const d = shape({ id: "d", left: 0, top: 500, width: 10, height: 100 });
    // span 0..600, heights 200, gaps = 400/3 ≈ 133.33
    const result = distribute([d, b, c, a], "v");
    expect(result).toHaveLength(2);
    expect(result[0]?.id).toBe("b");
    expect(result[0]?.top).toBeCloseTo(173.333, 2);
    expect(result[1]?.id).toBe("c");
    expect(result[1]?.top).toBeCloseTo(326.667, 2);
  });
});

describe("swapPositions", () => {
  it("swaps the centers of exactly two shapes", () => {
    const a = shape({ id: "a", left: 0, top: 0, width: 100, height: 100 });   // center 50,50
    const b = shape({ id: "b", left: 300, top: 200, width: 20, height: 40 }); // center 310,220
    expect(swapPositions([a, b])).toEqual([
      { id: "a", left: 260, top: 170 },
      { id: "b", left: 40, top: 30 },
    ]);
  });

  it("returns [] unless exactly two shapes", () => {
    const a = shape({ left: 0, top: 0, width: 1, height: 1 });
    expect(swapPositions([a])).toEqual([]);
    expect(swapPositions([a, a, a])).toEqual([]);
  });
});
```

- [ ] **Step 2: Ejecutar y ver que falla**

Run: `npx vitest run tests/core/distribute.test.ts`
Expected: FAIL, `distribute is not a function` (o import sin exportar).

- [ ] **Step 3: Añadir a `src/core/geometry.ts`**

```ts
export type Axis = "h" | "v";

/** Reparte el espacio entre formas de modo uniforme. Las dos extremas (por posición) no se mueven. */
export function distribute(shapes: ShapeModel[], axis: Axis): Patch[] {
  if (shapes.length < 3) return [];
  const pos = axis === "h" ? "left" : "top";
  const size = axis === "h" ? "width" : "height";
  const sorted = [...shapes].sort((a, b) => a[pos] + a[size] / 2 - (b[pos] + b[size] / 2));
  const first = sorted[0];
  const last = sorted[sorted.length - 1];
  if (!first || !last) return [];
  const span = last[pos] + last[size] - first[pos];
  const total = sorted.reduce((acc, s) => acc + s[size], 0);
  const gap = (span - total) / (sorted.length - 1);
  const patches: Patch[] = [];
  let cursor = first[pos] + first[size] + gap;
  for (const s of sorted.slice(1, -1)) {
    if (!near(cursor, s[pos])) patches.push({ id: s.id, [pos]: cursor } as Patch);
    cursor += s[size] + gap;
  }
  return patches;
}

/** Intercambia los centros de dos formas. */
export function swapPositions(shapes: ShapeModel[]): Patch[] {
  if (shapes.length !== 2) return [];
  const [a, b] = shapes;
  if (!a || !b) return [];
  const ca = { x: a.left + a.width / 2, y: a.top + a.height / 2 };
  const cb = { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  return [
    { id: a.id, left: cb.x - a.width / 2, top: cb.y - a.height / 2 },
    { id: b.id, left: ca.x - b.width / 2, top: ca.y - b.height / 2 },
  ];
}
```

- [ ] **Step 4: Ejecutar y ver que pasa**

Run: `npx vitest run tests/core`
Expected: todos `passed` (9 + 5).

- [ ] **Step 5: Commit**

```bash
git add src/core/geometry.ts tests/core/distribute.test.ts
git commit -m "feat(core): distribuir uniformemente e intercambiar posiciones"
```

---

### Task 6: `matchSize` y `stretchToEdge`

**Files:**
- Modify: `src/core/geometry.ts`
- Create: `tests/core/size.test.ts`

- [ ] **Step 1: Escribir el test que falla**

```ts
import { describe, expect, it } from "vitest";
import { matchSize, stretchToEdge } from "../../src/core/geometry";
import { shape } from "./fixtures";

describe("matchSize", () => {
  const ref = shape({ id: "ref", left: 0, top: 0, width: 200, height: 100 });
  const a = shape({ id: "a", left: 300, top: 0, width: 50, height: 40 });
  const b = shape({ id: "b", left: 600, top: 0, width: 300, height: 150 });

  it("returns [] with fewer than 2 shapes", () => {
    expect(matchSize([ref], "width", "first")).toEqual([]);
  });

  it("matches width to the first shape, keeping top-left", () => {
    expect(matchSize([ref, a, b], "width", "first")).toEqual([
      { id: "a", width: 200 },
      { id: "b", width: 200 },
    ]);
  });

  it("matches height and both", () => {
    expect(matchSize([ref, a], "height", "first")).toEqual([{ id: "a", height: 100 }]);
    expect(matchSize([ref, a], "both", "first")).toEqual([{ id: "a", width: 200, height: 100 }]);
  });

  it("matches to the largest in the selection, including the first", () => {
    expect(matchSize([ref, a, b], "width", "max")).toEqual([
      { id: "ref", width: 300 },
      { id: "a", width: 300 },
    ]);
  });

  it("matches to the smallest in the selection", () => {
    expect(matchSize([ref, a, b], "height", "min")).toEqual([
      { id: "ref", height: 40 },
      { id: "b", height: 40 },
    ]);
  });
});

describe("stretchToEdge", () => {
  const ref = shape({ id: "ref", left: 100, top: 100, width: 400, height: 200 }); // right 500, bottom 300
  const a = shape({ id: "a", left: 200, top: 150, width: 50, height: 30 });

  it("stretches right edge to the reference right edge, keeping left fixed", () => {
    expect(stretchToEdge([ref, a], "right")).toEqual([{ id: "a", width: 300 }]);
  });

  it("stretches left edge to the reference left edge, keeping right fixed", () => {
    expect(stretchToEdge([ref, a], "left")).toEqual([{ id: "a", left: 100, width: 150 }]);
  });

  it("stretches top and bottom", () => {
    expect(stretchToEdge([ref, a], "top")).toEqual([{ id: "a", top: 100, height: 80 }]);
    expect(stretchToEdge([ref, a], "bottom")).toEqual([{ id: "a", height: 150 }]);
  });

  it("returns [] when the result would have non-positive size", () => {
    const outside = shape({ id: "o", left: 600, top: 0, width: 50, height: 30 }); // left beyond ref.right
    expect(stretchToEdge([ref, outside], "right")).toEqual([]);
  });
});
```

- [ ] **Step 2: Ejecutar y ver que falla**

Run: `npx vitest run tests/core/size.test.ts`
Expected: FAIL por exportaciones inexistentes.

- [ ] **Step 3: Añadir a `src/core/geometry.ts`**

```ts
export type SizeDim = "width" | "height" | "both";
export type SizeMode = "first" | "max" | "min";
export type Edge = "left" | "right" | "top" | "bottom";

function pickTarget(shapes: ShapeModel[], key: "width" | "height", mode: SizeMode): number {
  if (mode === "first") return shapes[0]?.[key] ?? 0;
  const values = shapes.map((s) => s[key]);
  return mode === "max" ? Math.max(...values) : Math.min(...values);
}

/** Iguala ancho/alto. mode=first: al primero (no se toca). max/min: al mayor/menor de toda la selección. */
export function matchSize(shapes: ShapeModel[], dim: SizeDim, mode: SizeMode): Patch[] {
  if (shapes.length < 2) return [];
  const targets = shapes.slice(mode === "first" ? 1 : 0);
  const w = dim !== "height" ? pickTarget(shapes, "width", mode) : undefined;
  const h = dim !== "width" ? pickTarget(shapes, "height", mode) : undefined;
  const patches: Patch[] = [];
  for (const s of targets) {
    const patch: Patch = { id: s.id };
    if (w !== undefined && !near(w, s.width)) patch.width = w;
    if (h !== undefined && !near(h, s.height)) patch.height = h;
    if (patch.width !== undefined || patch.height !== undefined) patches.push(patch);
  }
  return patches;
}

/** Estira el borde indicado de cada forma hasta el mismo borde de la referencia; el borde opuesto no se mueve. */
export function stretchToEdge(shapes: ShapeModel[], edge: Edge): Patch[] {
  if (shapes.length < 2) return [];
  const [ref, ...rest] = shapes;
  if (!ref) return [];
  const refRight = ref.left + ref.width;
  const refBottom = ref.top + ref.height;
  const patches: Patch[] = [];
  for (const s of rest) {
    let patch: Patch | undefined;
    if (edge === "right") {
      const width = refRight - s.left;
      if (width > 0 && !near(width, s.width)) patch = { id: s.id, width };
    } else if (edge === "left") {
      const width = s.left + s.width - ref.left;
      if (width > 0 && !near(ref.left, s.left)) patch = { id: s.id, left: ref.left, width };
    } else if (edge === "bottom") {
      const height = refBottom - s.top;
      if (height > 0 && !near(height, s.height)) patch = { id: s.id, height };
    } else {
      const height = s.top + s.height - ref.top;
      if (height > 0 && !near(ref.top, s.top)) patch = { id: s.id, top: ref.top, height };
    }
    if (patch) patches.push(patch);
  }
  return patches;
}
```

- [ ] **Step 4: Ejecutar y ver que pasa**

Run: `npx vitest run tests/core`
Expected: todos `passed`.

- [ ] **Step 5: Commit**

```bash
git add src/core/geometry.ts tests/core/size.test.ts
git commit -m "feat(core): igualar tamaños y estirar hasta borde de referencia"
```

---

### Task 7: `arrangeMatrix`

**Files:**
- Modify: `src/core/geometry.ts`
- Create: `tests/core/matrix.test.ts`

- [ ] **Step 1: Escribir el test que falla**

```ts
import { describe, expect, it } from "vitest";
import { arrangeMatrix } from "../../src/core/geometry";
import { shape } from "./fixtures";

describe("arrangeMatrix", () => {
  it("returns [] with fewer than 2 shapes or invalid columns", () => {
    const a = shape({ left: 0, top: 0, width: 10, height: 10 });
    const b = shape({ left: 50, top: 0, width: 10, height: 10 });
    expect(arrangeMatrix([a], { columns: 2, gapX: 5, gapY: 5 })).toEqual([]);
    expect(arrangeMatrix([a, b], { columns: 0, gapX: 5, gapY: 5 })).toEqual([]);
  });

  it("lays shapes out in reading order from the selection's top-left, with equal column/row sizes", () => {
    // Orden de lectura (por top, luego left): c (0,0), a (0,100), b (60,20), d (60,200)
    const a = shape({ id: "a", left: 100, top: 0, width: 40, height: 20 });
    const b = shape({ id: "b", left: 20, top: 60, width: 30, height: 30 });
    const c = shape({ id: "c", left: 0, top: 0, width: 20, height: 10 });
    const d = shape({ id: "d", left: 200, top: 60, width: 10, height: 10 });
    // columna = ancho máximo 40, fila = alto máximo 30, gap 10 → cols x: 0, 50 ; rows y: 0, 40
    expect(arrangeMatrix([a, b, c, d], { columns: 2, gapX: 10, gapY: 10 })).toEqual([
      { id: "a", left: 50, top: 0 },
      { id: "b", left: 0, top: 40 },
      { id: "d", left: 50, top: 40 },
    ]);
  });
});
```

- [ ] **Step 2: Ejecutar y ver que falla**

Run: `npx vitest run tests/core/matrix.test.ts`
Expected: FAIL.

- [ ] **Step 3: Añadir a `src/core/geometry.ts`**

```ts
export interface MatrixOptions {
  columns: number;
  gapX: number;
  gapY: number;
}

/**
 * Coloca las formas en una rejilla de `columns` columnas, en orden de lectura (top, luego left),
 * partiendo de la esquina superior izquierda de la selección. Cada celda mide el máximo ancho/alto.
 */
export function arrangeMatrix(shapes: ShapeModel[], opts: MatrixOptions): Patch[] {
  if (shapes.length < 2 || !Number.isInteger(opts.columns) || opts.columns < 1) return [];
  const ordered = [...shapes].sort((a, b) => a.top - b.top || a.left - b.left);
  const originX = Math.min(...shapes.map((s) => s.left));
  const originY = Math.min(...shapes.map((s) => s.top));
  const cellW = Math.max(...shapes.map((s) => s.width));
  const cellH = Math.max(...shapes.map((s) => s.height));
  const patches: Patch[] = [];
  ordered.forEach((s, i) => {
    const col = i % opts.columns;
    const row = Math.floor(i / opts.columns);
    const left = originX + col * (cellW + opts.gapX);
    const top = originY + row * (cellH + opts.gapY);
    if (!near(left, s.left) || !near(top, s.top)) patches.push({ id: s.id, left, top });
  });
  return patches;
}
```

- [ ] **Step 4: Ejecutar y ver que pasa**

Run: `npx vitest run tests/core --coverage`
Expected: todos `passed`; cobertura de `src/core/geometry.ts` ≥ 90 % líneas.

- [ ] **Step 5: Commit**

```bash
git add src/core/geometry.ts tests/core/matrix.test.ts
git commit -m "feat(core): alinear en matriz"
```

---

### Task 8: Tamaño de diapositiva desde PNG (`core/png.ts`)

Office.js no expone el tamaño de la diapositiva. Se renderiza la diapositiva actual a PNG pequeño con `Slide.getImageAsBase64` y se deduce la proporción; el alto estándar de PowerPoint es 540 pt para 16:9, 4:3, A4 y Letter.

**Files:**
- Create: `src/core/png.ts`
- Create: `tests/core/png.test.ts`

- [ ] **Step 1: Escribir el test que falla**

```ts
import { describe, expect, it } from "vitest";
import { pngDimensions, slideSizeFromAspect } from "../../src/core/png";

// PNG 1x1 válido (cabecera estándar)
const PNG_1x1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

describe("pngDimensions", () => {
  it("reads width and height from the IHDR chunk", () => {
    expect(pngDimensions(PNG_1x1)).toEqual({ width: 1, height: 1 });
  });

  it("throws on non-PNG data", () => {
    expect(() => pngDimensions(btoa("not a png at all, really"))).toThrow(/PNG/);
  });
});

describe("slideSizeFromAspect", () => {
  it("assumes 540pt height and scales width by aspect ratio", () => {
    expect(slideSizeFromAspect({ width: 160, height: 90 })).toEqual({ width: 960, height: 540 });
    expect(slideSizeFromAspect({ width: 40, height: 30 })).toEqual({ width: 720, height: 540 });
  });
});
```

- [ ] **Step 2: Ejecutar y ver que falla**

Run: `npx vitest run tests/core/png.test.ts`
Expected: FAIL.

- [ ] **Step 3: Implementar `src/core/png.ts`**

```ts
import type { SlideSize } from "./types";

export interface PixelSize {
  width: number;
  height: number;
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
export const STANDARD_SLIDE_HEIGHT_PT = 540;

/** Lee ancho/alto del chunk IHDR de un PNG en base64. */
export function pngDimensions(base64: string): PixelSize {
  const bytes = Uint8Array.from(atob(base64.slice(0, 64)), (c) => c.charCodeAt(0));
  const isPng = PNG_SIGNATURE.every((b, i) => bytes[i] === b);
  if (!isPng || bytes.length < 24) throw new Error("Los datos no son un PNG");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return { width: view.getUint32(16), height: view.getUint32(20) };
}

/** Convierte la proporción de una imagen de la diapositiva a puntos, asumiendo alto estándar 540 pt. */
export function slideSizeFromAspect(px: PixelSize): SlideSize {
  const width = Math.round((STANDARD_SLIDE_HEIGHT_PT * px.width) / px.height);
  return { width, height: STANDARD_SLIDE_HEIGHT_PT };
}
```

- [ ] **Step 4: Ejecutar y ver que pasa**

Run: `npx vitest run tests/core/png.test.ts`
Expected: `3 passed`.

- [ ] **Step 5: Commit**

```bash
git add src/core/png.ts tests/core/png.test.ts
git commit -m "feat(core): deducir tamaño de diapositiva desde un PNG renderizado"
```

---

### Task 9: Adaptador Office: runtime, selección, tamaño de diapositiva y aplicar parches

**Files:**
- Create: `src/office/runtime.ts`
- Create: `src/office/selection.ts`
- Create: `src/office/slideSize.ts`
- Create: `src/office/apply.ts`
- Create: `tests/office/mockPowerPoint.ts`
- Create: `tests/office/apply.test.ts`
- Create: `tests/office/selection.test.ts`

- [ ] **Step 1: Crear el mock `tests/office/mockPowerPoint.ts`**

```ts
/**
 * Mock mínimo de los proxies de PowerPoint que usan selection.ts y apply.ts.
 * Registra escrituras en `writes` y cuenta llamadas a sync.
 */
export interface MockShapeInit {
  id: string;
  name?: string;
  type?: string;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
  zOrderPosition?: number;
  level?: number;
}

export class MockShape {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation: number;
  zOrderPosition: number;
  level: number;
  writes: Array<{ prop: string; value: number }> = [];

  constructor(init: MockShapeInit) {
    this.id = init.id;
    this.name = init.name ?? init.id;
    this.type = init.type ?? "GeometricShape";
    this.left = init.left;
    this.top = init.top;
    this.width = init.width;
    this.height = init.height;
    this.rotation = init.rotation ?? 0;
    this.zOrderPosition = init.zOrderPosition ?? 0;
    this.level = init.level ?? 0;
    for (const prop of ["left", "top", "width", "height"] as const) {
      let value = this[prop];
      Object.defineProperty(this, prop, {
        get: () => value,
        set: (v: number) => {
          value = v;
          this.writes.push({ prop, value: v });
        },
      });
    }
  }
}

export class MockContext {
  syncCount = 0;
  constructor(public shapes: MockShape[], public slidePng = "") {}
  presentation = {
    getSelectedShapes: () => ({ items: this.shapes, load: () => undefined }),
    getSelectedSlides: () => ({
      getItemAt: () => ({ getImageAsBase64: () => ({ value: this.slidePng }) }),
    }),
  };
  async sync(): Promise<void> {
    this.syncCount += 1;
  }
}
```

- [ ] **Step 2: Escribir `tests/office/apply.test.ts` (falla)**

```ts
import { describe, expect, it } from "vitest";
import { applyPatches } from "../../src/office/apply";
import { MockShape } from "./mockPowerPoint";

describe("applyPatches", () => {
  it("writes only the fields present in each patch", () => {
    const a = new MockShape({ id: "a", left: 0, top: 0, width: 10, height: 10 });
    const b = new MockShape({ id: "b", left: 0, top: 0, width: 10, height: 10 });
    const proxies = new Map<string, MockShape>([["a", a], ["b", b]]);
    applyPatches(proxies, [{ id: "a", left: 5 }, { id: "b", top: 7, height: 20 }]);
    expect(a.writes).toEqual([{ prop: "left", value: 5 }]);
    expect(b.writes).toEqual([{ prop: "top", value: 7 }, { prop: "height", value: 20 }]);
  });

  it("ignores patches whose id is not in the map", () => {
    const a = new MockShape({ id: "a", left: 0, top: 0, width: 10, height: 10 });
    applyPatches(new Map([["a", a]]), [{ id: "zzz", left: 1 }]);
    expect(a.writes).toEqual([]);
  });
});
```

- [ ] **Step 3: Escribir `tests/office/selection.test.ts` (falla)**

```ts
import { describe, expect, it } from "vitest";
import { readSelection } from "../../src/office/selection";
import { MockContext, MockShape } from "./mockPowerPoint";

const PNG_1x1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

describe("readSelection", () => {
  it("maps proxies to ShapeModel and keeps a proxy map by id", async () => {
    const ctx = new MockContext(
      [new MockShape({ id: "a", left: 1, top: 2, width: 3, height: 4, rotation: 10, zOrderPosition: 2, level: 1 })],
      PNG_1x1,
    );
    const { selection, proxies } = await readSelection(ctx as never, { slideSizeCache: new Map() });
    expect(selection.shapes).toEqual([
      { id: "a", name: "a", type: "GeometricShape", left: 1, top: 2, width: 3, height: 4, rotation: 10, zOrder: 2, isGroupChild: true },
    ]);
    expect(proxies.get("a")).toBeDefined();
    expect(selection.slide).toEqual({ width: 540, height: 540 }); // PNG 1x1 → aspecto 1:1
  });

  it("falls back to 960x540 when the slide image cannot be read", async () => {
    const ctx = new MockContext([], "not-base64-png");
    const { selection } = await readSelection(ctx as never, { slideSizeCache: new Map() });
    expect(selection.shapes).toEqual([]);
    expect(selection.slide).toEqual({ width: 960, height: 540 });
  });
});
```

- [ ] **Step 4: Ejecutar y ver que fallan**

Run: `npx vitest run tests/office`
Expected: FAIL por módulos inexistentes.

- [ ] **Step 5: Crear `src/office/apply.ts`**

```ts
import type { Patch } from "../core/types";

/** Subconjunto de PowerPoint.Shape que escribimos. Permite mocks en tests. */
export interface WritableShape {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** Escribe en los proxies solo los campos presentes en cada parche. No hace sync. */
export function applyPatches(proxies: Map<string, WritableShape>, patches: Patch[]): void {
  for (const p of patches) {
    const shape = proxies.get(p.id);
    if (!shape) continue;
    if (p.left !== undefined) shape.left = p.left;
    if (p.top !== undefined) shape.top = p.top;
    if (p.width !== undefined) shape.width = p.width;
    if (p.height !== undefined) shape.height = p.height;
  }
}
```

- [ ] **Step 6: Crear `src/office/slideSize.ts`**

```ts
import { pngDimensions, slideSizeFromAspect } from "../core/png";
import type { SlideSize } from "../core/types";

export const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };
const CACHE_KEY = "current";

/**
 * Deduce el tamaño de la diapositiva seleccionada renderizándola a 54 px de alto.
 * Cachea el resultado (el tamaño no cambia dentro de una presentación). Fallback: 16:9.
 */
export async function getSlideSize(
  context: PowerPoint.RequestContext,
  cache: Map<string, SlideSize>,
): Promise<SlideSize> {
  const cached = cache.get(CACHE_KEY);
  if (cached) return cached;
  try {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const image = slide.getImageAsBase64({ height: 54 });
    await context.sync();
    const size = slideSizeFromAspect(pngDimensions(image.value));
    cache.set(CACHE_KEY, size);
    return size;
  } catch {
    return DEFAULT_SLIDE_SIZE;
  }
}
```

- [ ] **Step 7: Crear `src/office/selection.ts`**

```ts
import type { SelectionContext, ShapeModel, SlideSize } from "../core/types";
import { getSlideSize } from "./slideSize";

export interface ReadSelectionDeps {
  slideSizeCache: Map<string, SlideSize>;
}

export interface SelectionRead {
  selection: SelectionContext;
  proxies: Map<string, PowerPoint.Shape>;
}

const SHAPE_PROPS = "items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/rotation,items/zOrderPosition,items/level";

/** Carga las formas seleccionadas y el tamaño de diapositiva. Hace un sync. */
export async function readSelection(context: PowerPoint.RequestContext, deps: ReadSelectionDeps): Promise<SelectionRead> {
  const collection = context.presentation.getSelectedShapes();
  collection.load(SHAPE_PROPS);
  await context.sync();

  const proxies = new Map<string, PowerPoint.Shape>();
  const shapes: ShapeModel[] = collection.items.map((s) => {
    proxies.set(s.id, s);
    return {
      id: s.id,
      name: s.name,
      type: String(s.type),
      left: s.left,
      top: s.top,
      width: s.width,
      height: s.height,
      rotation: s.rotation,
      zOrder: s.zOrderPosition,
      isGroupChild: (s.level ?? 0) > 0,
    };
  });

  const slide = await getSlideSize(context, deps.slideSizeCache);
  return { selection: { shapes, slide }, proxies };
}
```

- [ ] **Step 8: Crear `src/office/runtime.ts`**

```ts
export type CommandResult = { ok: true; changed: number } | { ok: false; message: string };

/** Mensajes de error legibles por código de OfficeExtension. */
function describeError(err: unknown): string {
  if (err instanceof OfficeExtension.Error) {
    if (err.code === "ApiNotFound") return "Esta función requiere una versión más reciente de PowerPoint.";
    if (err.code === "InvalidOperation") return "Operación no válida para la selección actual.";
    return `${err.code}: ${err.message}`;
  }
  return err instanceof Error ? err.message : String(err);
}

/** Ejecuta `fn` dentro de PowerPoint.run y normaliza el resultado. `fn` devuelve cuántas formas cambió. */
export async function runCommand(
  commandId: string,
  fn: (context: PowerPoint.RequestContext) => Promise<number>,
): Promise<CommandResult> {
  try {
    const changed = await PowerPoint.run(fn);
    return { ok: true, changed };
  } catch (err) {
    const message = describeError(err);
    console.error(`[slide-tools] ${commandId} falló:`, err);
    return { ok: false, message };
  }
}
```

- [ ] **Step 9: Ejecutar y ver que pasan**

Run: `npx vitest run tests/office && npm run typecheck`
Expected: `4 passed`; typecheck sin errores. Si `items/level` provoca error de tipos, confirmar que `@types/office-js` ≥ 1.0.400 (incluye `Shape.level`).

- [ ] **Step 10: Commit**

```bash
git add src/office tests/office
git commit -m "feat(office): leer selección, tamaño de diapositiva, aplicar parches y wrapper de ejecución"
```

---

### Task 10: Comandos declarativos y ejecutor

**Files:**
- Create: `src/commands/types.ts`
- Create: `src/commands/alignment.ts`
- Create: `src/commands/size.ts`
- Create: `src/commands/registry.ts`
- Create: `src/commands/execute.ts`
- Create: `tests/commands/registry.test.ts`

- [ ] **Step 1: Crear `src/commands/types.ts`**

```ts
import type { Patch, SelectionContext } from "../core/types";
import type { MatrixOptions } from "../core/geometry";

export type SectionId = "alignment" | "size";

export const SECTION_LABELS: Record<SectionId, string> = {
  alignment: "Alignment",
  size: "Size",
};

/** Opciones de usuario que algunos comandos necesitan (vienen del panel). */
export interface CommandOptions {
  matrix: MatrixOptions;
}

export interface GeometryCommand {
  kind: "geometry";
  id: string;
  section: SectionId;
  label: string;
  /** id en src/taskpane/icons.ts */
  icon: string;
  minShapes: number;
  maxShapes?: number;
  compute: (selection: SelectionContext, options: CommandOptions) => Patch[];
}

export type Command = GeometryCommand;
```

- [ ] **Step 2: Crear `src/commands/alignment.ts`**

```ts
import { alignShapes, alignToSlide, arrangeMatrix, distribute, swapPositions } from "../core/geometry";
import type { Command } from "./types";

export const alignmentCommands: Command[] = [
  { kind: "geometry", id: "align-left", section: "alignment", label: "Alinear a la izquierda", icon: "alignLeft", minShapes: 2, compute: (s) => alignShapes(s.shapes, "left") },
  { kind: "geometry", id: "align-center-h", section: "alignment", label: "Centrar horizontalmente", icon: "alignCenterH", minShapes: 2, compute: (s) => alignShapes(s.shapes, "centerH") },
  { kind: "geometry", id: "align-right", section: "alignment", label: "Alinear a la derecha", icon: "alignRight", minShapes: 2, compute: (s) => alignShapes(s.shapes, "right") },
  { kind: "geometry", id: "align-top", section: "alignment", label: "Alinear arriba", icon: "alignTop", minShapes: 2, compute: (s) => alignShapes(s.shapes, "top") },
  { kind: "geometry", id: "align-center-v", section: "alignment", label: "Centrar verticalmente", icon: "alignCenterV", minShapes: 2, compute: (s) => alignShapes(s.shapes, "centerV") },
  { kind: "geometry", id: "align-bottom", section: "alignment", label: "Alinear abajo", icon: "alignBottom", minShapes: 2, compute: (s) => alignShapes(s.shapes, "bottom") },
  { kind: "geometry", id: "distribute-h", section: "alignment", label: "Distribuir horizontalmente", icon: "distributeH", minShapes: 3, compute: (s) => distribute(s.shapes, "h") },
  { kind: "geometry", id: "distribute-v", section: "alignment", label: "Distribuir verticalmente", icon: "distributeV", minShapes: 3, compute: (s) => distribute(s.shapes, "v") },
  { kind: "geometry", id: "swap", section: "alignment", label: "Intercambiar posiciones", icon: "swap", minShapes: 2, maxShapes: 2, compute: (s) => swapPositions(s.shapes) },
  { kind: "geometry", id: "matrix", section: "alignment", label: "Alinear en matriz", icon: "matrix", minShapes: 2, compute: (s, o) => arrangeMatrix(s.shapes, o.matrix) },
  { kind: "geometry", id: "slide-center-h", section: "alignment", label: "Centrar en la diapositiva (horizontal)", icon: "slideCenterH", minShapes: 1, compute: (s) => alignToSlide(s.shapes, s.slide, "centerH") },
  { kind: "geometry", id: "slide-center-v", section: "alignment", label: "Centrar en la diapositiva (vertical)", icon: "slideCenterV", minShapes: 1, compute: (s) => alignToSlide(s.shapes, s.slide, "centerV") },
  { kind: "geometry", id: "slide-center", section: "alignment", label: "Centrar en la diapositiva", icon: "slideCenter", minShapes: 1, compute: (s) => [...alignToSlide(s.shapes, s.slide, "centerH"), ...alignToSlide(s.shapes, s.slide, "centerV")] },
];
```

- [ ] **Step 3: Crear `src/commands/size.ts`**

```ts
import { matchSize, stretchToEdge } from "../core/geometry";
import type { Command } from "./types";

export const sizeCommands: Command[] = [
  { kind: "geometry", id: "match-width", section: "size", label: "Igualar ancho al primero", icon: "matchWidth", minShapes: 2, compute: (s) => matchSize(s.shapes, "width", "first") },
  { kind: "geometry", id: "match-height", section: "size", label: "Igualar alto al primero", icon: "matchHeight", minShapes: 2, compute: (s) => matchSize(s.shapes, "height", "first") },
  { kind: "geometry", id: "match-both", section: "size", label: "Igualar tamaño al primero", icon: "matchBoth", minShapes: 2, compute: (s) => matchSize(s.shapes, "both", "first") },
  { kind: "geometry", id: "match-widest", section: "size", label: "Igualar al más ancho", icon: "matchWidest", minShapes: 2, compute: (s) => matchSize(s.shapes, "width", "max") },
  { kind: "geometry", id: "match-tallest", section: "size", label: "Igualar al más alto", icon: "matchTallest", minShapes: 2, compute: (s) => matchSize(s.shapes, "height", "max") },
  { kind: "geometry", id: "match-narrowest", section: "size", label: "Igualar al más estrecho", icon: "matchNarrowest", minShapes: 2, compute: (s) => matchSize(s.shapes, "width", "min") },
  { kind: "geometry", id: "match-shortest", section: "size", label: "Igualar al más bajo", icon: "matchShortest", minShapes: 2, compute: (s) => matchSize(s.shapes, "height", "min") },
  { kind: "geometry", id: "stretch-left", section: "size", label: "Estirar hasta el borde izquierdo", icon: "stretchLeft", minShapes: 2, compute: (s) => stretchToEdge(s.shapes, "left") },
  { kind: "geometry", id: "stretch-right", section: "size", label: "Estirar hasta el borde derecho", icon: "stretchRight", minShapes: 2, compute: (s) => stretchToEdge(s.shapes, "right") },
  { kind: "geometry", id: "stretch-top", section: "size", label: "Estirar hasta el borde superior", icon: "stretchTop", minShapes: 2, compute: (s) => stretchToEdge(s.shapes, "top") },
  { kind: "geometry", id: "stretch-bottom", section: "size", label: "Estirar hasta el borde inferior", icon: "stretchBottom", minShapes: 2, compute: (s) => stretchToEdge(s.shapes, "bottom") },
];
```

- [ ] **Step 4: Crear `src/commands/registry.ts`**

```ts
import { alignmentCommands } from "./alignment";
import { sizeCommands } from "./size";
import type { Command, SectionId } from "./types";

export const SECTION_ORDER: SectionId[] = ["alignment", "size"];

export const commands: Command[] = [...alignmentCommands, ...sizeCommands];

export function commandsBySection(section: SectionId): Command[] {
  return commands.filter((c) => c.section === section);
}

export function findCommand(id: string): Command | undefined {
  return commands.find((c) => c.id === id);
}

/** Devuelve el motivo por el que un comando no aplica a `count` formas, o undefined si aplica. */
export function disabledReason(cmd: Command, count: number): string | undefined {
  if (count < cmd.minShapes) return `Selecciona al menos ${cmd.minShapes} forma${cmd.minShapes > 1 ? "s" : ""}.`;
  if (cmd.maxShapes !== undefined && count > cmd.maxShapes) return `Selecciona como máximo ${cmd.maxShapes} formas.`;
  return undefined;
}
```

- [ ] **Step 5: Escribir `tests/commands/registry.test.ts` (falla hasta que exista registry)**

```ts
import { describe, expect, it } from "vitest";
import { commands, disabledReason, findCommand } from "../../src/commands/registry";
import { shape, slide16x9 } from "../core/fixtures";

describe("registry", () => {
  it("has unique ids and icons defined", () => {
    const ids = commands.map((c) => c.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(commands.every((c) => c.icon.length > 0 && c.label.length > 0)).toBe(true);
  });

  it("reports why a command is disabled", () => {
    const swap = findCommand("swap");
    expect(swap).toBeDefined();
    expect(disabledReason(swap!, 1)).toMatch(/al menos 2/);
    expect(disabledReason(swap!, 3)).toMatch(/como máximo 2/);
    expect(disabledReason(swap!, 2)).toBeUndefined();
  });

  it("align-left computes patches via core", () => {
    const cmd = findCommand("align-left")!;
    const ref = shape({ id: "r", left: 10, top: 0, width: 10, height: 10 });
    const b = shape({ id: "b", left: 50, top: 0, width: 10, height: 10 });
    const patches = cmd.compute({ shapes: [ref, b], slide: slide16x9 }, { matrix: { columns: 2, gapX: 8, gapY: 8 } });
    expect(patches).toEqual([{ id: "b", left: 10 }]);
  });
});
```

- [ ] **Step 6: Crear `src/commands/execute.ts`**

```ts
import { applyPatches } from "../office/apply";
import { runCommand, type CommandResult } from "../office/runtime";
import { readSelection } from "../office/selection";
import type { SlideSize } from "../core/types";
import { disabledReason } from "./registry";
import type { Command, CommandOptions } from "./types";

const slideSizeCache = new Map<string, SlideSize>();

/** Lee la selección, calcula los parches y los aplica en un único sync final. */
export async function executeCommand(cmd: Command, options: CommandOptions): Promise<CommandResult> {
  return runCommand(cmd.id, async (context) => {
    const { selection, proxies } = await readSelection(context, { slideSizeCache });
    const reason = disabledReason(cmd, selection.shapes.length);
    if (reason) throw new Error(reason);
    const patches = cmd.compute(selection, options);
    if (patches.length === 0) return 0;
    applyPatches(proxies, patches);
    await context.sync();
    return new Set(patches.map((p) => p.id)).size;
  });
}
```

- [ ] **Step 7: Ejecutar tests y typecheck**

Run: `npm test && npm run typecheck`
Expected: todos `passed`; sin errores de tipos.

- [ ] **Step 8: Commit**

```bash
git add src/commands tests/commands
git commit -m "feat(commands): registro declarativo de Alignment y Size y ejecutor con un solo sync"
```

---

### Task 11: Panel: iconos SVG, ajustes y renderizado

**Files:**
- Create: `src/taskpane/icons.ts`
- Create: `src/taskpane/settings.ts`
- Create: `src/taskpane/panel.ts`
- Modify: `src/taskpane/main.ts`
- Create: `tests/taskpane/icons.test.ts`

- [ ] **Step 1: Crear `src/taskpane/icons.ts`**

```ts
/** Iconos SVG 16x16 con trazo currentColor. Las barras rojas marcan la referencia. */
const S = 'xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"';
const REF = 'stroke="#d33"';

const box = (x: number, y: number, w: number, h: number) => `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="currentColor" fill-opacity=".25"/>`;

export const icons: Record<string, string> = {
  alignLeft: `<svg ${S}><line x1="2" y1="1" x2="2" y2="15" ${REF}/>${box(3, 3, 9, 3)}${box(3, 9, 5, 3)}</svg>`,
  alignCenterH: `<svg ${S}><line x1="8" y1="1" x2="8" y2="15" ${REF}/>${box(3, 3, 10, 3)}${box(5.5, 9, 5, 3)}</svg>`,
  alignRight: `<svg ${S}><line x1="14" y1="1" x2="14" y2="15" ${REF}/>${box(4, 3, 9, 3)}${box(8, 9, 5, 3)}</svg>`,
  alignTop: `<svg ${S}><line x1="1" y1="2" x2="15" y2="2" ${REF}/>${box(3, 3, 3, 9)}${box(9, 3, 3, 5)}</svg>`,
  alignCenterV: `<svg ${S}><line x1="1" y1="8" x2="15" y2="8" ${REF}/>${box(3, 3, 3, 10)}${box(9, 5.5, 3, 5)}</svg>`,
  alignBottom: `<svg ${S}><line x1="1" y1="14" x2="15" y2="14" ${REF}/>${box(3, 4, 3, 9)}${box(9, 8, 3, 5)}</svg>`,
  distributeH: `<svg ${S}>${box(1, 4, 3, 8)}${box(6.5, 4, 3, 8)}${box(12, 4, 3, 8)}<path d="M4.5 8h1.5M10 8h1.5"/></svg>`,
  distributeV: `<svg ${S}>${box(4, 1, 8, 3)}${box(4, 6.5, 8, 3)}${box(4, 12, 8, 3)}<path d="M8 4.5v1.5M8 10v1.5"/></svg>`,
  swap: `<svg ${S}>${box(1, 2, 6, 5)}${box(9, 9, 6, 5)}<path d="M8 4.5h5l-1.5-1.5M13 4.5l-1.5 1.5M8 11.5H3l1.5-1.5M3 11.5l1.5 1.5"/></svg>`,
  matrix: `<svg ${S}>${box(1.5, 1.5, 5, 5)}${box(9.5, 1.5, 5, 5)}${box(1.5, 9.5, 5, 5)}${box(9.5, 9.5, 5, 5)}</svg>`,
  slideCenterH: `<svg ${S}><rect x="1" y="2" width="14" height="12" rx="1"/><line x1="8" y1="2" x2="8" y2="14" stroke-dasharray="1.5 1.5" ${REF}/>${box(5, 6, 6, 4)}</svg>`,
  slideCenterV: `<svg ${S}><rect x="1" y="2" width="14" height="12" rx="1"/><line x1="1" y1="8" x2="15" y2="8" stroke-dasharray="1.5 1.5" ${REF}/>${box(4, 5, 8, 6)}</svg>`,
  slideCenter: `<svg ${S}><rect x="1" y="2" width="14" height="12" rx="1"/><line x1="8" y1="2" x2="8" y2="14" stroke-dasharray="1.5 1.5" ${REF}/><line x1="1" y1="8" x2="15" y2="8" stroke-dasharray="1.5 1.5" ${REF}/>${box(5, 5.5, 6, 5)}</svg>`,
  matchWidth: `<svg ${S}>${box(2, 2, 12, 4)}${box(2, 9, 12, 4)}<path d="M2 15h12" ${REF}/></svg>`,
  matchHeight: `<svg ${S}>${box(2, 2, 4, 12)}${box(9, 2, 4, 12)}<path d="M15 2v12" ${REF}/></svg>`,
  matchBoth: `<svg ${S}>${box(1, 1, 6, 6)}${box(9, 9, 6, 6)}<path d="M9 1h6v6" ${REF}/></svg>`,
  matchWidest: `<svg ${S}>${box(1, 2, 14, 4)}${box(1, 9, 8, 4)}<path d="M9 11h6" stroke-dasharray="1 1"/></svg>`,
  matchTallest: `<svg ${S}>${box(2, 1, 4, 14)}${box(9, 1, 4, 8)}<path d="M11 9v6" stroke-dasharray="1 1"/></svg>`,
  matchNarrowest: `<svg ${S}>${box(1, 2, 6, 4)}${box(1, 9, 14, 4)}<path d="M7 11h8" stroke-dasharray="1 1" ${REF}/></svg>`,
  matchShortest: `<svg ${S}>${box(2, 1, 4, 6)}${box(9, 1, 4, 14)}<path d="M11 7v8" stroke-dasharray="1 1" ${REF}/></svg>`,
  stretchLeft: `<svg ${S}><line x1="2" y1="1" x2="2" y2="15" ${REF}/>${box(7, 5, 7, 6)}<path d="M7 8H3M4.5 6.5 3 8l1.5 1.5"/></svg>`,
  stretchRight: `<svg ${S}><line x1="14" y1="1" x2="14" y2="15" ${REF}/>${box(2, 5, 7, 6)}<path d="M9 8h4M11.5 6.5 13 8l-1.5 1.5"/></svg>`,
  stretchTop: `<svg ${S}><line x1="1" y1="2" x2="15" y2="2" ${REF}/>${box(5, 7, 6, 7)}<path d="M8 7V3M6.5 4.5 8 3l1.5 1.5"/></svg>`,
  stretchBottom: `<svg ${S}><line x1="1" y1="14" x2="15" y2="14" ${REF}/>${box(5, 2, 6, 7)}<path d="M8 9v4M6.5 11.5 8 13l1.5-1.5"/></svg>`,
};

export function iconSvg(id: string): string {
  return icons[id] ?? `<svg ${S}><rect x="2" y="2" width="12" height="12"/></svg>`;
}
```

- [ ] **Step 2: Escribir `tests/taskpane/icons.test.ts`**

```ts
import { describe, expect, it } from "vitest";
import { commands } from "../../src/commands/registry";
import { icons } from "../../src/taskpane/icons";

describe("icons", () => {
  it("every command has an icon", () => {
    const missing = commands.filter((c) => !(c.icon in icons)).map((c) => c.id);
    expect(missing).toEqual([]);
  });
});
```

Run: `npx vitest run tests/taskpane`
Expected: `1 passed`.

- [ ] **Step 3: Crear `src/taskpane/settings.ts`**

```ts
import type { CommandOptions } from "../commands/types";

const KEY = "slide-tools.options";

export const DEFAULT_OPTIONS: CommandOptions = { matrix: { columns: 3, gapX: 8, gapY: 8 } };

export function loadOptions(): CommandOptions {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return DEFAULT_OPTIONS;
    const parsed = JSON.parse(raw) as Partial<CommandOptions>;
    return { matrix: { ...DEFAULT_OPTIONS.matrix, ...parsed.matrix } };
  } catch {
    return DEFAULT_OPTIONS;
  }
}

export function saveOptions(options: CommandOptions): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(options));
  } catch {
    /* almacenamiento no disponible: se ignora */
  }
}
```

- [ ] **Step 4: Crear `src/taskpane/panel.ts`**

```ts
import { executeCommand } from "../commands/execute";
import { SECTION_ORDER, commandsBySection, disabledReason } from "../commands/registry";
import { SECTION_LABELS, type Command, type CommandOptions } from "../commands/types";
import { iconSvg } from "./icons";
import { loadOptions, saveOptions } from "./settings";

export class Panel {
  private options: CommandOptions = loadOptions();
  private buttons = new Map<string, { el: HTMLButtonElement; cmd: Command }>();
  private noticeTimer: number | undefined;

  constructor(private root: HTMLElement, private notice: HTMLElement) {}

  render(): void {
    this.root.replaceChildren();
    SECTION_ORDER.forEach((section, i) => {
      if (i > 0) this.root.appendChild(Object.assign(document.createElement("hr"), { className: "st-divider" }));
      const sec = document.createElement("section");
      sec.className = "st-section";
      const title = document.createElement("div");
      title.className = "st-section-title";
      title.textContent = SECTION_LABELS[section];
      sec.appendChild(title);
      const row = document.createElement("div");
      row.className = "st-row";
      for (const cmd of commandsBySection(section)) row.appendChild(this.button(cmd));
      sec.appendChild(row);
      if (section === "alignment") sec.appendChild(this.matrixFields());
      this.root.appendChild(sec);
    });
    this.setSelectionCount(0);
  }

  /** Habilita/deshabilita botones según el número de formas seleccionadas. */
  setSelectionCount(count: number): void {
    for (const { el, cmd } of this.buttons.values()) {
      const reason = disabledReason(cmd, count);
      el.disabled = reason !== undefined;
      el.title = reason ? `${cmd.label} — ${reason}` : cmd.label;
    }
  }

  showNotice(text: string, isError = false): void {
    this.notice.textContent = text;
    this.notice.classList.toggle("is-error", isError);
    this.notice.hidden = false;
    window.clearTimeout(this.noticeTimer);
    this.noticeTimer = window.setTimeout(() => (this.notice.hidden = true), 3000);
  }

  private button(cmd: Command): HTMLButtonElement {
    const el = document.createElement("button");
    el.className = "st-btn";
    el.type = "button";
    el.innerHTML = iconSvg(cmd.icon);
    el.setAttribute("aria-label", cmd.label);
    el.addEventListener("click", () => void this.run(cmd));
    this.buttons.set(cmd.id, { el, cmd });
    return el;
  }

  private async run(cmd: Command): Promise<void> {
    const result = await executeCommand(cmd, this.options);
    if (!result.ok) this.showNotice(result.message, true);
    else if (result.changed === 0) this.showNotice("Nada que cambiar.");
  }

  private matrixFields(): HTMLElement {
    const wrap = document.createElement("div");
    wrap.className = "st-row";
    const make = (label: string, key: "columns" | "gapX" | "gapY", min: number) => {
      const field = document.createElement("label");
      field.className = "st-field";
      field.append(label);
      const input = document.createElement("input");
      input.type = "number";
      input.min = String(min);
      input.value = String(this.options.matrix[key]);
      input.addEventListener("change", () => {
        const value = Math.max(min, Number(input.value) || min);
        input.value = String(value);
        this.options = { ...this.options, matrix: { ...this.options.matrix, [key]: value } };
        saveOptions(this.options);
      });
      field.appendChild(input);
      return field;
    };
    wrap.append(make("Cols", "columns", 1), make("Gap X", "gapX", 0), make("Gap Y", "gapY", 0));
    return wrap;
  }
}
```

- [ ] **Step 5: Reescribir `src/taskpane/main.ts`**

```ts
import { Panel } from "./panel";

async function selectedShapeCount(): Promise<number> {
  try {
    return await PowerPoint.run(async (context) => {
      const count = context.presentation.getSelectedShapes().getCount();
      await context.sync();
      return count.value;
    });
  } catch {
    return 0;
  }
}

Office.onReady((info) => {
  const root = document.getElementById("panel");
  const notice = document.getElementById("notice");
  if (!root || !notice) return;

  if (info.host !== Office.HostType.PowerPoint) {
    root.textContent = "Este panel requiere PowerPoint.";
    return;
  }

  const panel = new Panel(root, notice);
  panel.render();

  const refresh = async () => panel.setSelectionCount(await selectedShapeCount());
  void refresh();
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => void refresh(), (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.warn("[slide-tools] no se pudo suscribir al cambio de selección:", result.error.message);
    }
  });
});
```

- [ ] **Step 6: Build y tests**

Run: `npm test && npm run build`
Expected: todos `passed`; `dist/` generado.

- [ ] **Step 7: Prueba manual en PowerPoint Mac**

Run (una sola vez, pide contraseña para instalar la CA de desarrollo): `npm run certs`
Run: `npm run manifest:dev && npm run sideload:mac`
Run en otra terminal: `npm run dev`

En PowerPoint: cerrar y reabrir, Insertar > Complementos > Mis complementos > Complementos para desarrolladores > "Slide Tools (dev)". Luego Inicio > "Abrir panel".

Checklist:
1. El panel muestra "Slide Tools / Victor Lopez" y las secciones Alignment y Size.
2. Sin selección, todos los botones deshabilitados; con 1 forma, solo los tres de "diapositiva"; con 2, todos salvo distribuir; con 3, distribuir activo y "Intercambiar" deshabilitado.
3. Seleccionar 3 rectángulos de tamaños distintos, pulsar "Alinear a la izquierda": se alinean al primero seleccionado. Cmd+Z lo deshace de una vez.
4. "Distribuir horizontalmente": separación uniforme, extremos fijos.
5. "Centrar en la diapositiva" con una forma en una presentación 16:9: queda centrada. Repetir en una presentación 4:3 (Diseño > Tamaño de diapositiva): también centrada.
6. "Igualar ancho al primero" y "Estirar hasta el borde derecho" se comportan como describen sus tooltips.
7. Cambiar "Cols" a 2 y pulsar "Alinear en matriz" con 4 formas: rejilla de 2×2.

Si algo falla, registrar el error de consola (PowerPoint Mac: clic derecho en el panel > Inspeccionar elemento) y corregir antes de seguir.

- [ ] **Step 8: Commit**

```bash
git add src/taskpane tests/taskpane
git commit -m "feat(taskpane): panel con secciones Alignment y Size, iconos SVG y ajustes de matriz"
```

---

### Task 12: Despliegue a GitHub Pages

**Files:**
- Create: `.github/workflows/deploy.yml`

- [ ] **Step 1: Crear `.github/workflows/deploy.yml`**

```yaml
name: Deploy to GitHub Pages

on:
  push:
    branches: [main]
  workflow_dispatch:

permissions:
  contents: read
  pages: write
  id-token: write

concurrency:
  group: pages
  cancel-in-progress: true

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 22
          cache: npm
      - run: npm ci
      - run: npm test
      - run: npm run build
      - run: cp manifest.xml dist/manifest.xml
      - uses: actions/upload-pages-artifact@v3
        with:
          path: dist

  deploy:
    needs: build
    runs-on: ubuntu-latest
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    steps:
      - id: deployment
        uses: actions/deploy-pages@v4
```

- [ ] **Step 2: Hacer público el repo y activar Pages con origen "GitHub Actions"**

Run:
```bash
gh repo edit xictorlr/skillsEEpowerPoint --visibility public --accept-visibility-change-consequences
gh api -X POST repos/xictorlr/skillsEEpowerPoint/pages -f build_type=workflow
```
Expected: segundo comando devuelve JSON con `"build_type": "workflow"` (o 409 si ya existía: aceptable).

- [ ] **Step 3: Commit, push y comprobar el despliegue**

```bash
git add .github/workflows/deploy.yml
git commit -m "ci: build, tests y despliegue a GitHub Pages"
git push
gh run watch --exit-status
curl -sI https://xictorlr.github.io/skillsEEpowerPoint/index.html | head -1
curl -sI https://xictorlr.github.io/skillsEEpowerPoint/assets/icon-32.png | head -1
```
Expected: run en verde; ambos `HTTP/2 200`.

- [ ] **Step 4: Instalar el manifest de producción en Mac y verificar**

Run: `bash scripts/sideload-mac.sh manifest.xml`
En PowerPoint (reiniciado): Mis complementos > "Slide Tools" (sin "dev"). Abrir panel y repetir los puntos 3 y 5 del checklist de la Task 11.

---

### Task 13: README

**Files:**
- Create: `README.md`

- [ ] **Step 1: Escribir `README.md`**

```markdown
# Slide Tools

Panel lateral para PowerPoint (Mac y Windows) con herramientas de alineación, tamaño, formato y texto,
inspirado en el flujo de trabajo de los paneles de productividad para presentaciones.

Autor: Victor Lopez · Add-in web (Office.js) · Hospedado en GitHub Pages.

## Instalación

### Mac

1. Descarga `manifest.xml` desde https://xictorlr.github.io/skillsEEpowerPoint/manifest.xml
2. Cópialo a `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` (crea la carpeta si no existe).
   Con el repo clonado: `bash scripts/sideload-mac.sh manifest.xml`
3. Reinicia PowerPoint. Insertar > Complementos > Mis complementos > Complementos para desarrolladores > Slide Tools.
4. En la pestaña Inicio aparece el botón "Abrir panel".

### Windows

1. Descarga `manifest.xml` a una carpeta y compártela (clic derecho > Propiedades > Compartir > compartir contigo mismo).
   Copia la ruta de red (`\\TU-PC\carpeta`).
2. PowerPoint > Archivo > Opciones > Centro de confianza > Configuración > Catálogos de complementos de confianza.
   Añade la ruta de red, marca "Mostrar en el menú" y acepta.
3. Reinicia PowerPoint. Insertar > Mis complementos > Carpeta compartida > Slide Tools.

## Comandos (Fase 1)

**Alignment:** alinear izquierda / centro / derecha / arriba / medio / abajo (a la primera forma seleccionada),
distribuir horizontal y vertical, intercambiar posiciones, alinear en matriz (columnas y separación configurables),
centrar en la diapositiva (horizontal, vertical, ambos).

**Size:** igualar ancho / alto / ambos al primero, igualar al más ancho / alto / estrecho / bajo,
estirar hasta el borde izquierdo / derecho / superior / inferior de la referencia.

La **referencia** es siempre la primera forma que seleccionaste. Cada comando es un único paso de deshacer.

## Desarrollo

```bash
npm install
npm run certs          # una vez: certificados HTTPS de desarrollo
npm run manifest:dev   # genera manifest.dev.xml (localhost:3000)
npm run sideload:mac   # copia el manifest dev a PowerPoint
npm run dev            # servidor Vite
npm test               # Vitest
npm run build          # dist/
```

Estructura: `src/core` (lógica pura con tests), `src/office` (adaptador Office.js), `src/commands` (comandos
declarativos), `src/taskpane` (UI). Diseño completo en `docs/superpowers/specs/`.

## Hoja de ruta

Fase 2 Format + Color · Fase 3 Text · Fase 4 Navigation & Expert Tools · Fase 5 Check + Presentation · Fase 6 Wizards.
```

- [ ] **Step 2: Commit y push**

```bash
git add README.md
git commit -m "docs: README con instalación Mac/Windows, comandos y desarrollo"
git push
```

---

## Autorrevisión del plan

**Cobertura del spec (Fase 1 + andamiaje):**
- Alinear 6 bordes/centros a la primera forma → Task 4, 10. ✔
- Distribuir H/V → Task 5, 10. ✔
- Matriz configurable → Task 7, 10, 11 (campos Cols/Gap). ✔
- Intercambiar posiciones → Task 5, 10. ✔
- Alinear a la diapositiva → Task 4, 8, 9, 10. ✔
- Igualar ancho/alto/ambos al primero; al más ancho/alto (y además al más estrecho/bajo) → Task 6, 10. ✔
- Estirar hasta bordes → Task 6, 10. ✔
- "Ajustar alto al contenido": requiere `TextFrame.autoSizeSetting`; se pospone a la Fase 3 (Text), donde vive el resto de autofit. Anotado en el spec como parte de Text.
- Un `sync` final por comando → Task 10 (`executeCommand`). ✔
- Botones deshabilitados con motivo + evento de selección → Task 10 (`disabledReason`), Task 11. ✔
- Wrapper de errores → Task 9 (`runCommand`). ✔
- Tests core ≥ 90 % → Task 7 Step 4 ejecuta cobertura. ✔
- Manifest, sideload Mac, instrucciones Windows, Pages → Task 3, 12, 13. ✔

**Placeholders:** ninguno; todo el código está en los pasos.

**Consistencia de tipos:** `Patch`, `ShapeModel`, `SelectionContext` (Task 4) se usan igual en Task 9–11. `readSelection(context, { slideSizeCache })` coincide entre Task 9 y Task 10. `CommandResult` definido en runtime.ts y consumido en execute.ts y panel.ts. `MatrixOptions` exportado de geometry.ts y reutilizado en commands/types.ts.
