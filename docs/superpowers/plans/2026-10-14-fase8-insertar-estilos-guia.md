# Fase 8: Insertar elementos, estilos rápidos de texto, guía de instalación y créditos — Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Añadir (a) una sección **Insert** inspirada en los grupos "Text Elements" y "Graphic Elements" de la cinta de referencia: columnas de texto, rejilla, cadena de chevrones, círculos numerados, sticker, líneas y formas básicas; (b) estilos rápidos en **Text**: negrita/cursiva/subrayado/tachado (conmutadores), tamaños predefinidos, viñetas sí/no, sangría ±; (c) una guía de instalación para principiantes (Windows y Mac) y (d) créditos del autor.

**Architecture:** Sin cambios: `core/insert.ts` (layouts puros), `office/insert.ts` (alta de formas con estilo y selección de lo creado), comandos `action` con `needsSelection: false` para Insert y `text` para estilos; nuevos campos `TextShape` (bold/italic/underline/strikethrough/bulletsVisible/indentLevel).

**API:** `addTextBox`, `addGeometricShape` (Rectangle, RoundRectangle, Ellipse, Chevron), `addLine("Straight"|"Elbow")` (1.4); `ShapeFont.bold/italic/underline/strikethrough` (1.4/1.8); `BulletFormat.visible` (1.4); `ParagraphFormat.indentLevel` (1.10, gated); `Slide.setSelectedShapes` (1.5).

---

### Task 1: Core y adaptador de inserción

**Files:** `src/core/insert.ts`, `src/office/insert.ts`, `tests/core/insert.test.ts`, `tests/office/insert.test.ts`, `tests/office/mockPowerPoint.ts` (additivo).

Core (`insert.ts`), área de contenido = márgenes 8 % lateral, 20 % superior, 8 % inferior:
```ts
export interface ContentArea { left; top; width; height }
export function contentArea(slide: SlideSize): Rect
export function columnsLayout(slide, count: number, gapRatio = 0.02): Rect[]          // 1..6 columnas iguales
export function gridLayout(slide, rows: number, cols: number, gapRatio = 0.02): Rect[]  // fila a fila
export function chevronChainLayout(slide, count: number): Rect[]   // fila de chevrones de alto 12 % centrada verticalmente en el área, solape 0.
export function numberedCirclesLayout(slide, count: number): Rect[] // círculos de diámetro 7 % del ancho, repartidos en una fila
export function stickerRect(slide): Rect                            // esquina superior derecha: 18 % × 7 %, margen 3 %
export const BASIC_SHAPE_RECT = (slide) => Rect                     // centro, 25 % × 25 % (cuadrado por el menor lado)
```
Tests con 960×540: `columnsLayout(3)` → 3 rects de ancho (806.4 − 2·19.2)/3 con lefts 76.8, 76.8+w+19.2…; `gridLayout(2,2)` → 4; `chevronChainLayout(4)`; `numberedCirclesLayout(5)`; `stickerRect` → `{ left: 960·(1−0.03−0.18), top: 540·0.03, width: 172.8, height: 37.8 }`; count fuera de 1..12 → clamp.

Office (`insert.ts`):
```ts
export interface ShapeStyle { fill?: string | null; line?: string | null; fontSize?: number; bold?: boolean; fontColor?: string | null; text?: string }
export function addTextBoxes(slide, rects: Rect[], texts: string[], style?: ShapeStyle): PowerPoint.Shape[]
export function addGeometric(slide, type: "Rectangle" | "RoundRectangle" | "Ellipse" | "Chevron", rect: Rect, style?: ShapeStyle): PowerPoint.Shape
export function addLine(slide, kind: "Straight" | "Elbow", rect: Rect): PowerPoint.Shape
/** Tras un sync con los ids cargados, selecciona las formas creadas. */
export function selectCreated(slide, shapes: PowerPoint.Shape[]): void   // slide.setSelectedShapes(ids)
```
Mock: `shapes.addGeometricShape(type, opts)` y `shapes.addLine(kind, opts)` crean `MockShape` (type "GeometricShape"/"Line") y registran en `ctx.calls`; `fill.setSolidColor`, `lineFormat.visible/color`, `textFrame.textRange.text/font` ya existen.

Commit: `feat(insert): layouts y alta de elementos (columnas, rejilla, chevrones, círculos, sticker, líneas, formas)`.

### Task 2: Estilos rápidos de texto

**Files:** `src/core/text.ts`, `src/office/text.ts`, tests.

- `TextShape` + `bold: boolean | null`, `italic`, `underline: string | null`, `strikethrough: boolean | null`, `bulletsVisible: boolean | null`, `indentLevel: number | null`. `TextPatch` + `bold?`, `italic?`, `underline?: string`, `strikethrough?`, `bulletsVisible?`, `indentLevel?`.
- `toggleBoldPatches(shapes)`: si TODAS las formas con valor no nulo son `true` → `false` para todas; si no → `true`. Igual `toggleItalicPatches`, `toggleStrikethroughPatches`; `toggleUnderlinePatches` (todas "Single"/no "None" → "None"; si no → "Single"). `toggleBulletsPatches`. `fontSizePresetPatches(shapes, size)` (omite mixtos no: aplica a todas). `indentPatches(shapes, delta)` clamp 0..8, omite `null`.
- `office/text.ts`: cargar `font.bold,italic,underline,strikethrough`, `paragraphFormat.bulletFormat.visible`, y `paragraphFormat.indentLevel` solo si `supportsApi("1.10")`; escribir cada campo del parche. Mock: `paragraphFormat.bulletFormat.visible` accessor, `indentLevel` accessor.
- Tests para cada generador (todas true → false; mixto → true; nulos ignorados) y para las escrituras.

Commit: `feat(text): conmutadores de estilo, tamaños predefinidos, viñetas y sangría`.

### Task 3: Guía de instalación y créditos

**Files:** `docs/INSTALACION.md` (nuevo), `README.md`, `package.json` (author), `manifest.xml` (ProviderName), `src/taskpane/index.html`, `src/taskpane/styles.css`.

- `docs/INSTALACION.md`: guía paso a paso sin jerga, con secciones "¿Qué es esto?", "Mac (5 minutos)", "Windows (5 minutos)", "Cómo saber si funciona", "Problemas frecuentes" (no aparece el complemento → reiniciar PowerPoint / vaciar caché Wef; el panel sale en blanco → conexión a internet; versión antigua de PowerPoint → qué botones faltan). Mac: descargar el manifest desde la URL, Finder > Ir a la carpeta `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef` (crear si no existe), copiar, reiniciar, Insertar > Complementos > Mis complementos > Complementos para desarrolladores. Windows: guardar el manifest en una carpeta, compartirla (clic derecho > Propiedades > Compartir), copiar la ruta de red `\\EQUIPO\carpeta`, PowerPoint > Archivo > Opciones > Centro de confianza > Configuración > Catálogos de complementos de confianza > Añadir ruta > marcar "Mostrar en el menú" > reiniciar > Insertar > Mis complementos > Carpeta compartida. Incluir comandos de terminal como alternativa para Mac.
- README: enlace a la guía al principio de "Instalación"; sección "Créditos": Victor Lopez Rapado · https://xictorlr.eu · https://www.linkedin.com/in/victorlopezrapado/.
- `package.json`: `"author": "Victor Lopez Rapado (https://xictorlr.eu)"`; `manifest.xml`: `<ProviderName>Victor Lopez Rapado</ProviderName>`.
- Panel: cabecera con el nombre completo enlazado a `https://xictorlr.eu` (`target="_blank" rel="noopener"`) y un icono/enlace "in" a LinkedIn; estilo discreto (`.st-author a`). Actualizar el test DOM si comprueba el texto de la cabecera.

Commit: `docs: guía de instalación para principiantes y créditos del autor`.

### Task 4: Comandos y panel

**Files:** `src/commands/types.ts`, `registry.ts`, `src/commands/insert.ts` (nuevo), `src/commands/text.ts`, `src/taskpane/sections/insertSection.ts` (nuevo), `panel.ts`, `sanitize.ts`, `icons.ts`, README, versión.

- `SectionId` + `"insert"` (después de "text"); `CommandOptions.insert: { count: number; stickerText: string; stickerColor: string }` persistido (count 1..12 default 3; stickerText ≤ 30 default "BORRADOR"; color hex default "#FFE066").
- Comandos `action` (`needsSelection: false`, usan `readSlideNav` + `getSlideSize` + `slides.getItem(id)`; crean, `load("items/id")`→ sync → `selectCreated` → sync → `{ changed: n, message }`):
  `insert-columns` (N columnas de texto "Texto"), `insert-grid` (2×2), `insert-chevrons` (N chevrones con "Paso i", relleno accent/`#1F6FEB`, texto blanco), `insert-circles` (N círculos numerados), `insert-sticker` (RoundRectangle con texto, relleno `stickerColor`, negrita 12), `insert-line`, `insert-elbow`, `insert-rect`, `insert-rounded`, `insert-ellipse`, `insert-textbox`.
- Comandos `text`: `style-bold`, `style-italic`, `style-underline`, `style-strike`, `bullets-toggle`, `size-title` (24), `size-subtitle` (18), `size-body` (14), `size-small` (11), `indent-more`, `indent-less` (`requiresApi: "1.10"`).
- `insertSection.ts`: campo "N" (1..12), campo "Sticker" (texto), color del sticker (swatch/`input type=color`). Iconos para todos.
- README "Comandos (Fase 8)", prueba manual paso 10, versión 1.6.0 / 1.6.0.0.

Commits: `feat(commands): sección Insert y estilos rápidos de texto` · `feat(taskpane): sección Insert` · `docs: README Fase 8 y versión 1.6.0`.

---

## Autorrevisión
Cubre la petición del usuario (elementos de texto/gráficos de la cinta de referencia que el API permite, guía de instalación, créditos). Tipos: `Rect[]` de `core/insert` ↔ `office/insert` ↔ comandos; nuevos campos `TextShape`/`TextPatch` ↔ `office/text` ↔ comandos `text`.
