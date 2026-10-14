# Fase 4: Navigation & View + Expert Tools — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Añadir las secciones Navigation (ir a diapositiva, anterior/siguiente/primera/última) y Expert (seleccionar formas por tipo o nombre, renombrar con prefijo numerado, orden Z, agrupar/desagrupar, inspector de posición y tamaño), y aprovechar PowerPointApi 1.10 cuando exista para el tamaño exacto de diapositiva y los colores del tema.

**Architecture:** Igual que las fases anteriores. Se añade un tipo de comando `action` cuyo `run` recibe el contexto Office completo (lectura de selección ya hecha) y devuelve `CommandOutcome`, para operaciones que no encajan en el modelo de parches (seleccionar, agrupar, cambiar de diapositiva). La lógica decidible sin Office (filtros por tipo/nombre, orden de aplicación del z-order, nombres numerados, caja envolvente e inspector) va en `src/core/expert.ts` y se testea en Node. Las capacidades por versión de API se consultan en `src/office/capabilities.ts` (`supportsApi("1.10")`), con fallback a la heurística actual.

**Tech Stack:** el existente. API: `Presentation.setSelectedSlides` y `Slide.setSelectedShapes` (1.5), `Shape.setZOrder` y `ShapeScopedCollection.group()` / `ShapeGroup.ungroup()` (1.8), `Slide.index` (1.8), `Presentation.pageSetup.slideWidth/slideHeight` y `Slide.themeColorScheme.getThemeColor` (1.10).

**Spec:** `docs/superpowers/specs/2026-10-14-slide-tools-design.md` § 5 Fase 4.

---

## Mapa de archivos

| Archivo | Responsabilidad |
|---|---|
| `src/office/capabilities.ts` | `supportsApi(version)` sobre `Office.context.requirements.isSetSupported`, con guardas para tests |
| `src/office/slideSize.ts` (modificar) | Si 1.10: `pageSetup` exacto; si no, heurística PNG actual |
| `src/office/theme.ts` | `getThemeColors(context)` → 12 colores del tema (1.10) o `null` |
| `src/core/expert.ts` | `ShapeKind`, `kindOf(type)`, `filterByKind`, `filterByName`, `numberedNames`, `zOrderApplyOrder`, `boundingBox`, `applyBoundingBox` |
| `src/office/navigation.ts` | `readSlideNav`, `selectSlideByIndex`, `readSlideShapes`, `selectShapes`, `groupSelected`, `ungroupSelected`, `renameShapes`, `setZOrder` |
| `src/commands/types.ts` (modificar) | `ActionCommand`, `ActionContext`; `SectionId` + `"navigation" | "expert"`; `CommandOptions` + `gotoSlide`, `nameFilter`, `renamePrefix`, `inspector` |
| `src/commands/navigation.ts`, `src/commands/expert.ts` | Comandos |
| `src/commands/registry.ts`, `execute.ts` (modificar) | Registro y rama `action` |
| `src/taskpane/sections/navigationSection.ts`, `expertSection.ts` | Controles (campo "Ir a", filtro de nombre, prefijo, inspector) |
| `src/taskpane/sections/colorSection.ts` (modificar) | Fila "Tema" con los colores del tema cuando existan |
| `src/taskpane/main.ts`, `panel.ts` (modificar) | El refresco de selección devuelve también `slideId`, `slideIndex`, `slideCount` y `bbox`; el panel actualiza el inspector y la fila Tema |
| `src/taskpane/icons.ts` (modificar) | ~20 iconos |
| `tests/core/expert.test.ts`, `tests/office/{capabilities,theme,navigation,slideSize}.test.ts`, `tests/commands/expert.test.ts`, `tests/taskpane/settings.test.ts` | Tests |
| `README.md`, `manifest.xml`, `package.json` | Docs y versión 1.3.0 |

---

### Task 0: Capacidades, tamaño exacto de diapositiva y colores del tema (API 1.10)

**Files:**
- Create: `src/office/capabilities.ts`, `src/office/theme.ts`
- Modify: `src/office/slideSize.ts`
- Create: `tests/office/capabilities.test.ts`, `tests/office/theme.test.ts`; Modify: `tests/office/selection.test.ts`, `tests/office/mockPowerPoint.ts`

- [ ] **Step 1: `src/office/capabilities.ts`**

```ts
/** Consulta soporte de un requirement set de PowerPointApi. Devuelve false si Office no está disponible. */
export function supportsApi(version: string): boolean {
  try {
    return typeof Office !== "undefined" && Office.context?.requirements?.isSetSupported("PowerPointApi", version) === true;
  } catch {
    return false;
  }
}
```

Test: con `vi.stubGlobal("Office", { context: { requirements: { isSetSupported: (n: string, v: string) => n === "PowerPointApi" && v === "1.10" } } })` → `supportsApi("1.10")` true, `supportsApi("1.11")` false; sin global → false.

- [ ] **Step 2: `src/office/slideSize.ts`**

`getSlideSize(context, cache, override?)`: si `override` → devolverlo; si caché → devolverla; si `supportsApi("1.10")`: `const ps = context.presentation.pageSetup; ps.load("slideWidth,slideHeight"); await context.sync(); size = { width: ps.slideWidth, height: ps.slideHeight }` (validar finitos > 0; si no, caer a la heurística); cachear y devolver. Si no 1.10: heurística PNG actual. Mock: `MockContext.presentation.pageSetup = { slideWidth, slideHeight, load() }` configurable (`new MockContext(shapes, png, { pageSetup: { slideWidth: 720, slideHeight: 405 } })`). Tests: con `Office` stub 1.10 y pageSetup 720×405 → `{720,405}`, `imageOptions` undefined, 1 sync; sin 1.10 → heurística como hoy.

- [ ] **Step 3: `src/office/theme.ts`**

```ts
export const THEME_COLOR_NAMES = ["Dark1", "Light1", "Dark2", "Light2", "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink"] as const;
export type ThemeColorName = (typeof THEME_COLOR_NAMES)[number];
export interface ThemeColor { name: ThemeColorName; hex: string }

/** Colores del tema de la diapositiva activa (API 1.10). null si no hay soporte o falla. Hace un sync. */
export async function getThemeColors(context: PowerPoint.RequestContext): Promise<ThemeColor[] | null>
```

Implementación: si `!supportsApi("1.10")` → null; `const slide = context.presentation.getSelectedSlides().getItemAt(0); const scheme = slide.themeColorScheme; const results = THEME_COLOR_NAMES.map((n) => scheme.getThemeColor(n)); await context.sync();` mapear con `normalizeHex(results[i].value)` (descartar nulos). Capturar errores → null. Mock: `getSelectedSlides().getItemAt(0).themeColorScheme.getThemeColor(name)` → `{ value }` desde un mapa configurable. Tests: 12 colores normalizados; sin soporte → null; error → null.

- [ ] **Step 4: Verificar y commit**

Run: `npm test && npm run typecheck`.

```bash
git add src/office tests/office
git commit -m "feat(office): tamaño exacto de diapositiva y colores del tema con PowerPointApi 1.10"
```

---

### Task 1: Lógica pura de expert (`core/expert.ts`)

**Files:**
- Create: `src/core/expert.ts`, `tests/core/expert.test.ts`

- [ ] **Step 1: Test que falla** (casos clave)

```ts
import { describe, expect, it } from "vitest";
import { applyBoundingBox, boundingBox, filterByKind, filterByName, kindOf, numberedNames, zOrderApplyOrder } from "../../src/core/expert";
import { shape } from "./fixtures";

describe("kindOf / filterByKind", () => {
  it("maps PowerPoint types to kinds", () => {
    expect(kindOf("TextBox")).toBe("text"); expect(kindOf("GeometricShape")).toBe("shape"); expect(kindOf("Image")).toBe("image");
    expect(kindOf("Line")).toBe("line"); expect(kindOf("Table")).toBe("table"); expect(kindOf("Group")).toBe("group");
    expect(kindOf("Chart")).toBe("other"); expect(kindOf("Placeholder")).toBe("text");
  });
  it("filters by kind or returns all", () => {
    const a = shape({ id: "a", type: "TextBox", left: 0, top: 0, width: 1, height: 1 });
    const b = shape({ id: "b", type: "Image", left: 0, top: 0, width: 1, height: 1 });
    expect(filterByKind([a, b], "image").map((s) => s.id)).toEqual(["b"]);
    expect(filterByKind([a, b], "all").map((s) => s.id)).toEqual(["a", "b"]);
  });
});

describe("filterByName", () => {
  it("is case- and accent-insensitive substring match; empty filter matches nothing", () => {
    const a = shape({ id: "a", name: "Título principal", left: 0, top: 0, width: 1, height: 1 });
    expect(filterByName([a], "titulo").map((s) => s.id)).toEqual(["a"]);
    expect(filterByName([a], "")).toEqual([]);
  });
});

describe("numberedNames", () => {
  it("numbers in reading order with zero padding by count", () => {
    const a = shape({ id: "a", left: 100, top: 0, width: 1, height: 1 });
    const b = shape({ id: "b", left: 0, top: 0, width: 1, height: 1 });
    expect(numberedNames([a, b], "Caja")).toEqual([{ id: "b", name: "Caja 1" }, { id: "a", name: "Caja 2" }]);
    const many = Array.from({ length: 10 }, (_, i) => shape({ id: `s${i}`, left: i, top: 0, width: 1, height: 1 }));
    expect(numberedNames(many, "X")[0]?.name).toBe("X 01");
  });
  it("skips shapes that already have the target name", () => {
    const a = shape({ id: "a", name: "Caja 1", left: 0, top: 0, width: 1, height: 1 });
    expect(numberedNames([a], "Caja")).toEqual([]);
  });
});

describe("zOrderApplyOrder", () => {
  it("applies bottom-most first for BringToFront/SendBackward and top-most first for SendToBack/BringForward", () => {
    const low = shape({ id: "low", zOrder: 1, left: 0, top: 0, width: 1, height: 1 });
    const high = shape({ id: "high", zOrder: 5, left: 0, top: 0, width: 1, height: 1 });
    expect(zOrderApplyOrder([high, low], "BringToFront").map((s) => s.id)).toEqual(["low", "high"]);
    expect(zOrderApplyOrder([low, high], "SendToBack").map((s) => s.id)).toEqual(["high", "low"]);
  });
});

describe("boundingBox / applyBoundingBox", () => {
  const a = shape({ id: "a", left: 10, top: 10, width: 100, height: 50 });
  const b = shape({ id: "b", left: 60, top: 40, width: 100, height: 50 });
  it("computes the union rect", () => {
    expect(boundingBox([a, b])).toEqual({ left: 10, top: 10, width: 150, height: 80 });
    expect(boundingBox([])).toBeNull();
  });
  it("moves all shapes when only left/top change", () => {
    expect(applyBoundingBox([a, b], { left: 20, top: 10, width: 150, height: 80 })).toEqual([{ id: "a", left: 20 }, { id: "b", left: 70 }]);
  });
  it("scales positions and sizes proportionally when width/height change", () => {
    expect(applyBoundingBox([a, b], { left: 10, top: 10, width: 300, height: 80 })).toEqual([
      { id: "a", width: 200 }, { id: "b", left: 110, width: 200 },
    ]);
  });
  it("ignores non-finite or non-positive targets", () => {
    expect(applyBoundingBox([a], { left: NaN, top: 10, width: 0, height: 50 })).toEqual([]);
  });
});
```

- [ ] **Step 2: Implementar** `src/core/expert.ts` (tipos: `ShapeKind = "all" | "text" | "shape" | "image" | "line" | "table" | "group" | "other"`; `ZOrderOp = "BringToFront" | "BringForward" | "SendBackward" | "SendToBack"`; `boundingBox(shapes): Rect | null`; `applyBoundingBox(shapes, target: Rect): Patch[]` con `near` y escalado `sx = target.width / bbox.width`, `sy = ...`, posiciones relativas `left = target.left + (s.left - bbox.left) * sx`, tamaños `width = s.width * sx`; `filterByName` normaliza con `normalize("NFD").replace(/\p{M}/gu, "").toLocaleLowerCase("es")`; `numberedNames` ordena por `top` con tolerancia 2 pt y luego `left` (reutilizar la idea de `readingOrder` de geometry.ts: exportarla desde geometry si está privada, o duplicar 10 líneas con comentario); padding `String(n).padStart(String(total).length, "0")`).

- [ ] **Step 3: Verificar y commit**: `npx vitest run tests/core`, `git commit -m "feat(core): filtros, nombres numerados, orden de z-order y caja envolvente"`.

---

### Task 2: Adaptador Office de navegación y expert

**Files:**
- Create: `src/office/navigation.ts`, `tests/office/navigation.test.ts`; Modify: `tests/office/mockPowerPoint.ts`

API del módulo (todas sin sync propio salvo donde se indica):

```ts
export interface SlideNav { slideId: string; index: number; count: number; ids: string[] }
/** Un sync. Diapositiva activa = primer item de getSelectedSlides(). Lanza si no hay diapositivas. */
export async function readSlideNav(context): Promise<SlideNav>
/** Selecciona la diapositiva por índice (0..count-1) con setSelectedSlides([id]). Requiere sync del llamador. */
export function selectSlideByIndex(context, nav: SlideNav, index: number): boolean  // false si fuera de rango
/** Un sync: formas de la diapositiva activa como ShapeModel[] (id, name, type, geometría, zOrder). */
export async function readSlideShapes(context, slideId: string): Promise<ShapeModel[]>
export function selectShapes(context, slideId: string, ids: string[]): void   // slide.setSelectedShapes(ids)
export function groupSelected(context): void                                   // getSelectedShapes().group()
export function ungroupShapes(proxies: Map<string, PowerPoint.Shape>, groupIds: string[]): void  // shape.group.ungroup()
export function renameShapes(proxies, names: Array<{ id: string; name: string }>): void
export function applyZOrder(proxies, orderedIds: string[], op: ZOrderOp): void  // shape.setZOrder(op) en el orden dado
```

Mock: `MockContext.presentation.slides = { items: MockSlide[], load, getItem(id) }`, `getSelectedSlides()` devuelve `{ getItemAt(0) → MockSlide, items, load }`; `MockSlide { id, index, shapes: { items: MockShape[], load }, setSelectedShapes(ids) registra en ctx.selectedShapeIds, themeColorScheme, getImageAsBase64 }`; `presentation.setSelectedSlides(ids)` registra `ctx.selectedSlideIds`; `MockShape.setZOrder(op)` registra `{ path: "setZOrder", value: op }`; `MockShape.group = { ungroup() }` registra `{ path: "ungroup" }`; `name` accessor registra `{ path: "name", value }`; `presentation.getSelectedShapes().group()` registra `ctx.grouped = true`.

Tests: `readSlideNav` con 3 slides y la segunda activa → `{ index: 1, count: 3 }`; `selectSlideByIndex` fuera de rango → false y nada seleccionado; `readSlideShapes` mapea; `selectShapes`/`groupSelected`/`ungroupShapes`/`renameShapes`/`applyZOrder` registran las llamadas en orden.

Commit: `feat(office): navegación de diapositivas, selección, agrupar, renombrar y z-order`.

---

### Task 3: Comandos `action` y ejecutor

**Files:**
- Modify: `src/commands/types.ts`, `registry.ts`, `execute.ts`; Create: `src/commands/navigation.ts`, `src/commands/expert.ts`, `tests/commands/expert.test.ts`

- [ ] `types.ts`:

```ts
export type SectionId = "alignment" | "size" | "format" | "color" | "text" | "navigation" | "expert";
// SECTION_LABELS: navigation: "Navigation", expert: "Expert"
export interface ActionContext {
  context: PowerPoint.RequestContext;
  selection: SelectionContext;
  proxies: Map<string, PowerPoint.Shape>;
  options: CommandOptions;
}
export interface ActionCommand extends BaseCommand {
  kind: "action";
  run: (ctx: ActionContext) => Promise<CommandOutcome>;
}
// CommandOptions + : gotoSlide?: number (1-based); nameFilter: string; renamePrefix: string; inspector?: Partial<Rect>; shapeKind: ShapeKind
```

- [ ] `execute.ts`: `case "action": return cmd.run({ context, selection, proxies, options });` (sin sync automático: cada acción hace el suyo). `withSlideSize` solo para geometry.

- [ ] `navigation.ts` (minShapes 0 en todos):
  - `slide-first/prev/next/last`: `readSlideNav` → `selectSlideByIndex` → `sync` → `{ changed: 0, message: "Diapositiva N de M" }`; en los extremos, mensaje "Ya estás en la primera/última diapositiva.".
  - `slide-goto`: usa `options.gotoSlide` (1-based); fuera de rango → `{ changed: 0, message: "La presentación tiene M diapositivas." }`.
- [ ] `expert.ts`:
  - `select-kind` (oculto; el panel lo dispara con `shapeKind`): `readSlideNav` → `readSlideShapes` → `filterByKind` → `selectShapes` → sync → `{ changed: n, message: n === 0 ? "No hay formas de ese tipo." : undefined }`.
  - `select-by-name`: igual con `filterByName(shapes, options.nameFilter)`; filtro vacío → mensaje "Escribe parte del nombre.".
  - `rename-numbered` (minShapes 1): `numberedNames(selection.shapes, options.renamePrefix || "Forma")` → `renameShapes` → sync.
  - `z-front/z-forward/z-backward/z-back` (minShapes 1): `zOrderApplyOrder` → `applyZOrder` → sync → `{ changed: n }`.
  - `group` (minShapes 2): `groupSelected` → sync → `{ changed: 1, message: "Agrupado." }`.
  - `ungroup` (minShapes 1): ids de `selection.shapes` con `type === "Group"`; si ninguno → `{ changed: 0, message: "No hay grupos en la selección." }`; si no `ungroupShapes` → sync.
  - `inspector-apply` (oculto; minShapes 1): `applyBoundingBox(selection.shapes, { ...boundingBox(shapes), ...options.inspector })` → `applyPatches` → sync.
- [ ] `registry.ts`: `SECTION_ORDER` += `"navigation", "expert"`; `hiddenCommandIds` += `select-kind`, `inspector-apply`.
- [ ] Tests (mock): slide-next desde índice 1 de 3 selecciona el id del índice 2 y devuelve "Diapositiva 3 de 3"; slide-next en la última devuelve el mensaje de extremo sin cambios; slide-goto fuera de rango; select-kind "image" selecciona solo imágenes; select-by-name vacío → mensaje; rename-numbered escribe nombres en orden de lectura; z-front aplica en orden ascendente de zOrder; group con 1 forma → `ok:false` por `minShapes`; ungroup sin grupos → mensaje; inspector-apply mueve la selección.

Commit: `feat(commands): comandos de navegación y expert con tipo action`.

---

### Task 4: Panel: secciones Navigation y Expert, fila Tema, inspector vivo

**Files:**
- Create: `src/taskpane/sections/navigationSection.ts`, `expertSection.ts`; Modify: `colorSection.ts`, `panel.ts`, `main.ts`, `icons.ts`, `sanitize.ts`, `settings.ts`; `tests/taskpane/settings.test.ts`

- [ ] `main.ts`: el refresco deja de ser solo un conteo. Nueva función `readSelectionSummary(): Promise<SelectionSummary | null>` en `src/office/summary.ts` con `SelectionSummary { count; slideId; slideIndex; slideCount; bbox: Rect | null }` en un `PowerPoint.run` con un sync (carga `getSelectedShapes().items/left,top,width,height`, `getSelectedSlides().items/id,index`, `slides.getCount()`). El panel recibe `setSummary(summary | null)` (que llama a `setSelectionCount(summary?.count ?? null)` y actualiza inspector/navegación). Cuando `slideId` cambia respecto al anterior, el panel pide `getThemeColors` (otro `PowerPoint.run`) y actualiza la fila Tema.
- [ ] `navigationSection.ts`: fila con botones de la sección + campo "Ir a" (`input type=number min=1`) con un botón "Ir" que ejecuta `slide-goto` con `{ gotoSlide }`; etiqueta "N / M" actualizada desde el summary.
- [ ] `expertSection.ts`: (a) `<select>` de tipo (Todas, Texto, Formas, Imágenes, Líneas, Tablas, Grupos, Otras) + botón "Seleccionar" → `select-kind` con `{ shapeKind }`; (b) campo "Nombre contiene" + botón → `select-by-name` con `{ nameFilter }`; (c) campo "Prefijo" (persistido en `renamePrefix`) junto al botón `rename-numbered`; (d) inspector: 4 inputs (X, Y, Ancho, Alto, en pt, un decimal) rellenados desde `summary.bbox` en cada refresco; al cambiar uno y pulsar Enter/`change`, ejecuta `inspector-apply` con `{ inspector: { [campo]: valor } }`; deshabilitados sin selección.
- [ ] `colorSection.ts`: fila "Tema" (subtítulo + swatches) visible solo cuando hay colores; método `setThemeColors(colors | null)`.
- [ ] `sanitize.ts`: `renamePrefix` persistido (string ≤ 40 chars, default "Forma"); `toCommandOptions` incluye `renamePrefix`, `nameFilter: ""`, `shapeKind: "all"`.
- [ ] Iconos: slideFirst, slidePrev, slideNext, slideLast, slideGoto, selectKind, selectByName, renameNumbered, zFront, zForward, zBackward, zBack, group, ungroup, inspector.
- [ ] Verificar `npm test && npm run typecheck && npm run build`.

Commit: `feat(taskpane): secciones Navigation y Expert, fila de colores del tema e inspector`.

---

### Task 5: README y versión 1.3.0

Sección "Comandos (Fase 4)"; nota: con PowerPoint reciente (API 1.10) el tamaño de diapositiva es exacto y aparece la fila "Tema" con los colores del tema; limitaciones: el inspector escala la selección de forma proporcional; "Seleccionar por tipo" actúa sobre la diapositiva activa. Prueba manual paso 8. Versiones 1.3.0 / 1.3.0.0. Commit `docs: README Fase 4 y versión 1.3.0`.

---

## Autorrevisión del plan

**Cobertura del spec § Fase 4:** ir a diapositiva N / anterior / siguiente ✔ (Task 3, 4); seleccionar por tipo o nombre ✔; renombrar con prefijo numerado ✔; z-order ✔; agrupar/desagrupar ✔; inspector editable ✔. Extras justificados: tamaño exacto y colores del tema con 1.10 (mejoran Fases 1 y 2 sin romper el fallback).

**Tipos:** `ActionCommand.run(ctx)` → `CommandOutcome` (runtime.ts); `SelectionSummary` (Task 4) consumido por `Panel.setSummary`; `ShapeKind`/`ZOrderOp` (Task 1) usados en Tasks 2–4; `applyBoundingBox` devuelve `Patch[]` compatible con `applyPatches`.
