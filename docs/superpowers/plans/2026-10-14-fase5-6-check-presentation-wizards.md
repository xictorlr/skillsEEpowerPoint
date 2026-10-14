# Fases 5 y 6: Check + Presentation + Wizards — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Cerrar el alcance del spec con las secciones Check (limpieza y diagnóstico de la presentación), Presentation (nueva presentación, exportar diapositivas seleccionadas) y Wizards (asistente de agenda).

**Architecture:** Igual que la Fase 4: comandos `action`. Lógica decidible sin Office en `src/core/check.ts` (detección fuera de diapositiva, fuentes distintas, layout de la agenda). Adaptador `src/office/deck.ts` para recorrer toda la presentación por fases (tipos primero, texto después), borrar formas, añadir diapositivas y cuadros de texto, exportar y crear presentaciones. Las operaciones que recorren todo el mazo cargan solo lo imprescindible y hacen como máximo 3 syncs de lectura.

**API:** `Slide.shapes` (1.3), `Shape.delete()` (1.3), `TextFrame.hasText` (1.4), `ShapeCollection.addTextBox/addGeometricShape` (1.4), `SlideCollection.add` (1.3), `Slide.moveTo` (1.8), `SlideCollection.exportAsBase64Presentation(ids)` (1.10, gated), `PowerPoint.createPresentation(base64?)` (1.1).

**No disponible en Office.js (documentar):** notas del orador, animaciones, corrector, impresión, correo.

---

## Mapa de archivos

| Archivo | Responsabilidad |
|---|---|
| `src/core/check.ts` | `isOffSlide`, `distinctFonts`, `agendaLayout`, `parseAgendaSections` |
| `src/office/deck.ts` | `readDeckShapes`, `loadHasText`, `deleteShapes`, `readFontNames`, `addSlideAfter`, `addTextBox`, `exportSlides`, `createPresentation` |
| `src/commands/check.ts`, `presentation.ts`, `wizards.ts` | Comandos `action` |
| `src/commands/types.ts`, `registry.ts` (modificar) | `SectionId` + `"check" | "presentation" | "wizards"`; `CommandOptions.agenda` |
| `src/taskpane/sections/wizardSection.ts` | Textarea de secciones + opciones + botón "Crear agenda" |
| `src/taskpane/panel.ts`, `sanitize.ts`, `icons.ts` (modificar) | Registro de la sección, persistencia del texto de agenda, iconos |
| `tests/core/check.test.ts`, `tests/office/deck.test.ts`, `tests/commands/check.test.ts`, `tests/taskpane/settings.test.ts` | Tests |
| `README.md`, `manifest.xml`, `package.json` | Docs y versión 1.4.0 |

---

### Task 1: Lógica pura (`core/check.ts`)

```ts
import type { Rect, ShapeModel, SlideSize } from "./types";

/** true si la forma está completamente fuera del lienzo (marco sin rotar). */
export function isOffSlide(shape: Rect, slide: SlideSize): boolean

/** Nombres de fuente únicos, ordenados alfabéticamente, sin nulos ni vacíos. */
export function distinctFonts(names: Array<string | null | undefined>): string[]

/** Separa el texto del textarea en secciones no vacías (una por línea, recortadas), máximo 20. */
export function parseAgendaSections(text: string): string[]

export interface AgendaItem { text: string; rect: Rect; highlighted: boolean }
export interface AgendaSlideSpec { title: { text: string; rect: Rect }; items: AgendaItem[] }
export interface AgendaOptions { title: string; sections: string[] }

/**
 * Layout de una diapositiva de agenda para `currentIndex` (0-based): título arriba, lista de secciones
 * en columna, la actual resaltada. Márgenes del 8 % lateral y 10 % superior; filas del 8 % de alto
 * hasta un máximo de 10 filas; si hay más secciones, dos columnas.
 */
export function agendaLayout(slide: SlideSize, opts: AgendaOptions, currentIndex: number): AgendaSlideSpec
```

Tests: `isOffSlide` (fuera por la derecha, parcialmente dentro → false, negativa completa → true); `distinctFonts` (dedup, orden, nulos); `parseAgendaSections` (líneas vacías, recorte, límite 20); `agendaLayout` (960×540, 3 secciones: título en `{76.8, 54, 806.4, 64.8}`, 3 filas de 43.2 de alto a partir de y=151.2, solo el índice actual `highlighted`; 12 secciones → dos columnas de 6).

Commit: `feat(core): comprobaciones de mazo y layout de agenda`.

---

### Task 2: Adaptador (`office/deck.ts`)

```ts
export interface DeckShape { slideId: string; slideIndex: number; shape: ShapeModel }
/** Un sync: carga slides/items/id y shapes/items/{id,name,type,left,top,width,height,rotation,zOrderPosition,level}. */
export async function readDeckShapes(context): Promise<{ shapes: DeckShape[]; proxies: Map<string, PowerPoint.Shape> }>
/** Encola textFrame.load("hasText") en las formas con texto (por tipo, como format.ts); el llamador hace sync. */
export function loadHasText(proxies: Map<string, PowerPoint.Shape>, ids: string[]): void
export function deleteShapes(proxies, ids: string[]): void      // shape.delete()
/** Encola font.load("name") en las formas con texto; el llamador hace sync; luego `readFontName(proxy)`. */
export function loadFontNames(proxies, ids: string[]): void
export function readFontName(shape: PowerPoint.Shape): string | null
/** Añade una diapositiva al final, sync, la mueve a `index` (moveTo), devuelve el Slide (sin sync final). */
export async function addSlideAfter(context, index: number): Promise<PowerPoint.Slide>
export function addTextBox(slide: PowerPoint.Slide, text: string, rect: Rect, style: { size: number; bold: boolean; color: string | null }): PowerPoint.Shape
/** API 1.10: exporta ids a base64 (un sync). Devuelve null si no hay soporte. */
export async function exportSlides(context, ids: string[]): Promise<string | null>
export function createPresentation(base64?: string): Promise<void>   // PowerPoint.createPresentation
```

Mock: `MockSlide.shapes.items` ya existe; añadir `shape.delete()` → `formatWrites {path:"delete"}`; `textFrame.hasText` configurable (`textProps.hasText`, default true); `slides.add()` crea un `MockSlide` nuevo al final (`ctx.addedSlides++`), `slide.moveTo(i)` registra; `slide.shapes.addTextBox(text, opts)` crea un `MockShape` y registra `{ path: "addTextBox", value: { text, ...opts } }`; `slides.exportAsBase64Presentation(ids)` → `{ value: "BASE64:" + ids.join(",") }`; `vi.stubGlobal("PowerPoint", { run, createPresentation: vi.fn() })` en tests.

Tests: `readDeckShapes` con 2 diapositivas → índices y proxies; `loadHasText` solo en tipos con texto; `deleteShapes` registra; `addSlideAfter(ctx, 1)` → add + moveTo(2); `exportSlides` sin 1.10 → null, con 1.10 → string.

Commit: `feat(office): recorrido del mazo, borrado, alta de diapositivas y exportación`.

---

### Task 3: Comandos

`types.ts`: `SectionId` + `"check" | "presentation" | "wizards"`; `SECTION_LABELS`: "Check", "Presentation", "Wizards"; `CommandOptions.agenda: { title: string; sectionsText: string }`.

`check.ts` (todos `action`, minShapes 0):
- `check-empty-text-slide`: en la diapositiva activa, cuadros de texto (`TextBox`) con `hasText === false` → `deleteShapes` → sync → `{ changed: n, message: n === 0 ? "No hay cuadros vacíos." : \`${n} cuadro(s) vacío(s) eliminado(s).\` }`.
- `check-empty-text-deck`: igual en todo el mazo (readDeckShapes → loadHasText → sync → delete → sync).
- `check-offslide-select`: formas de la diapositiva activa con `isOffSlide` → `selectShapes` → `{ changed: n, message: n === 0 ? "No hay formas fuera de la diapositiva." : \`${n} forma(s) fuera de la diapositiva seleccionada(s).\` }` (usa `getSlideSize`).
- `check-fonts`: `readDeckShapes` → `loadFontNames` en formas con texto → sync → `distinctFonts` → `{ changed: 0, message: \`Fuentes: ${list.join(", ")}\` }` (máx. 12 nombres, luego "…").

`presentation.ts`:
- `new-presentation`: `createPresentation()` → `{ changed: 0, message: "Nueva presentación creada." }`.
- `export-selected-slides`: ids de `getSelectedSlides()`; `exportSlides` → si null: `{ changed: 0, message: "Exportar requiere PowerPoint reciente (API 1.10)." }`; si no: `createPresentation(base64)` → `{ changed: n, message: \`${n} diapositiva(s) exportada(s) a una nueva presentación.\` }`.

`wizards.ts`:
- `agenda-create` (oculto; el panel lo dispara): `parseAgendaSections(options.agenda.sectionsText)`; si < 1 → "Escribe al menos una sección (una por línea)."; `readSlideNav` para el índice actual y `getSlideSize`; para cada sección i: `addSlideAfter(context, nav.index + 1 + i)`, `agendaLayout(size, opts, i)`, `addTextBox` del título (size 28, bold, color null) y de cada item (size 18, bold = highlighted, color = highlighted ? accent : "#7F7F7F" donde accent = primer color del tema Accent1 si `getThemeColors` lo da, si no "#1F6FEB"); sync al final → `{ changed: n, message: \`${n} diapositiva(s) de agenda creada(s).\` }`.

`registry.ts`: `SECTION_ORDER` += `"check", "presentation", "wizards"`; `hiddenCommandIds` += `agenda-create`.

Tests (mock): cada comando con casos de "nada que hacer" y "con cambios"; export sin 1.10; agenda con 3 secciones crea 3 slides con 4 cuadros cada una y resalta el correcto.

Commit: `feat(commands): secciones Check, Presentation y Wizards`.

---

### Task 4: Panel, README y versión

- `sections/wizardSection.ts`: campo "Título" (default "Agenda"), `<textarea rows=5>` "Secciones (una por línea)", botón "Crear agenda" → `agenda-create` con `{ agenda }`. Texto persistido (`sanitize`: título ≤ 60 chars, secciones ≤ 2000 chars).
- Secciones Check y Presentation: solo botones (render automático). Iconos: emptyTextSlide, emptyTextDeck, offSlide, fonts, newPresentation, exportSlides, agenda.
- Confirmación para acciones destructivas: `check-empty-text-deck` requiere pulsar dos veces en 5 s (el botón cambia su `title` a "Pulsa de nuevo para confirmar") — implementar en `Panel.run` para comandos con `confirm: true` en `BaseCommand`.
- README "Comandos (Fases 5 y 6)" + limitaciones (sin notas/animaciones; exportar requiere 1.10; la agenda usa cuadros de texto nativos) + prueba manual paso 9. Versiones 1.4.0 / 1.4.0.0.

Commits: `feat(taskpane): secciones Check, Presentation y asistente de agenda` y `docs: README Fases 5 y 6 y versión 1.4.0`.

---

## Autorrevisión del plan

Spec § Fase 5: eliminar notas → **no disponible** (documentado); formas vacías ✔; fuentes ✔; nueva presentación ✔; exportar seleccionadas ✔ (1.10). § Fase 6: agenda ✔. Tipos: `AgendaOptions`/`agendaLayout` (T1) ↔ `agenda-create` (T3) ↔ `CommandOptions.agenda` (T3/T4); `DeckShape` (T2) ↔ comandos check (T3).
