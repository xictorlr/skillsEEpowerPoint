# Fase 2: Format + Color — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Añadir al panel las secciones Format (copiar/pegar formato por partes) y Color (paleta, recientes, color personalizado, sin relleno/sin línea, transparencia) operando sobre la selección.

**Architecture:** Igual que la Fase 1: `src/core/format.ts` y `src/core/color.ts` son lógica pura sobre `ShapeFormat`/`FormatPatch`; `src/office/format.ts` lee y escribe `fill`, `lineFormat` y `textFrame.textRange.font`; `src/commands` gana dos tipos de comando (`format` y `copy-format`) y el ejecutor hace `switch (cmd.kind)`; el panel añade las dos secciones y persiste colores recientes. Las formas sin marco de texto (imagen, línea, tabla, gráfico, grupo…) se detectan por `type`, porque `shape.textFrame` lanza `InvalidArgument` en ellas y `getTextFrameOrNullObject` exige PowerPointApi 1.10.

**Tech Stack:** el existente (TypeScript, Vite 8, Vitest 5, Office.js PowerPointApi ≥ 1.8).

**Spec:** `docs/superpowers/specs/2026-10-14-slide-tools-design.md` § 5 Fase 2. Decisión de esta fase: el API no expone los colores del tema, así que la paleta es una lista fija y neutra (`DEFAULT_PALETTE`) más colores recientes y un selector de color personalizado.

---

## Mapa de archivos

| Archivo | Responsabilidad |
|---|---|
| `src/core/format.ts` | Tipos `FillFormat`, `LineFormat`, `FontFormat`, `ShapeFormat`, `FormatParts`, `FormatPatch`, `ColorTarget`; funciones puras `buildPastePatches`, `colorPatches`, `transparencyPatches`, `clearPatches` |
| `src/core/color.ts` | `normalizeHex`, `DEFAULT_PALETTE`, `pushRecent` |
| `src/office/format.ts` | `FORMAT_LOAD_PROPS`, `hasTextFrameByType`, `loadFonts`, `toShapeFormat`, `applyFormatPatch` |
| `src/office/selection.ts` (modificar) | `deps.withFormat` → carga formato (2 syncs) y devuelve `formats` |
| `src/office/runtime.ts` (modificar) | `fn` devuelve `CommandOutcome { changed, message? }`; `CommandResult` ok lleva `message?` |
| `src/commands/types.ts` (modificar) | `BaseCommand`, `GeometryCommand`, `FormatCommand`, `CopyFormatCommand`; `SectionId` + `"format" | "color"`; `CommandOptions` + `color`, `colorTarget`, `pasteParts` |
| `src/commands/clipboard.ts` | Portapapeles de formato en memoria (`getClipboard`, `setClipboard`, `resetClipboard`) |
| `src/commands/format.ts`, `src/commands/color.ts` | Comandos de las dos secciones |
| `src/commands/registry.ts` (modificar) | Nuevas secciones en `SECTION_ORDER`; `disabledReason` tiene en cuenta el portapapeles para "pegar" |
| `src/commands/execute.ts` (modificar) | `switch (cmd.kind)` |
| `src/taskpane/icons.ts` (modificar) | 13 iconos nuevos |
| `src/taskpane/sanitize.ts` (modificar) | `sanitizeRecentColors`, `PersistedOptions.recentColors` |
| `src/taskpane/panel.ts` (modificar) | Sección Color con selector de destino, paleta, recientes, `<input type="color">`; botones de Format |
| `src/taskpane/styles.css` (modificar) | `.st-swatch`, `.st-toggle`, `.st-toggle.is-active` |
| `tests/core/format.test.ts`, `tests/core/color.test.ts`, `tests/office/format.test.ts`, `tests/commands/format.test.ts`, `tests/taskpane/settings.test.ts` (ampliar) | Tests |

---

### Task 1: Tipos y lógica pura de formato (`core/format.ts`)

**Files:**
- Create: `src/core/format.ts`
- Create: `tests/core/format.test.ts`

- [ ] **Step 1: Escribir el test que falla**

```ts
import { describe, expect, it } from "vitest";
import {
  buildPastePatches,
  clearPatches,
  colorPatches,
  transparencyPatches,
  type ShapeFormat,
} from "../../src/core/format";

function fmt(id: string, over: Partial<ShapeFormat> = {}): ShapeFormat {
  return {
    id,
    fill: { type: "Solid", color: "#FF0000", transparency: 0 },
    line: { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 },
    font: { name: "Calibri", size: 12, color: "#111111", bold: false, italic: false, underline: "None" },
    size: { width: 100, height: 50 },
    ...over,
  };
}

const ALL = { fill: true, line: true, font: true, size: true };

describe("buildPastePatches", () => {
  it("copies every part from the source to each target, skipping the source itself", () => {
    const src = fmt("src");
    const t = fmt("t", { fill: { type: "NoFill", color: null, transparency: null } });
    expect(buildPastePatches(src, [src, t], ALL)).toEqual([
      {
        id: "t",
        fill: { type: "Solid", color: "#FF0000", transparency: 0 },
        line: { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 },
        font: { name: "Calibri", size: 12, color: "#111111", bold: false, italic: false, underline: "None" },
        width: 100,
        height: 50,
      },
    ]);
  });

  it("only includes the requested parts", () => {
    const src = fmt("src");
    const t = fmt("t");
    expect(buildPastePatches(src, [t], { fill: true, line: false, font: false, size: false })).toEqual([
      { id: "t", fill: { type: "Solid", color: "#FF0000", transparency: 0 } },
    ]);
  });

  it("omits null source values and never pastes font onto shapes without text", () => {
    const src = fmt("src", { fill: { type: "Solid", color: null, transparency: 0.5 } });
    const picture = fmt("pic", { font: null });
    expect(buildPastePatches(src, [picture], ALL)).toEqual([
      {
        id: "pic",
        fill: { type: "Solid", transparency: 0.5 },
        line: { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 },
        width: 100,
        height: 50,
      },
    ]);
  });

  it("returns [] when there is no target", () => {
    expect(buildPastePatches(fmt("src"), [fmt("src")], ALL)).toEqual([]);
  });
});

describe("colorPatches", () => {
  it("sets fill, line or font color on every shape", () => {
    const a = fmt("a");
    expect(colorPatches([a], "fill", "#00FF00")).toEqual([{ id: "a", fill: { type: "Solid", color: "#00FF00" } }]);
    expect(colorPatches([a], "line", "#00FF00")).toEqual([{ id: "a", line: { visible: true, color: "#00FF00" } }]);
    expect(colorPatches([a], "font", "#00FF00")).toEqual([{ id: "a", font: { color: "#00FF00" } }]);
  });

  it("skips font color for shapes without text", () => {
    expect(colorPatches([fmt("pic", { font: null })], "font", "#00FF00")).toEqual([]);
  });
});

describe("transparencyPatches", () => {
  it("steps transparency and clamps to [0, 1]", () => {
    const a = fmt("a", { fill: { type: "Solid", color: "#FF0000", transparency: 0.95 } });
    expect(transparencyPatches([a], "fill", 0.1)).toEqual([{ id: "a", fill: { transparency: 1 } }]);
    expect(transparencyPatches([a], "fill", -0.1)).toEqual([{ id: "a", fill: { transparency: 0.85 } }]);
  });

  it("treats null transparency as 0 and skips invisible lines", () => {
    const a = fmt("a", {
      fill: { type: "Solid", color: "#FF0000", transparency: null },
      line: { visible: false, color: null, weight: null, dashStyle: null, transparency: null },
    });
    expect(transparencyPatches([a], "fill", 0.2)).toEqual([{ id: "a", fill: { transparency: 0.2 } }]);
    expect(transparencyPatches([a], "line", 0.2)).toEqual([]);
  });

  it("omits shapes whose value would not change", () => {
    const a = fmt("a", { fill: { type: "Solid", color: "#FF0000", transparency: 1 } });
    expect(transparencyPatches([a], "fill", 0.1)).toEqual([]);
  });
});

describe("clearPatches", () => {
  it("removes fill or hides line, skipping shapes already cleared", () => {
    const a = fmt("a");
    const b = fmt("b", {
      fill: { type: "NoFill", color: null, transparency: null },
      line: { visible: false, color: null, weight: null, dashStyle: null, transparency: null },
    });
    expect(clearPatches([a, b], "fill")).toEqual([{ id: "a", fill: { type: "NoFill" } }]);
    expect(clearPatches([a, b], "line")).toEqual([{ id: "a", line: { visible: false } }]);
  });
});
```

- [ ] **Step 2: Ejecutar y ver que falla**

Run: `npx vitest run tests/core/format.test.ts`
Expected: FAIL, módulo inexistente.

- [ ] **Step 3: Implementar `src/core/format.ts`**

```ts
/** Formato de una forma tal como lo expone Office.js. `null` = valor mixto o no aplicable. */
export interface FillFormat {
  type: "Solid" | "NoFill" | "Other";
  color: string | null;
  transparency: number | null;
}

export interface LineFormat {
  visible: boolean | null;
  color: string | null;
  weight: number | null;
  dashStyle: string | null;
  transparency: number | null;
}

export interface FontFormat {
  name: string | null;
  size: number | null;
  color: string | null;
  bold: boolean | null;
  italic: boolean | null;
  underline: string | null;
}

export interface ShapeFormat {
  id: string;
  fill: FillFormat;
  line: LineFormat;
  /** null cuando la forma no tiene marco de texto (imagen, línea, tabla, gráfico…). */
  font: FontFormat | null;
  size: { width: number; height: number };
}

export interface FormatParts {
  fill: boolean;
  line: boolean;
  font: boolean;
  size: boolean;
}

export type ColorTarget = "fill" | "line" | "font";

/** Cambio parcial de formato. Solo se incluyen las partes y campos a escribir. */
export interface FormatPatch {
  id: string;
  fill?: Partial<FillFormat>;
  line?: Partial<LineFormat>;
  font?: Partial<FontFormat>;
  width?: number;
  height?: number;
}

const EPS = 0.001;

/** Copia solo los campos no nulos de `src`. Devuelve undefined si no queda ninguno. */
function nonNull<T extends object>(src: T): Partial<T> | undefined {
  const out: Partial<T> = {};
  let any = false;
  for (const key of Object.keys(src) as Array<keyof T>) {
    const v = src[key];
    if (v !== null && v !== undefined) {
      out[key] = v;
      any = true;
    }
  }
  return any ? out : undefined;
}

/** Pega las partes indicadas del formato `source` en cada `target` distinto de la fuente. */
export function buildPastePatches(source: ShapeFormat, targets: ShapeFormat[], parts: FormatParts): FormatPatch[] {
  const patches: FormatPatch[] = [];
  for (const t of targets) {
    if (t.id === source.id) continue;
    const patch: FormatPatch = { id: t.id };
    if (parts.fill) patch.fill = nonNull(source.fill);
    if (parts.line) patch.line = nonNull(source.line);
    if (parts.font && source.font && t.font) patch.font = nonNull(source.font);
    if (parts.size) {
      patch.width = source.size.width;
      patch.height = source.size.height;
    }
    if (patch.fill === undefined) delete patch.fill;
    if (patch.line === undefined) delete patch.line;
    if (patch.font === undefined) delete patch.font;
    if (Object.keys(patch).length > 1) patches.push(patch);
  }
  return patches;
}

/** Aplica un color sólido al relleno, la línea o el texto de cada forma. */
export function colorPatches(targets: ShapeFormat[], target: ColorTarget, color: string): FormatPatch[] {
  const patches: FormatPatch[] = [];
  for (const t of targets) {
    if (target === "fill") patches.push({ id: t.id, fill: { type: "Solid", color } });
    else if (target === "line") patches.push({ id: t.id, line: { visible: true, color } });
    else if (t.font) patches.push({ id: t.id, font: { color } });
  }
  return patches;
}

/** Suma `delta` a la transparencia (0..1) de relleno o línea. Líneas invisibles se omiten. */
export function transparencyPatches(targets: ShapeFormat[], target: "fill" | "line", delta: number): FormatPatch[] {
  const patches: FormatPatch[] = [];
  for (const t of targets) {
    if (target === "line" && t.line.visible === false) continue;
    const current = (target === "fill" ? t.fill.transparency : t.line.transparency) ?? 0;
    const next = Math.min(1, Math.max(0, Math.round((current + delta) * 100) / 100));
    if (Math.abs(next - current) < EPS) continue;
    patches.push(target === "fill" ? { id: t.id, fill: { transparency: next } } : { id: t.id, line: { transparency: next } });
  }
  return patches;
}

/** Quita el relleno o la línea de cada forma que aún los tenga. */
export function clearPatches(targets: ShapeFormat[], target: "fill" | "line"): FormatPatch[] {
  const patches: FormatPatch[] = [];
  for (const t of targets) {
    if (target === "fill" && t.fill.type !== "NoFill") patches.push({ id: t.id, fill: { type: "NoFill" } });
    if (target === "line" && t.line.visible !== false) patches.push({ id: t.id, line: { visible: false } });
  }
  return patches;
}
```

- [ ] **Step 4: Ejecutar y ver que pasa**

Run: `npx vitest run tests/core/format.test.ts`
Expected: `11 passed`.

- [ ] **Step 5: Commit**

```bash
git add src/core/format.ts tests/core/format.test.ts
git commit -m "feat(core): modelo de formato y parches de pegar, color, transparencia y limpiar"
```

---

### Task 2: Colores (`core/color.ts`)

**Files:**
- Create: `src/core/color.ts`
- Create: `tests/core/color.test.ts`

- [ ] **Step 1: Escribir el test que falla**

```ts
import { describe, expect, it } from "vitest";
import { DEFAULT_PALETTE, normalizeHex, pushRecent } from "../../src/core/color";

describe("normalizeHex", () => {
  it("accepts 6-digit hex with or without #, upper-casing it", () => {
    expect(normalizeHex("#ffa500")).toBe("#FFA500");
    expect(normalizeHex("ffa500")).toBe("#FFA500");
  });
  it("expands 3-digit hex", () => {
    expect(normalizeHex("#abc")).toBe("#AABBCC");
  });
  it("rejects anything else", () => {
    expect(normalizeHex("orange")).toBeNull();
    expect(normalizeHex("#12345")).toBeNull();
    expect(normalizeHex("")).toBeNull();
  });
});

describe("DEFAULT_PALETTE", () => {
  it("has 16 unique normalized colors", () => {
    expect(DEFAULT_PALETTE).toHaveLength(16);
    expect(new Set(DEFAULT_PALETTE).size).toBe(16);
    expect(DEFAULT_PALETTE.every((c) => normalizeHex(c) === c)).toBe(true);
  });
});

describe("pushRecent", () => {
  it("prepends, de-duplicates and caps the list", () => {
    expect(pushRecent(["#111111", "#222222"], "#222222", 8)).toEqual(["#222222", "#111111"]);
    const eight = ["#000001", "#000002", "#000003", "#000004", "#000005", "#000006", "#000007", "#000008"];
    expect(pushRecent(eight, "#000009", 8)).toEqual(["#000009", ...eight.slice(0, 7)]);
  });
  it("ignores invalid colors", () => {
    expect(pushRecent(["#111111"], "nope", 8)).toEqual(["#111111"]);
  });
});
```

- [ ] **Step 2: Ejecutar y ver que falla**

Run: `npx vitest run tests/core/color.test.ts`
Expected: FAIL.

- [ ] **Step 3: Implementar `src/core/color.ts`**

```ts
/** Normaliza a "#RRGGBB" en mayúsculas. Acepta "abc", "#abc", "aabbcc", "#AABBCC". */
export function normalizeHex(input: string): string | null {
  const s = input.trim().replace(/^#/, "");
  if (/^[0-9a-fA-F]{6}$/.test(s)) return `#${s.toUpperCase()}`;
  if (/^[0-9a-fA-F]{3}$/.test(s)) {
    const [r, g, b] = s.toUpperCase();
    return `#${r}${r}${g}${g}${b}${b}`;
  }
  return null;
}

/** Paleta neutra: grises, azules, verdes, ámbar, rojo. Sin colores de marca. */
export const DEFAULT_PALETTE: readonly string[] = [
  "#000000", "#404040", "#808080", "#BFBFBF", "#D9D9D9", "#FFFFFF",
  "#0B3D91", "#1F6FEB", "#7FB3FF",
  "#0F6E56", "#2EA043", "#9BE9A8",
  "#B35900", "#F2A900",
  "#B42318", "#7A1FA2",
];

/** Inserta `color` al principio de `recent` sin duplicados, limitado a `max` entradas. */
export function pushRecent(recent: readonly string[], color: string, max: number): string[] {
  const hex = normalizeHex(color);
  if (!hex) return [...recent];
  return [hex, ...recent.filter((c) => c !== hex)].slice(0, max);
}
```

- [ ] **Step 4: Ejecutar y ver que pasa**

Run: `npx vitest run tests/core/color.test.ts`
Expected: `6 passed`.

- [ ] **Step 5: Commit**

```bash
git add src/core/color.ts tests/core/color.test.ts
git commit -m "feat(core): normalización de colores, paleta neutra y recientes"
```

---

### Task 3: Adaptador Office de formato y lectura de selección con formato

**Files:**
- Create: `src/office/format.ts`
- Modify: `src/office/selection.ts`
- Modify: `src/office/runtime.ts`
- Modify: `tests/office/mockPowerPoint.ts`
- Create: `tests/office/format.test.ts`
- Modify: `tests/office/selection.test.ts`, `tests/office/runtime.test.ts`

- [ ] **Step 1: Ampliar el mock `tests/office/mockPowerPoint.ts`**

Añadir a `MockShapeInit`: `fill?: { type: string; foregroundColor: string; transparency: number | null }`, `line?: { visible: boolean | null; color: string | null; weight: number | null; dashStyle: string | null; transparency: number | null }`, `font?: { name; size; color; bold; italic; underline } | null` (null = sin marco de texto). Añadir a `MockShape`:

```ts
  fill: {
    type: string; foregroundColor: string; transparency: number | null;
    load: () => void; setSolidColor: (c: string) => void; clear: () => void;
  };
  lineFormat: { visible: boolean | null; color: string | null; weight: number | null; dashStyle: string | null; transparency: number | null; load: () => void };
  textFrame: { textRange: { font: Record<string, unknown> & { load: (p: string) => void } } };
  formatWrites: Array<{ path: string; value: unknown }> = [];
  fontLoaded = false;
```

En el constructor: inicializar `fill` desde `init.fill ?? { type: "Solid", foregroundColor: "#FF0000", transparency: 0 }`, con `setSolidColor: (c) => { this.fill.type = "Solid"; this.fill.foregroundColor = c; this.formatWrites.push({ path: "fill.setSolidColor", value: c }); }` y `clear: () => { this.fill.type = "NoFill"; this.formatWrites.push({ path: "fill.clear", value: null }); }`; convertir `fill.transparency`, cada campo de `lineFormat` y cada campo de `font` en accessors que registran `{ path: "fill.transparency", value }`, `{ path: "lineFormat.color", value }`, `{ path: "font.size", value }`… (mismo patrón `Object.defineProperty` que ya usa para left/top). Si `init.font === null`, `textFrame` debe ser un getter que lanza `new Error("InvalidArgument")` (imita a Office.js). `font.load = () => { this.fontLoaded = true; }`.

- [ ] **Step 2: Escribir el test que falla `tests/office/format.test.ts`**

```ts
import { describe, expect, it } from "vitest";
import { applyFormatPatch, hasTextFrameByType, loadFonts, toShapeFormat } from "../../src/office/format";
import { MockShape } from "./mockPowerPoint";

describe("hasTextFrameByType", () => {
  it("is true for shapes that can hold text and false otherwise", () => {
    for (const t of ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"]) expect(hasTextFrameByType(t)).toBe(true);
    for (const t of ["Image", "Line", "Table", "Chart", "Group", "SmartArt", "Media", "Unsupported"]) expect(hasTextFrameByType(t)).toBe(false);
  });
});

describe("loadFonts / toShapeFormat", () => {
  it("loads fonts only on text shapes and maps the format, normalizing colors", () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20,
      fill: { type: "Solid", foregroundColor: "ffa500", transparency: 0.25 },
      font: { name: "Arial", size: 14, color: "#112233", bold: true, italic: null, underline: "None" } });
    const pic = new MockShape({ id: "p", type: "Image", left: 0, top: 0, width: 10, height: 20, font: null });
    loadFonts([box, pic] as never);
    expect(box.fontLoaded).toBe(true);
    expect(toShapeFormat(box as never)).toEqual({
      id: "a",
      fill: { type: "Solid", color: "#FFA500", transparency: 0.25 },
      line: { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 },
      font: { name: "Arial", size: 14, color: "#112233", bold: true, italic: null, underline: "None" },
      size: { width: 10, height: 20 },
    });
    expect(toShapeFormat(pic as never).font).toBeNull();
  });

  it("maps unknown fill types to Other and NoFill to NoFill", () => {
    const g = new MockShape({ id: "g", left: 0, top: 0, width: 1, height: 1, fill: { type: "Gradient", foregroundColor: "#000000", transparency: null } });
    const n = new MockShape({ id: "n", left: 0, top: 0, width: 1, height: 1, fill: { type: "NoFill", foregroundColor: "#000000", transparency: null } });
    expect(toShapeFormat(g as never).fill.type).toBe("Other");
    expect(toShapeFormat(n as never).fill.type).toBe("NoFill");
  });
});

describe("applyFormatPatch", () => {
  it("writes solid color via setSolidColor, NoFill via clear, and other fields directly", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20 });
    applyFormatPatch(s as never, { id: "a", fill: { type: "Solid", color: "#00FF00", transparency: 0.5 }, line: { visible: true, color: "#0000FF", weight: 2 }, font: { size: 18, bold: true }, width: 30 });
    expect(s.formatWrites).toEqual([
      { path: "fill.setSolidColor", value: "#00FF00" },
      { path: "fill.transparency", value: 0.5 },
      { path: "lineFormat.visible", value: true },
      { path: "lineFormat.color", value: "#0000FF" },
      { path: "lineFormat.weight", value: 2 },
      { path: "font.size", value: 18 },
      { path: "font.bold", value: true },
    ]);
    expect(s.writes).toEqual([{ prop: "width", value: 30 }]);
  });

  it("clears the fill for NoFill and never touches font on shapes without text", () => {
    const pic = new MockShape({ id: "p", type: "Image", left: 0, top: 0, width: 1, height: 1, font: null });
    applyFormatPatch(pic as never, { id: "p", fill: { type: "NoFill" }, font: { size: 10 } });
    expect(pic.formatWrites).toEqual([{ path: "fill.clear", value: null }]);
  });
});
```

- [ ] **Step 3: Ejecutar y ver que falla**

Run: `npx vitest run tests/office/format.test.ts`
Expected: FAIL.

- [ ] **Step 4: Implementar `src/office/format.ts`**

```ts
import { normalizeHex } from "../core/color";
import type { FillFormat, FormatPatch, LineFormat, ShapeFormat } from "../core/format";

/** Propiedades de formato que se cargan junto a la geometría (válidas para cualquier tipo de forma). */
export const FORMAT_LOAD_PROPS =
  "items/fill/type,items/fill/foregroundColor,items/fill/transparency," +
  "items/lineFormat/visible,items/lineFormat/color,items/lineFormat/weight,items/lineFormat/dashStyle,items/lineFormat/transparency";

const FONT_PROPS = "name,size,color,bold,italic,underline";
const TEXT_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"]);

/** `shape.textFrame` lanza InvalidArgument en formas sin texto; se decide por tipo. */
export function hasTextFrameByType(type: string): boolean {
  return TEXT_TYPES.has(type);
}

/** Encola la carga de la fuente en las formas con texto. Requiere un sync posterior. */
export function loadFonts(shapes: PowerPoint.Shape[]): void {
  for (const s of shapes) {
    if (hasTextFrameByType(String(s.type))) s.textFrame.textRange.font.load(FONT_PROPS);
  }
}

function color(v: string | null | undefined): string | null {
  return v == null ? null : normalizeHex(v);
}

export function toShapeFormat(shape: PowerPoint.Shape): ShapeFormat {
  const fillType = String(shape.fill.type);
  const fill: FillFormat = {
    type: fillType === "Solid" ? "Solid" : fillType === "NoFill" ? "NoFill" : "Other",
    color: color(shape.fill.foregroundColor),
    transparency: shape.fill.transparency ?? null,
  };
  const lf = shape.lineFormat;
  const line: LineFormat = {
    visible: lf.visible ?? null,
    color: color(lf.color),
    weight: lf.weight ?? null,
    dashStyle: lf.dashStyle == null ? null : String(lf.dashStyle),
    transparency: lf.transparency ?? null,
  };
  let font: ShapeFormat["font"] = null;
  if (hasTextFrameByType(String(shape.type))) {
    const f = shape.textFrame.textRange.font;
    font = {
      name: f.name ?? null,
      size: f.size ?? null,
      color: color(f.color),
      bold: f.bold ?? null,
      italic: f.italic ?? null,
      underline: f.underline == null ? null : String(f.underline),
    };
  }
  return { id: shape.id, fill, line, font, size: { width: shape.width, height: shape.height } };
}

/** Escribe un parche de formato. No hace sync. */
export function applyFormatPatch(shape: PowerPoint.Shape, patch: FormatPatch): void {
  if (patch.fill) {
    if (patch.fill.type === "NoFill") shape.fill.clear();
    else if (patch.fill.color) shape.fill.setSolidColor(patch.fill.color);
    if (patch.fill.type !== "NoFill" && patch.fill.transparency != null) shape.fill.transparency = patch.fill.transparency;
  }
  if (patch.line) {
    const l = patch.line;
    if (l.visible != null) shape.lineFormat.visible = l.visible;
    if (l.color) shape.lineFormat.color = l.color;
    if (l.weight != null) shape.lineFormat.weight = l.weight;
    if (l.dashStyle) shape.lineFormat.dashStyle = l.dashStyle as PowerPoint.ShapeLineDashStyle;
    if (l.transparency != null) shape.lineFormat.transparency = l.transparency;
  }
  if (patch.font && hasTextFrameByType(String(shape.type))) {
    const f = shape.textFrame.textRange.font;
    const p = patch.font;
    if (p.name) f.name = p.name;
    if (p.size != null) f.size = p.size;
    if (p.color) f.color = p.color;
    if (p.bold != null) f.bold = p.bold;
    if (p.italic != null) f.italic = p.italic;
    if (p.underline) f.underline = p.underline as PowerPoint.ShapeFontUnderlineStyle;
  }
  if (patch.width != null && Number.isFinite(patch.width)) shape.width = patch.width;
  if (patch.height != null && Number.isFinite(patch.height)) shape.height = patch.height;
}
```

Nota sobre el test de `applyFormatPatch`: el orden de `formatWrites` esperado es exactamente el orden de escritura de esta implementación.

- [ ] **Step 5: Ejecutar y ver que pasa**

Run: `npx vitest run tests/office/format.test.ts`
Expected: `5 passed`.

- [ ] **Step 6: Modificar `src/office/selection.ts`**

`ReadSelectionDeps` gana `withFormat?: boolean`. `SelectionRead` gana `formats: ShapeFormat[]` (vacío si `withFormat` es falso). En `readSelection`: el string de carga es `SHAPE_PROPS` o `${SHAPE_PROPS},${FORMAT_LOAD_PROPS}` según `withFormat`; tras el primer sync, si `withFormat`, `loadFonts(collection.items)` y `await context.sync()`; `formats = collection.items.map(toShapeFormat)`. Doc: "con formato hace dos syncs de lectura (el segundo para las fuentes de las formas con texto)".

Añadir a `tests/office/selection.test.ts` (dentro del `describe("readSelection")` existente):

```ts
  it("loads format props and fonts when withFormat is set (two read syncs) and returns formats", async () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20,
      font: { name: "Arial", size: 14, color: "#112233", bold: true, italic: null, underline: "None" } });
    const ctx = new MockContext([box], "bad-png");
    const { formats } = await readSelection(ctx as never, { slideSizeCache: new Map(), slideSizeOverride: { width: 960, height: 540 }, withFormat: true });
    expect(ctx.loadedProps).toContain("items/fill/type");
    expect(ctx.loadedProps).toContain("items/lineFormat/weight");
    expect(box.fontLoaded).toBe(true);
    expect(ctx.syncCount).toBe(2);
    expect(formats[0]?.font?.name).toBe("Arial");
  });

  it("does not load format props by default", async () => {
    const ctx = new MockContext([new MockShape({ id: "a", left: 0, top: 0, width: 1, height: 1 })], "bad-png");
    const { formats } = await readSelection(ctx as never, { slideSizeCache: new Map(), slideSizeOverride: { width: 960, height: 540 } });
    expect(ctx.loadedProps).not.toContain("items/fill/type");
    expect(formats).toEqual([]);
  });
```

- [ ] **Step 7: Modificar `src/office/runtime.ts`**

```ts
export interface CommandOutcome {
  changed: number;
  /** Mensaje informativo opcional para el panel (p. ej. "Formato copiado"). */
  message?: string;
}

export type CommandResult = ({ ok: true } & CommandOutcome) | { ok: false; message: string };

export async function runCommand(
  commandId: string,
  fn: (context: PowerPoint.RequestContext) => Promise<CommandOutcome>,
): Promise<CommandResult> {
  try {
    const outcome = await PowerPoint.run(fn);
    return { ok: true, ...outcome };
  } catch (err) { /* sin cambios */ }
}
```

Actualizar `tests/office/runtime.test.ts`: la función de éxito devuelve `{ changed: 2 }` y se espera `{ ok: true, changed: 2 }`; añadir un caso con `{ changed: 0, message: "Formato copiado" }` → `{ ok: true, changed: 0, message: "Formato copiado" }`. `src/commands/execute.ts` deja de compilar hasta la Task 5; para que `npm run typecheck` pase en este commit, cambiar en `execute.ts` el `return 0` por `return { changed: 0 }` y el `return new Set(...).size` por `return { changed: new Set(...).size }` (cambio mínimo; el `switch` llega en la Task 5).

- [ ] **Step 8: Ejecutar todo**

Run: `npm test && npm run typecheck`
Expected: todo en verde (los tests de `tests/commands/execute.test.ts` siguen pasando porque comparan `{ ok: true, changed: n }`).

- [ ] **Step 9: Commit**

```bash
git add src/office tests/office src/commands/execute.ts
git commit -m "feat(office): lectura y escritura de formato, selección con formato y resultado con mensaje"
```

---

### Task 4: Tipos de comando, portapapeles y comandos de Format y Color

**Files:**
- Modify: `src/commands/types.ts`
- Create: `src/commands/clipboard.ts`
- Create: `src/commands/format.ts`
- Create: `src/commands/color.ts`
- Modify: `src/commands/registry.ts`
- Modify: `tests/commands/registry.test.ts`

- [ ] **Step 1: Reescribir `src/commands/types.ts`**

```ts
import type { MatrixOptions } from "../core/geometry";
import type { ColorTarget, FormatParts, FormatPatch, ShapeFormat } from "../core/format";
import type { Patch, SelectionContext, SlideSize } from "../core/types";

export type SectionId = "alignment" | "size" | "format" | "color";

export const SECTION_LABELS: Readonly<Record<SectionId, string>> = {
  alignment: "Alignment",
  size: "Size",
  format: "Format",
  color: "Color",
};

export interface CommandOptions {
  matrix: MatrixOptions;
  /** Tamaño de diapositiva forzado por el usuario; si falta se deduce automáticamente. */
  slideSize?: SlideSize;
  /** Color elegido en la paleta ("#RRGGBB"); lo usan los comandos de color. */
  color?: string;
  /** A qué se aplica el color: relleno, línea o texto. */
  colorTarget: ColorTarget;
  /** Qué partes pega "Pegar formato". */
  pasteParts: FormatParts;
}

interface BaseCommand {
  id: string;
  section: SectionId;
  label: string;
  /** id en src/taskpane/icons.ts */
  icon: string;
  minShapes: number;
  maxShapes?: number;
}

export interface GeometryCommand extends BaseCommand {
  kind: "geometry";
  compute: (selection: SelectionContext, options: CommandOptions) => Patch[];
}

export interface FormatCommand extends BaseCommand {
  kind: "format";
  /** Si es true el botón se deshabilita mientras no haya formato copiado. */
  needsClipboard?: boolean;
  compute: (formats: ShapeFormat[], options: CommandOptions, clipboard: ShapeFormat | null) => FormatPatch[];
}

/** Guarda el formato de la primera forma seleccionada en el portapapeles interno. */
export interface CopyFormatCommand extends BaseCommand {
  kind: "copy-format";
}

export type Command = GeometryCommand | FormatCommand | CopyFormatCommand;
```

- [ ] **Step 2: Crear `src/commands/clipboard.ts`**

```ts
import type { ShapeFormat } from "../core/format";

let clipboard: ShapeFormat | null = null;

export function getClipboard(): ShapeFormat | null {
  return clipboard;
}

export function setClipboard(format: ShapeFormat | null): void {
  clipboard = format;
}

/** Solo para tests. */
export function resetClipboard(): void {
  clipboard = null;
}
```

- [ ] **Step 3: Crear `src/commands/format.ts`**

```ts
import { buildPastePatches } from "../core/format";
import type { Command } from "./types";

const paste = (id: string, label: string, icon: string, parts: { fill: boolean; line: boolean; font: boolean; size: boolean } | "options"): Command => ({
  kind: "format",
  id,
  section: "format",
  label,
  icon,
  minShapes: 1,
  needsClipboard: true,
  compute: (formats, options, clipboard) =>
    clipboard ? buildPastePatches(clipboard, formats, parts === "options" ? options.pasteParts : parts) : [],
});

export const formatCommands: readonly Command[] = [
  { kind: "copy-format", id: "copy-format", section: "format", label: "Copiar formato de la primera forma", icon: "copyFormat", minShapes: 1 },
  paste("paste-format", "Pegar formato (según partes elegidas)", "pasteFormat", "options"),
  paste("paste-fill", "Pegar solo relleno", "pasteFill", { fill: true, line: false, font: false, size: false }),
  paste("paste-line", "Pegar solo línea", "pasteLine", { fill: false, line: true, font: false, size: false }),
  paste("paste-font", "Pegar solo fuente", "pasteFont", { fill: false, line: false, font: true, size: false }),
  paste("paste-size", "Pegar solo tamaño", "pasteSize", { fill: false, line: false, font: false, size: true }),
];
```

- [ ] **Step 4: Crear `src/commands/color.ts`**

```ts
import { clearPatches, colorPatches, transparencyPatches } from "../core/format";
import type { Command } from "./types";

export const TRANSPARENCY_STEP = 0.1;

export const colorCommands: readonly Command[] = [
  {
    kind: "format", id: "set-color", section: "color", label: "Aplicar color", icon: "setColor", minShapes: 1,
    compute: (formats, options) => (options.color ? colorPatches(formats, options.colorTarget, options.color) : []),
  },
  { kind: "format", id: "no-fill", section: "color", label: "Sin relleno", icon: "noFill", minShapes: 1, compute: (f) => clearPatches(f, "fill") },
  { kind: "format", id: "no-line", section: "color", label: "Sin línea", icon: "noLine", minShapes: 1, compute: (f) => clearPatches(f, "line") },
  {
    kind: "format", id: "transparency-up", section: "color", label: "Más transparente (+10 %)", icon: "transparencyUp", minShapes: 1,
    compute: (f, o) => (o.colorTarget === "font" ? [] : transparencyPatches(f, o.colorTarget, TRANSPARENCY_STEP)),
  },
  {
    kind: "format", id: "transparency-down", section: "color", label: "Menos transparente (−10 %)", icon: "transparencyDown", minShapes: 1,
    compute: (f, o) => (o.colorTarget === "font" ? [] : transparencyPatches(f, o.colorTarget, -TRANSPARENCY_STEP)),
  },
];
```

El comando `set-color` no se muestra como botón: el panel lo ejecuta al pulsar una muestra de color (ver Task 6). Para ello el registro expone `hiddenCommandIds`.

- [ ] **Step 5: Modificar `src/commands/registry.ts`**

```ts
import { alignmentCommands } from "./alignment";
import { colorCommands } from "./color";
import { formatCommands } from "./format";
import { sizeCommands } from "./size";
import type { Command, SectionId } from "./types";

export const SECTION_ORDER: readonly SectionId[] = ["alignment", "size", "format", "color"];

export const commands: readonly Command[] = [...alignmentCommands, ...sizeCommands, ...formatCommands, ...colorCommands];

/** Comandos que el panel no renderiza como botón (se disparan desde otros controles). */
export const hiddenCommandIds: ReadonlySet<string> = new Set(["set-color"]);

export function commandsBySection(section: SectionId): Command[] {
  return commands.filter((c) => c.section === section && !hiddenCommandIds.has(c.id));
}

export function findCommand(id: string): Command | undefined {
  return commands.find((c) => c.id === id);
}

function formas(n: number): string {
  return `${n} forma${n === 1 ? "" : "s"}`;
}

/** Motivo por el que un comando no aplica, o undefined si aplica. `hasClipboard` solo importa para pegar formato. */
export function disabledReason(cmd: Command, count: number, hasClipboard = true): string | undefined {
  if (count < cmd.minShapes) return `Selecciona al menos ${formas(cmd.minShapes)}.`;
  if (cmd.maxShapes !== undefined && count > cmd.maxShapes) return `Selecciona como máximo ${formas(cmd.maxShapes)}.`;
  if (cmd.kind === "format" && cmd.needsClipboard && !hasClipboard) return "Primero copia el formato de una forma.";
  return undefined;
}
```

- [ ] **Step 6: Ampliar `tests/commands/registry.test.ts`**

```ts
  it("includes format and color sections and hides set-color from buttons", () => {
    expect(SECTION_ORDER).toEqual(["alignment", "size", "format", "color"]);
    expect(commandsBySection("format").map((c) => c.id)).toEqual(["copy-format", "paste-format", "paste-fill", "paste-line", "paste-font", "paste-size"]);
    expect(commandsBySection("color").map((c) => c.id)).toEqual(["no-fill", "no-line", "transparency-up", "transparency-down"]);
    expect(findCommand("set-color")).toBeDefined();
  });

  it("disables paste commands without a clipboard", () => {
    const paste = findCommand("paste-format")!;
    expect(disabledReason(paste, 1, false)).toMatch(/copia el formato/);
    expect(disabledReason(paste, 1, true)).toBeUndefined();
    expect(disabledReason(findCommand("copy-format")!, 1, false)).toBeUndefined();
  });
```

(Importar `SECTION_ORDER` y `commandsBySection` del registro.) Actualizar el test existente "align-left computes patches via core" para que pase las opciones completas: `{ matrix: {...}, colorTarget: "fill", pasteParts: { fill: true, line: true, font: true, size: true } }`. Añadir un helper `const OPTS: CommandOptions = {...}` al principio del archivo y reutilizarlo. Hacer lo mismo en `tests/commands/execute.test.ts` (`opts`).

- [ ] **Step 7: Ejecutar**

Run: `npx vitest run tests/commands/registry.test.ts && npm run typecheck`
Expected: tests en verde. El typecheck fallará en `execute.ts` (`cmd.compute` ya no existe en todas las variantes): se arregla en la Task 5. Si prefieres commit verde, haz la Task 5 antes de commitear; en ese caso un único commit para Tasks 4+5 con el mensaje de la Task 5.

- [ ] **Step 8: Commit (si typecheck pasa) o continuar a Task 5**

```bash
git add src/commands tests/commands
git commit -m "feat(commands): tipos de comando de formato, portapapeles y comandos Format y Color"
```

---

### Task 5: Ejecutor con `switch (cmd.kind)` y tests

**Files:**
- Modify: `src/commands/execute.ts`
- Create: `tests/commands/format.test.ts`

- [ ] **Step 1: Reescribir `src/commands/execute.ts`**

```ts
import { applyPatches } from "../office/apply";
import { applyFormatPatch } from "../office/format";
import { runCommand, type CommandOutcome, type CommandResult } from "../office/runtime";
import { readSelection } from "../office/selection";
import type { SlideSize } from "../core/types";
import { getClipboard, setClipboard } from "./clipboard";
import { disabledReason } from "./registry";
import type { Command, CommandOptions } from "./types";

// La caché no se invalida si el usuario cambia Diseño > Tamaño de diapositiva con el panel abierto
// (limitación conocida; la opción slideSize es la alternativa).
const slideSizeCache = new Map<string, SlideSize>();

/** Solo para tests. */
export function resetSlideSizeCache(): void {
  slideSizeCache.clear();
}

/** Lee la selección, calcula los parches del comando y los aplica en un único sync de escritura. */
export async function executeCommand(cmd: Command, options: CommandOptions): Promise<CommandResult> {
  return runCommand(cmd.id, async (context): Promise<CommandOutcome> => {
    const withFormat = cmd.kind !== "geometry";
    const { selection, proxies, formats } = await readSelection(context, {
      slideSizeCache,
      slideSizeOverride: options.slideSize,
      withFormat,
    });
    const reason = disabledReason(cmd, selection.shapes.length, getClipboard() !== null);
    if (reason) throw new Error(reason);

    switch (cmd.kind) {
      case "geometry": {
        const patches = cmd.compute(selection, options);
        if (patches.length === 0) return { changed: 0 };
        applyPatches(proxies, patches);
        await context.sync();
        return { changed: new Set(patches.map((p) => p.id)).size };
      }
      case "copy-format": {
        const first = formats[0];
        if (!first) return { changed: 0 };
        setClipboard(first);
        return { changed: 0, message: "Formato copiado." };
      }
      case "format": {
        const patches = cmd.compute(formats, options, getClipboard());
        if (patches.length === 0) return { changed: 0 };
        for (const p of patches) {
          const shape = proxies.get(p.id);
          if (shape) applyFormatPatch(shape, p);
        }
        await context.sync();
        return { changed: new Set(patches.map((p) => p.id)).size };
      }
    }
  });
}
```

- [ ] **Step 2: Escribir `tests/commands/format.test.ts`**

```ts
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { getClipboard, resetClipboard } from "../../src/commands/clipboard";
import { executeCommand, resetSlideSizeCache } from "../../src/commands/execute";
import { findCommand } from "../../src/commands/registry";
import type { CommandOptions } from "../../src/commands/types";
import { MockContext, MockShape } from "../office/mockPowerPoint";

const OPTS: CommandOptions = {
  matrix: { columns: 2, gapX: 8, gapY: 8 },
  slideSize: { width: 960, height: 540 },
  colorTarget: "fill",
  pasteParts: { fill: true, line: true, font: true, size: true },
};

let ctx: MockContext;

beforeEach(() => {
  resetClipboard();
  resetSlideSizeCache();
  vi.spyOn(console, "error").mockImplementation(() => {});
  vi.stubGlobal("PowerPoint", { run: (fn: (c: unknown) => Promise<unknown>) => fn(ctx) });
});
afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const box = (id: string, color: string) =>
  new MockShape({ id, type: "TextBox", left: 0, top: 0, width: 10, height: 10,
    fill: { type: "Solid", foregroundColor: color, transparency: 0 },
    font: { name: "Arial", size: 12, color: "#000000", bold: false, italic: false, underline: "None" } });

describe("copy-format / paste-format", () => {
  it("copies the first shape's format and pastes it onto the others with one write sync", async () => {
    const a = box("a", "#FF0000");
    ctx = new MockContext([a], "bad-png");
    const copied = await executeCommand(findCommand("copy-format")!, OPTS);
    expect(copied).toEqual({ ok: true, changed: 0, message: "Formato copiado." });
    expect(getClipboard()?.fill.color).toBe("#FF0000");
    expect(ctx.syncCount).toBe(2); // geometría+formato, fuentes

    const b = box("b", "#00FF00");
    ctx = new MockContext([b], "bad-png");
    const pasted = await executeCommand(findCommand("paste-fill")!, OPTS);
    expect(pasted).toEqual({ ok: true, changed: 1 });
    expect(b.formatWrites).toEqual([{ path: "fill.setSolidColor", value: "#FF0000" }, { path: "fill.transparency", value: 0 }]);
    expect(ctx.syncCount).toBe(3);
  });

  it("refuses to paste without a clipboard", async () => {
    ctx = new MockContext([box("a", "#FF0000")], "bad-png");
    const r = await executeCommand(findCommand("paste-format")!, OPTS);
    expect(r).toEqual({ ok: false, message: "Primero copia el formato de una forma." });
  });
});

describe("color commands", () => {
  it("set-color applies the chosen color to the chosen target", async () => {
    const a = box("a", "#FF0000");
    ctx = new MockContext([a], "bad-png");
    const r = await executeCommand(findCommand("set-color")!, { ...OPTS, color: "#123456", colorTarget: "font" });
    expect(r).toEqual({ ok: true, changed: 1 });
    expect(a.formatWrites).toEqual([{ path: "font.color", value: "#123456" }]);
  });

  it("no-fill clears and transparency-up steps by 0.1", async () => {
    const a = box("a", "#FF0000");
    ctx = new MockContext([a], "bad-png");
    await executeCommand(findCommand("transparency-up")!, OPTS);
    expect(a.formatWrites).toEqual([{ path: "fill.transparency", value: 0.1 }]);
    await executeCommand(findCommand("no-fill")!, OPTS);
    expect(a.formatWrites.at(-1)).toEqual({ path: "fill.clear", value: null });
  });

  it("set-color without a color is a no-op", async () => {
    ctx = new MockContext([box("a", "#FF0000")], "bad-png");
    expect(await executeCommand(findCommand("set-color")!, OPTS)).toEqual({ ok: true, changed: 0 });
  });
});
```

- [ ] **Step 3: Ejecutar**

Run: `npm test && npm run typecheck`
Expected: todo en verde.

- [ ] **Step 4: Commit**

```bash
git add src/commands tests/commands
git commit -m "feat(commands): ejecutor por tipo de comando con formato y portapapeles"
```

---

### Task 6: Panel: secciones Format y Color, paleta y recientes

**Files:**
- Modify: `src/taskpane/icons.ts`
- Modify: `src/taskpane/sanitize.ts`
- Modify: `src/taskpane/settings.ts`
- Modify: `src/taskpane/panel.ts`
- Modify: `src/taskpane/styles.css`
- Modify: `tests/taskpane/settings.test.ts`

- [ ] **Step 1: Iconos nuevos en `src/taskpane/icons.ts`**

Añadir a `icons` (mismo estilo `S`, `REF`, `box`):

```ts
  copyFormat: `<svg ${S}>${box(2, 2, 7, 9)}<path d="M9 4h4v10H5v-2" /><path d="M4 5.5 6 7l2.5-3" ${REF}/></svg>`,
  pasteFormat: `<svg ${S}><path d="M3 3h10v11H3z"/><path d="M6 3V1.5h4V3"/><path d="M5 8h6M5 11h6" ${REF}/></svg>`,
  pasteFill: `<svg ${S}><path d="M3 3h10v11H3z"/>${box(5, 6, 6, 6)}</svg>`,
  pasteLine: `<svg ${S}><path d="M3 3h10v11H3z"/><rect x="5" y="6" width="6" height="6" ${REF}/></svg>`,
  pasteFont: `<svg ${S}><path d="M3 3h10v11H3z"/><path d="M5.5 12 8 5.5 10.5 12M6.3 10h3.4" ${REF}/></svg>`,
  pasteSize: `<svg ${S}><path d="M3 3h10v11H3z"/><path d="M5.5 11.5v-6h6" ${REF}/><path d="M5.5 5.5 7 7M5.5 5.5 7 4"/></svg>`,
  noFill: `<svg ${S}><rect x="3" y="3" width="10" height="10"/><path d="M3 13 13 3" ${REF}/></svg>`,
  noLine: `<svg ${S}><rect x="3" y="3" width="10" height="10" stroke-dasharray="2 2"/><path d="M3 13 13 3" ${REF}/></svg>`,
  transparencyUp: `<svg ${S}>${box(2, 4, 8, 8)}<rect x="6" y="2" width="8" height="8" fill="currentColor" fill-opacity=".1"/><path d="M11 13v3M9.5 14.5h3" ${REF}/></svg>`,
  transparencyDown: `<svg ${S}>${box(2, 4, 8, 8)}<rect x="6" y="2" width="8" height="8" fill="currentColor" fill-opacity=".6"/><path d="M9.5 14.5h3" ${REF}/></svg>`,
  targetFill: `<svg ${S}>${box(3, 3, 10, 10)}</svg>`,
  targetLine: `<svg ${S}><rect x="3" y="3" width="10" height="10" stroke-width="2"/></svg>`,
  targetFont: `<svg ${S}><path d="M4 13 8 3l4 10M5.5 9.5h5"/></svg>`,
  setColor: `<svg ${S}><circle cx="8" cy="8" r="5"/></svg>`,
```

- [ ] **Step 2: Persistencia en `src/taskpane/sanitize.ts`**

Añadir `recentColors: string[]` y `colorTarget: ColorTarget` y `pasteParts: FormatParts` a `PersistedOptions` y `DEFAULT_PERSISTED` (`recentColors: []`, `colorTarget: "fill"`, `pasteParts: { fill: true, line: true, font: true, size: true }`). Añadir:

```ts
export const MAX_RECENT_COLORS = 8;

export function sanitizeRecentColors(input: unknown): string[] {
  if (!Array.isArray(input)) return [];
  const out: string[] = [];
  for (const v of input) {
    if (typeof v !== "string") continue;
    const hex = normalizeHex(v);
    if (hex && !out.includes(hex)) out.push(hex);
    if (out.length === MAX_RECENT_COLORS) break;
  }
  return out;
}

export function sanitizeColorTarget(input: unknown): ColorTarget {
  return input === "line" || input === "font" ? input : "fill";
}

export function sanitizePasteParts(input: unknown): FormatParts {
  const o = (typeof input === "object" && input !== null ? input : {}) as Record<string, unknown>;
  const b = (v: unknown) => (typeof v === "boolean" ? v : true);
  return { fill: b(o.fill), line: b(o.line), font: b(o.font), size: b(o.size) };
}
```

`toCommandOptions(persisted)` devuelve también `colorTarget: persisted.colorTarget` y `pasteParts: persisted.pasteParts` (sin `color`: ese se pasa en el momento del clic). `settings.ts`: `loadPersistedOptions` sanea los tres campos nuevos.

Tests en `tests/taskpane/settings.test.ts`: `sanitizeRecentColors` (no array → [], filtra inválidos, dedup, máximo 8, normaliza), `sanitizeColorTarget`, `sanitizePasteParts` (booleans por defecto true), `toCommandOptions` incluye `colorTarget` y `pasteParts`.

- [ ] **Step 3: Panel `src/taskpane/panel.ts`**

Cambios:
- `render()`: para cada sección de `SECTION_ORDER`, los botones normales (`commandsBySection`). Tras la fila de botones de "format", añadir una fila de 4 casillas (`<label class="st-check"><input type="checkbox"> Relleno</label>`, Línea, Fuente, Tamaño) ligadas a `persisted.pasteParts`. Tras la fila de botones de "color", añadir: (a) selector de destino: tres `<button class="st-btn st-toggle">` con iconos `targetFill`/`targetLine`/`targetFont`, `aria-pressed`, el activo con clase `is-active`; (b) paleta: `DEFAULT_PALETTE.map(swatch)`; (c) fila "Recientes" con `persisted.recentColors.map(swatch)` (se re-renderiza al cambiar); (d) `<label class="st-field">Otro <input type="color"></label>` cuyo `change` aplica el color.
- `swatch(color)`: `<button class="st-swatch" type="button" aria-label="Color ${color}" title="${color}" style="--swatch:${color}">`; al hacer clic: `applyColor(color)`.
- `applyColor(color)`: `const hex = normalizeHex(color); if (!hex) return; await this.run(findCommand("set-color")!, { color: hex })` y, si el resultado es ok, `persisted.recentColors = pushRecent(persisted.recentColors, hex, MAX_RECENT_COLORS)`, guardar, re-renderizar solo la fila de recientes.
- `run(cmd, extra?: Partial<CommandOptions>)`: ejecuta con `{ ...this.options, ...extra }`; si `result.ok && result.message` → `showNotice(result.message)`; si `result.ok` tras un `copy-format` → `this.hasClipboard = true` y `setSelectionCount(this.lastCount)` para habilitar los botones de pegar.
- `setSelectionCount`: usa `disabledReason(cmd, count, this.hasClipboard)`.
- `style` del swatch: el color va en una custom property (`--swatch`) y el CSS lo pinta con `background: var(--swatch)`; no hay HTML inyectado con datos del usuario (el color ya pasó por `normalizeHex`).

- [ ] **Step 4: CSS `src/taskpane/styles.css`**

```css
.st-swatch { width: 18px; height: 18px; border-radius: 3px; border: 1px solid var(--st-border); background: var(--swatch); cursor: pointer; padding: 0; }
.st-swatch:focus-visible { outline: 2px solid var(--st-accent); outline-offset: 1px; }
.st-toggle[aria-pressed="true"], .st-toggle.is-active { background: var(--st-hover); border-color: var(--st-accent); }
.st-check { display: inline-flex; align-items: center; gap: 3px; margin-right: 6px; color: var(--st-muted); }
.st-subtitle { color: var(--st-muted); font-size: 11px; margin: 4px 0 2px; }
```

- [ ] **Step 5: Verificar**

Run: `npm test && npm run typecheck && npm run build`
Expected: todo en verde; el test de iconos sigue garantizando que todo comando tiene icono (incluido `set-color`).

- [ ] **Step 6: Prueba manual en PowerPoint (Mac)**

`npm run sideload:mac` + `npm run dev` (certificados ya instalados). Checklist:
1. Seleccionar una forma con relleno azul y texto; "Copiar formato" muestra "Formato copiado." y habilita los botones de pegar.
2. Seleccionar dos formas distintas; "Pegar formato" copia relleno, línea, fuente y tamaño; desmarcar "Tamaño" y repetir sobre una tercera: el tamaño no cambia.
3. Destino "Relleno" + pulsar una muestra: cambia el relleno; destino "Texto" + muestra: cambia el color del texto; una imagen seleccionada junto a un cuadro de texto no da error (la imagen se omite para texto).
4. "Sin relleno", "Sin línea", transparencia + y − se comportan según sus tooltips; Cmd+Z deshace cada acción de una vez.
5. El color personalizado y los recientes persisten tras cerrar y abrir el panel.

- [ ] **Step 7: Commit**

```bash
git add src/taskpane tests/taskpane
git commit -m "feat(taskpane): secciones Format y Color con paleta, recientes y destino de color"
```

---

### Task 7: README y versión

**Files:**
- Modify: `README.md`
- Modify: `manifest.xml` (`<Version>1.1.0.0</Version>`), `package.json` (`"version": "1.1.0"`)

- [ ] **Step 1:** Añadir en README la sección "Comandos (Fase 2)": Format (copiar formato; pegar todo/relleno/línea/fuente/tamaño; casillas de partes) y Color (destino relleno/línea/texto, paleta neutra de 16 colores, 8 recientes, color personalizado, sin relleno, sin línea, transparencia ±10 %). Nota: "La paleta es fija porque Office.js no expone los colores del tema". Actualizar la hoja de ruta (Fase 2 hecha).
- [ ] **Step 2:** Subir versiones y validar: `npm run validate`.
- [ ] **Step 3:** Commit: `git commit -am "docs: README Fase 2 y versión 1.1.0"`.

---

## Autorrevisión del plan

**Cobertura del spec § Fase 2:** copiar formato (Task 4/5) ✔; pegar todo/solo fill/line/font/size (Task 4) ✔; paleta + recientes aplicable a relleno, línea o texto (Task 2, 6) ✔ (colores del tema sustituidos por paleta fija: limitación del API, documentada en Task 7); transparencia en pasos (Task 1, 4) ✔; sin relleno / sin línea (Task 1, 4) ✔.

**Placeholders:** ninguno.

**Consistencia de tipos:** `ShapeFormat`/`FormatPatch` (Task 1) se usan igual en Tasks 3–5. `readSelection` devuelve `formats` (Task 3) y lo consume `execute.ts` (Task 5). `CommandOptions` gana `color`, `colorTarget`, `pasteParts` (Task 4) y los produce `toCommandOptions` (Task 6). `disabledReason(cmd, count, hasClipboard)` (Task 4) lo usa el panel (Task 6) y el ejecutor (Task 5). `CommandOutcome`/`CommandResult.message` (Task 3) lo muestra el panel (Task 6).
