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
