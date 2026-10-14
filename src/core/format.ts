import { normalizeHex } from "./color";

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
  /** null cuando el tipo de forma no admite relleno (p. ej. una línea, una tabla, un gráfico). */
  fill: FillFormat | null;
  /** null cuando el tipo de forma no admite línea (p. ej. una tabla, un gráfico, un grupo). */
  line: LineFormat | null;
  /** null cuando la forma no tiene marco de texto (imagen, tabla, gráfico, un marcador de posición
   *  cuyo contenido no es texto…). */
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

/**
 * Cambio parcial de formato. Solo se incluyen las partes y campos a escribir.
 *
 * Contrato para quien aplique el patch (el adaptador de Office):
 * - `fill.type === "NoFill"` → `shape.fill.clear()`. No se pasan `color` ni `transparency` junto a `NoFill`.
 * - `fill.type === "Solid"` (con o sin `color`) → si viene `color`, `shape.fill.setSolidColor(color)`;
 *   `transparency`, si viene, se aplica sobre el relleno sólido resultante. Un relleno de tipo
 *   `"Other"` (degradado, patrón, imagen) nunca aparece aquí: no se puede pegar/generar ese tipo.
 * - Dentro de `line`: si `visible === true`, se escribe primero (antes que color/peso/estilo/
 *   transparencia); si `visible === false`, se escribe al final. En Office.js, alternar la
 *   visibilidad de la línea puede resetear los demás campos, así que el orden importa según la
 *   dirección del cambio: al mostrar la línea hay que fijar sus demás campos después de hacerla
 *   visible; al ocultarla hay que fijarlos antes de que desaparezca.
 * - `fill`/`line`/`font` solo se incluyen cuando la forma los admite; quien aplique el patch recibe
 *   además la información de capacidades de la forma (ver `FormatCaps` en el adaptador de Office) y
 *   debe ignorar cualquier parte que la forma no soporte.
 * - Los colores (`fill.color`, `line.color`, `font.color`) ya vienen normalizados ("#RRGGBB" en
 *   mayúsculas) al leer el formato desde Office.js, y los que generan estas funciones también lo están.
 */
export interface FormatPatch {
  id: string;
  fill?: { type?: "Solid" | "NoFill"; color?: string; transparency?: number };
  line?: Partial<LineFormat>;
  font?: Partial<FontFormat>;
  width?: number;
  height?: number;
}

/** Tolerancia para comparar transparencias (0..1). */
const EPS = 0.001;

/** Tolerancia (en puntos) para considerar que dos tamaños son iguales. */
const SIZE_EPS = 0.01;

function near(a: number, b: number): boolean {
  return Math.abs(a - b) < SIZE_EPS;
}

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

/**
 * Parte de relleno a pegar desde `fill`. `NoFill` se reduce a su tipo (ignora cualquier color/
 * transparencia residual); `Other` (degradado, patrón, imagen) no se puede pegar y se omite; `Solid`
 * conserva el tipo y los campos no nulos de color/transparencia.
 */
function fillPart(fill: FillFormat): { type?: "Solid" | "NoFill"; color?: string; transparency?: number } | undefined {
  if (fill.type === "NoFill") return { type: "NoFill" };
  if (fill.type === "Other") return undefined;
  const part: { type: "Solid"; color?: string; transparency?: number } = { type: "Solid" };
  if (fill.color !== null) part.color = fill.color;
  if (fill.transparency !== null) part.transparency = fill.transparency;
  return part;
}

/** Parte de línea a pegar desde `line`. Una línea invisible se reduce a `{ visible: false }`, ignorando
 *  color/peso/estilo/transparencia residuales; de lo contrario se copian los campos no nulos. */
function linePart(line: LineFormat): Partial<LineFormat> | undefined {
  if (line.visible === false) return { visible: false };
  return nonNull(line);
}

/**
 * Pega las partes indicadas del formato `source` en cada `target`. No excluye al propio `source`:
 * los ids de forma pueden repetirse entre diapositivas, y volver a escribir su propio formato es
 * inofensivo. Ojo: fill/línea/fuente se pegan siempre que se piden y están presentes en `source`
 * (no se comparan con el valor actual del target); solo `size` se compara con el target y se omite
 * si ya coincide. Por eso, si `source` aparece también como target (mismo id u objeto) y tiene
 * algún fill/línea/fuente para pegar, SÍ aparece en los patches devueltos.
 */
export function buildPastePatches(source: ShapeFormat, targets: ShapeFormat[], parts: FormatParts): FormatPatch[] {
  const fill = parts.fill && source.fill ? fillPart(source.fill) : undefined;
  const line = parts.line && source.line ? linePart(source.line) : undefined;
  const font = parts.font && source.font ? nonNull(source.font) : undefined;

  const patches: FormatPatch[] = [];
  for (const t of targets) {
    const includeFill = fill !== undefined && t.fill !== null;
    const includeLine = line !== undefined && t.line !== null;
    const includeFont = parts.font && !!source.font && !!t.font && font !== undefined;
    const includeSize =
      parts.size && (!near(source.size.width, t.size.width) || !near(source.size.height, t.size.height));
    const patch: FormatPatch = {
      id: t.id,
      ...(includeFill ? { fill } : {}),
      ...(includeLine ? { line } : {}),
      ...(includeFont ? { font } : {}),
      ...(includeSize ? { width: source.size.width, height: source.size.height } : {}),
    };
    if (Object.keys(patch).length > 1) patches.push(patch);
  }
  return patches;
}

/** Aplica un color sólido al relleno, la línea o el texto de cada forma que admita esa parte. */
export function colorPatches(targets: ShapeFormat[], target: ColorTarget, color: string): FormatPatch[] {
  const hex = normalizeHex(color);
  if (!hex) return [];
  const patches: FormatPatch[] = [];
  for (const t of targets) {
    if (target === "fill" && t.fill) patches.push({ id: t.id, fill: { type: "Solid", color: hex } });
    else if (target === "line" && t.line) patches.push({ id: t.id, line: { visible: true, color: hex } });
    else if (target === "font" && t.font) patches.push({ id: t.id, font: { color: hex } });
  }
  return patches;
}

/** Suma `delta` a la transparencia (0..1) de relleno o línea. Se omiten: formas sin esa parte (null)
 *  o sin relleno sólido cuando `target` es "fill", líneas invisibles cuando `target` es "line", y,
 *  siempre, `delta === 0`. */
export function transparencyPatches(targets: ShapeFormat[], target: "fill" | "line", delta: number): FormatPatch[] {
  if (delta === 0) return [];
  const patches: FormatPatch[] = [];
  for (const t of targets) {
    if (target === "fill" && (!t.fill || t.fill.type !== "Solid")) continue;
    if (target === "line" && (!t.line || t.line.visible === false)) continue;
    const current = (target === "fill" ? t.fill?.transparency : t.line?.transparency) ?? 0;
    const next = Math.min(1, Math.max(0, Math.round((current + delta) * 100) / 100));
    if (Math.abs(next - current) < EPS) continue;
    patches.push(target === "fill" ? { id: t.id, fill: { transparency: next } } : { id: t.id, line: { transparency: next } });
  }
  return patches;
}

/** Quita el relleno o la línea de cada forma que los admita y aún los tenga. */
export function clearPatches(targets: ShapeFormat[], target: "fill" | "line"): FormatPatch[] {
  const patches: FormatPatch[] = [];
  for (const t of targets) {
    if (target === "fill" && t.fill && t.fill.type !== "NoFill") patches.push({ id: t.id, fill: { type: "NoFill" } });
    if (target === "line" && t.line && t.line.visible !== false) patches.push({ id: t.id, line: { visible: false } });
  }
  return patches;
}
