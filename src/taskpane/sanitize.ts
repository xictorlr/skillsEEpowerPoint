import type { CommandOptions } from "../commands/types";
import { normalizeHex } from "../core/color";
import type { ColorTarget, FormatParts } from "../core/format";
import type { MatrixOptions } from "../core/geometry";
import type { SlideSize } from "../core/types";

/**
 * Valores por defecto de la matriz. Se usan cuando falta un campo o su valor no es válido.
 * Archivo puro (sin DOM ni localStorage) para poder testearlo bajo Node.
 */
export const DEFAULT_MATRIX: MatrixOptions = { columns: 3, gapX: 8, gapY: 8 };

function pick(value: unknown, min: number, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value >= min ? value : fallback;
}

/** Sanea un valor arbitrario (p.ej. JSON.parse de localStorage, o un <input>) a una MatrixOptions válida. */
export function sanitizeMatrix(input: unknown): MatrixOptions {
  const obj = input !== null && typeof input === "object" ? (input as Record<string, unknown>) : {};
  const columns = pick(obj.columns, 1, DEFAULT_MATRIX.columns);
  return {
    columns: Math.round(columns),
    gapX: pick(obj.gapX, 0, DEFAULT_MATRIX.gapX),
    gapY: pick(obj.gapY, 0, DEFAULT_MATRIX.gapY),
  };
}

/**
 * Tamaños de diapositiva predefinidos. `auto` deja la decisión a la detección automática (que asume
 * 540pt de alto); el resto fuerza un tamaño explícito — útil para mazos 16:9 "clásicos" (720x405) u
 * otros formatos donde esa heurística no aplica. "4x3" y "letter" comparten medidas (720x540): se
 * mantienen como entradas separadas porque la etiqueta que ve el usuario es distinta.
 */
export const SLIDE_SIZE_PRESETS = {
  auto: undefined,
  "16x9": { width: 960, height: 540 },
  "16x9-legacy": { width: 720, height: 405 },
  "4x3": { width: 720, height: 540 },
  a4: { width: 780, height: 540 },
  letter: { width: 720, height: 540 },
} as const satisfies Record<string, SlideSize | undefined>;

export type SlideSizePreset = keyof typeof SLIDE_SIZE_PRESETS;

const SLIDE_SIZE_PRESET_KEYS = Object.keys(SLIDE_SIZE_PRESETS) as SlideSizePreset[];

/** Sanea un valor arbitrario a una clave válida de SLIDE_SIZE_PRESETS; "auto" si no la reconoce. */
export function sanitizeSlideSizePreset(input: unknown): SlideSizePreset {
  return typeof input === "string" && (SLIDE_SIZE_PRESET_KEYS as string[]).includes(input)
    ? (input as SlideSizePreset)
    : "auto";
}

/** Número máximo de colores recientes que se recuerdan. */
export const MAX_RECENT_COLORS = 8;

/** Sanea un valor arbitrario a una lista de colores "#RRGGBB" válidos, sin duplicados y acotada. */
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

/** Sanea un valor arbitrario a un ColorTarget válido; "fill" si no lo reconoce. */
export function sanitizeColorTarget(input: unknown): ColorTarget {
  return input === "line" || input === "font" ? input : "fill";
}

/** Sanea un valor arbitrario a un FormatParts válido; cada campo ausente o no booleano se asume true. */
export function sanitizePasteParts(input: unknown): FormatParts {
  const o = (typeof input === "object" && input !== null ? input : {}) as Record<string, unknown>;
  const b = (v: unknown) => (typeof v === "boolean" ? v : true);
  return { fill: b(o.fill), line: b(o.line), font: b(o.font), size: b(o.size) };
}

/** Forma persistida en localStorage (ver settings.ts): matriz, preset de tamaño de diapositiva,
 *  colores recientes, destino de color y partes a pegar por defecto. */
export interface PersistedOptions {
  matrix: MatrixOptions;
  slideSizePreset: SlideSizePreset;
  recentColors: string[];
  colorTarget: ColorTarget;
  pasteParts: FormatParts;
}

export const DEFAULT_PERSISTED: PersistedOptions = {
  matrix: DEFAULT_MATRIX,
  slideSizePreset: "auto",
  recentColors: [],
  colorTarget: "fill",
  pasteParts: { fill: true, line: true, font: true, size: true },
};

/** Convierte las opciones persistidas (ya saneadas) en las CommandOptions que consumen los comandos.
 *  No incluye `color`: ese se pasa en el momento del clic sobre una muestra. */
export function toCommandOptions(persisted: PersistedOptions): CommandOptions {
  return {
    matrix: persisted.matrix,
    slideSize: SLIDE_SIZE_PRESETS[persisted.slideSizePreset],
    colorTarget: persisted.colorTarget,
    pasteParts: persisted.pasteParts,
    // TODO Task 4: persistir el paso de tamaño de fuente en PersistedOptions.
    fontSizeStep: 1,
  };
}
