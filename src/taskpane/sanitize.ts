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
 * Tamaños de diapositiva predefinidos. `auto` deja la decisión a la detección automática: es exacta
 * (vía `pageSetup`) en hosts con PowerPointApi 1.10; en el resto, asume 540pt de alto (heurística
 * PNG). El resto de entradas fuerza un tamaño explícito — útil para mazos 16:9 "clásicos" (720x405)
 * u otros formatos donde esa heurística no aplica. "4x3" y "letter" comparten medidas (720x540): se
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

/** Longitud máxima del prefijo de "Renombrar numerado". */
export const MAX_RENAME_PREFIX_LENGTH = 40;

/** Prefijo por defecto de "Renombrar numerado". */
export const DEFAULT_RENAME_PREFIX = "Forma";

/** Sanea un valor arbitrario a un prefijo válido para "Renombrar numerado": recortado (trim) y
 *  acotado a MAX_RENAME_PREFIX_LENGTH caracteres. Un valor no-string usa el valor por defecto; una
 *  cadena vacía (tras recortar) se persiste tal cual -- no hace falta que esta función la sustituya
 *  por "Forma", porque el comando `rename-numbered` ya cae a "Forma" si el prefijo que recibe
 *  está vacío (ver `options.renamePrefix || "Forma"` en commands/expert.ts). */
export function sanitizeRenamePrefix(input: unknown): string {
  if (typeof input !== "string") return DEFAULT_RENAME_PREFIX;
  return input.trim().slice(0, MAX_RENAME_PREFIX_LENGTH);
}

/** Paso por defecto (en puntos) al aumentar/reducir el tamaño de fuente. */
export const DEFAULT_FONT_SIZE_STEP = 1;

/** Sanea un valor arbitrario a un paso de tamaño de fuente válido: un número finito se acota a
 *  [0.5, 20] y se redondea al 0.5 más cercano; cualquier otro valor usa el valor por defecto. */
export function sanitizeFontSizeStep(input: unknown): number {
  if (typeof input !== "number" || !Number.isFinite(input)) {
    return DEFAULT_FONT_SIZE_STEP;
  }
  const clamped = Math.min(20, Math.max(0.5, input));
  return Math.round(clamped * 2) / 2;
}

/** Título por defecto del asistente "Crear agenda" (sección Wizards). */
export const DEFAULT_AGENDA_TITLE = "Agenda";

/** Longitud máxima del título de la agenda. */
export const MAX_AGENDA_TITLE_LENGTH = 60;

/** Longitud máxima del texto de secciones (textarea, una sección por línea) de la agenda. */
export const MAX_AGENDA_SECTIONS_TEXT_LENGTH = 2000;

/** Título y texto de secciones persistidos del asistente "Crear agenda"; ver `commands/wizards.ts`
 *  para el formato esperado de `sectionsText` (una sección por línea, `parseAgendaSections`). */
export interface AgendaSettings {
  title: string;
  sectionsText: string;
}

/** Sanea el título: no-string o vacío tras recortar (trim) cae a DEFAULT_AGENDA_TITLE; el resultado
 *  se acota a MAX_AGENDA_TITLE_LENGTH caracteres. */
export function sanitizeAgendaTitle(input: unknown): string {
  if (typeof input !== "string") return DEFAULT_AGENDA_TITLE;
  const trimmed = input.trim();
  return (trimmed.length > 0 ? trimmed : DEFAULT_AGENDA_TITLE).slice(0, MAX_AGENDA_TITLE_LENGTH);
}

/** Sanea el texto de secciones: un no-string se convierte en cadena vacía; el resultado se acota a
 *  MAX_AGENDA_SECTIONS_TEXT_LENGTH caracteres. No se recorta (trim): los saltos de línea y espacios
 *  internos son significativos para `parseAgendaSections`. */
export function sanitizeAgendaSectionsText(input: unknown): string {
  return typeof input === "string" ? input.slice(0, MAX_AGENDA_SECTIONS_TEXT_LENGTH) : "";
}

/** Sanea un valor arbitrario (p.ej. JSON.parse de localStorage) a una AgendaSettings válida. */
export function sanitizeAgenda(input: unknown): AgendaSettings {
  const obj = input !== null && typeof input === "object" ? (input as Record<string, unknown>) : {};
  return {
    title: sanitizeAgendaTitle(obj.title),
    sectionsText: sanitizeAgendaSectionsText(obj.sectionsText),
  };
}

export const DEFAULT_AGENDA: AgendaSettings = { title: DEFAULT_AGENDA_TITLE, sectionsText: "" };

/** Forma persistida en localStorage (ver settings.ts): matriz, preset de tamaño de diapositiva,
 *  colores recientes, destino de color, partes a pegar por defecto, paso de tamaño de fuente,
 *  prefijo de "Renombrar numerado" y título/secciones del asistente "Crear agenda". */
export interface PersistedOptions {
  matrix: MatrixOptions;
  slideSizePreset: SlideSizePreset;
  recentColors: string[];
  colorTarget: ColorTarget;
  pasteParts: FormatParts;
  fontSizeStep: number;
  renamePrefix: string;
  agenda: AgendaSettings;
}

export const DEFAULT_PERSISTED: PersistedOptions = {
  matrix: DEFAULT_MATRIX,
  slideSizePreset: "auto",
  recentColors: [],
  colorTarget: "fill",
  pasteParts: { fill: true, line: true, font: true, size: true },
  fontSizeStep: DEFAULT_FONT_SIZE_STEP,
  renamePrefix: DEFAULT_RENAME_PREFIX,
  agenda: DEFAULT_AGENDA,
};

/** Convierte las opciones persistidas (ya saneadas) en las CommandOptions que consumen los comandos.
 *  No incluye `color`: ese se pasa en el momento del clic sobre una muestra. Tampoco incluye
 *  `nameFilter` ni `shapeKind`: "Seleccionar por nombre" y "Seleccionar por tipo" los pasan en el
 *  momento del clic (ver expertSection.ts), no se persisten. */
export function toCommandOptions(persisted: PersistedOptions): CommandOptions {
  return {
    matrix: persisted.matrix,
    slideSize: SLIDE_SIZE_PRESETS[persisted.slideSizePreset],
    colorTarget: persisted.colorTarget,
    pasteParts: persisted.pasteParts,
    fontSizeStep: persisted.fontSizeStep,
    nameFilter: "",
    renamePrefix: persisted.renamePrefix,
    shapeKind: "all",
    agenda: persisted.agenda,
  };
}
