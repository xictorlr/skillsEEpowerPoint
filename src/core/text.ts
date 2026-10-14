/** Formato de texto de una forma con marco de texto. `null` = valor mixto. */
export interface Margins { left: number; right: number; top: number; bottom: number; }

export interface TextShape {
  id: string;
  text: string;
  fontSize: number | null;
  horizontalAlignment: string | null;   // PowerPoint.ParagraphHorizontalAlignment
  verticalAlignment: string;            // PowerPoint.TextVerticalAlignment
  autoSize: string;                     // PowerPoint.ShapeAutoSize
  wordWrap: boolean;
  margins: Margins;
}

export interface TextPatch {
  id: string;
  text?: string;
  fontSize?: number;
  horizontalAlignment?: string;
  verticalAlignment?: string;
  autoSize?: string;
  wordWrap?: boolean;
  margins?: Margins;
}

export type CaseMode = "upper" | "lower" | "sentence" | "title";
export type MarginPreset = "zero" | "standard";

export const ZERO_MARGINS: Margins = { left: 0, right: 0, top: 0, bottom: 0 };
/** Márgenes por defecto de PowerPoint: 0,1" lateral y 0,05" vertical. */
export const STANDARD_MARGINS: Margins = { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 };
export const MIN_FONT_SIZE = 1;
export const MAX_FONT_SIZE = 400;
const LOCALE = "es";

export interface Line { text: string; sep: string; }

/** Divide por separadores de párrafo/línea conservando cada separador. */
export function splitLines(text: string): Line[] {
  const lines: Line[] = [];
  let current = "";
  for (const ch of text) {
    if (ch === "\r" || ch === "\n" || ch === "\v") {
      lines.push({ text: current, sep: ch });
      current = "";
    } else current += ch;
  }
  lines.push({ text: current, sep: "" });
  return lines;
}

function joinLines(lines: Line[]): string {
  return lines.map((l) => l.text + l.sep).join("");
}

function sentenceCase(line: string): string {
  const lower = line.toLocaleLowerCase(LOCALE);
  let out = "";
  let capitalizeNext = true;
  for (const ch of lower) {
    if (capitalizeNext && /\p{L}/u.test(ch)) {
      out += ch.toLocaleUpperCase(LOCALE);
      capitalizeNext = false;
    } else {
      out += ch;
      if (ch === "." || ch === "!" || ch === "?") capitalizeNext = true;
    }
  }
  return out;
}

function titleCase(line: string): string {
  return line
    .toLocaleLowerCase(LOCALE)
    .replace(/(^|[\s(\["'¿¡-])(\p{L})/gu, (_m, pre: string, letter: string) => pre + letter.toLocaleUpperCase(LOCALE));
}

export function changeCase(text: string, mode: CaseMode): string {
  const lines = splitLines(text).map((l) => {
    if (mode === "upper") return { ...l, text: l.text.toLocaleUpperCase(LOCALE) };
    if (mode === "lower") return { ...l, text: l.text.toLocaleLowerCase(LOCALE) };
    if (mode === "sentence") return { ...l, text: sentenceCase(l.text) };
    return { ...l, text: titleCase(l.text) };
  });
  return joinLines(lines);
}

/** Colapsa espacios repetidos, recorta extremos de cada línea y elimina líneas vacías. */
export function cleanText(text: string): string {
  const lines = splitLines(text)
    .map((l) => ({ ...l, text: l.text.replace(/[ \t ]+/g, " ").trim() }))
    .filter((l) => l.text.length > 0);
  if (lines.length === 0) return "";
  const last = lines[lines.length - 1];
  if (last) last.sep = "";
  return joinLines(lines);
}

export function casePatches(shapes: TextShape[], mode: CaseMode): TextPatch[] {
  const out: TextPatch[] = [];
  for (const s of shapes) {
    const text = changeCase(s.text, mode);
    if (text !== s.text) out.push({ id: s.id, text });
  }
  return out;
}

export function cleanPatches(shapes: TextShape[]): TextPatch[] {
  const out: TextPatch[] = [];
  for (const s of shapes) {
    const text = cleanText(s.text);
    if (text !== s.text) out.push({ id: s.id, text });
  }
  return out;
}

/** Suma `delta` puntos al tamaño de fuente. Omite formas con tamaño mixto (null). */
export function fontSizePatches(shapes: TextShape[], delta: number): TextPatch[] {
  const out: TextPatch[] = [];
  for (const s of shapes) {
    if (s.fontSize === null) continue;
    const next = Math.min(MAX_FONT_SIZE, Math.max(MIN_FONT_SIZE, s.fontSize + delta));
    if (next !== s.fontSize) out.push({ id: s.id, fontSize: next });
  }
  return out;
}

export function paragraphAlignPatches(shapes: TextShape[], align: string): TextPatch[] {
  return shapes.filter((s) => s.horizontalAlignment !== align).map((s) => ({ id: s.id, horizontalAlignment: align }));
}

export function verticalAlignPatches(shapes: TextShape[], align: string): TextPatch[] {
  return shapes.filter((s) => s.verticalAlignment !== align).map((s) => ({ id: s.id, verticalAlignment: align }));
}

function sameMargins(a: Margins, b: Margins): boolean {
  return Math.abs(a.left - b.left) < 0.01 && Math.abs(a.right - b.right) < 0.01 && Math.abs(a.top - b.top) < 0.01 && Math.abs(a.bottom - b.bottom) < 0.01;
}

export function marginPatches(shapes: TextShape[], preset: MarginPreset): TextPatch[] {
  const target = preset === "zero" ? ZERO_MARGINS : STANDARD_MARGINS;
  return shapes.filter((s) => !sameMargins(s.margins, target)).map((s) => ({ id: s.id, margins: target }));
}

export function autoSizePatches(shapes: TextShape[], mode: string): TextPatch[] {
  return shapes.filter((s) => s.autoSize !== mode).map((s) => ({ id: s.id, autoSize: mode }));
}

export function wordWrapPatches(shapes: TextShape[], on: boolean): TextPatch[] {
  return shapes.filter((s) => s.wordWrap !== on).map((s) => ({ id: s.id, wordWrap: on }));
}
