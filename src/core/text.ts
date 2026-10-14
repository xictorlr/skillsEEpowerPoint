export interface Margins { left: number; right: number; top: number; bottom: number; }

/** Formato de texto de una forma con marco de texto. `null` = valor mixto. */
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

export const ZERO_MARGINS: Margins = Object.freeze({ left: 0, right: 0, top: 0, bottom: 0 });
/** Márgenes por defecto de PowerPoint: 0,1" lateral y 0,05" vertical. */
export const STANDARD_MARGINS: Margins = Object.freeze({ left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 });
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

/**
 * Mayúscula inicial de cada oración. Un terminador (`.`, `!`, `?`, `…`) cierra una
 * oración solo cuando le sigue un espacio (puede haber cierres de comillas/paréntesis
 * `"'”’»)]` entre medias). Un dígito puede iniciar oración, pero solo las letras cambian
 * de caja. Trata `\v` como espacio, de modo que el estado de capitalización se mantiene
 * a través de saltos de línea suaves.
 */
function sentenceCase(line: string): string {
  let out = "";
  let cap = true;
  let pending = false;
  for (const ch of line.toLocaleLowerCase(LOCALE)) {
    if (cap && /[\p{L}\p{N}]/u.test(ch)) {
      out += /\p{L}/u.test(ch) ? ch.toLocaleUpperCase(LOCALE) : ch;
      cap = false;
      continue;
    }
    out += ch;
    if (ch === "." || ch === "!" || ch === "?" || ch === "…") pending = true;
    else if (pending && /\s/u.test(ch)) { cap = true; pending = false; }
    else if (pending && !/["'”’»)\]]/u.test(ch)) pending = false;
  }
  return out;
}

/**
 * Aplica `sentenceCase` agrupando líneas unidas por saltos suaves (`\v`) en un único
 * párrafo (el estado de capitalización cruza el `\v`), y reinicia solo en separadores
 * duros (`\r`, `\n`) o al final del texto.
 */
function sentenceCaseAcrossSoftBreaks(text: string): string {
  const lines = splitLines(text);
  const out: Line[] = [];
  let i = 0;
  while (i < lines.length) {
    let j = i;
    let paragraph = lines[i]?.text ?? "";
    while (j < lines.length - 1 && lines[j]?.sep === "\v") {
      j++;
      paragraph += "\v" + (lines[j]?.text ?? "");
    }
    const parts = sentenceCase(paragraph).split("\v");
    for (let k = i; k <= j; k++) {
      out.push({ text: parts[k - i] ?? "", sep: lines[k]?.sep ?? "" });
    }
    i = j + 1;
  }
  return joinLines(out);
}

function titleCase(line: string): string {
  return line
    .toLocaleLowerCase(LOCALE)
    .replace(/(^|[\s(\[«"“‘¿¡\-–—/]|(?:^|\s)')(\p{L})/gu, (_m, pre: string, letter: string) => pre + letter.toLocaleUpperCase(LOCALE));
}

export function changeCase(text: string, mode: CaseMode): string {
  if (mode === "sentence") return sentenceCaseAcrossSoftBreaks(text);
  const lines = splitLines(text).map((l) => {
    if (mode === "upper") return { ...l, text: l.text.toLocaleUpperCase(LOCALE) };
    if (mode === "lower") return { ...l, text: l.text.toLocaleLowerCase(LOCALE) };
    return { ...l, text: titleCase(l.text) };
  });
  return joinLines(lines);
}

/** Elimina espacios dobles, espacios al inicio/fin de línea y párrafos vacíos; conserva tabuladores y espacios duros. */
export function cleanText(text: string): string {
  const lines = splitLines(text)
    .map((l) => ({ ...l, text: l.text.replace(/ {2,}/g, " ").replace(/^ +/, "").replace(/ +$/, "") }))
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
  if (!Number.isFinite(delta) || delta === 0) return [];
  const out: TextPatch[] = [];
  for (const s of shapes) {
    if (s.fontSize === null) continue;
    const clamped = Math.min(MAX_FONT_SIZE, Math.max(MIN_FONT_SIZE, s.fontSize + delta));
    const next = Math.round(clamped * 10) / 10;
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
  return shapes.filter((s) => !sameMargins(s.margins, target)).map((s) => ({ id: s.id, margins: { ...target } }));
}

export function autoSizePatches(shapes: TextShape[], mode: string): TextPatch[] {
  return shapes.filter((s) => s.autoSize !== mode).map((s) => ({ id: s.id, autoSize: mode }));
}

export function wordWrapPatches(shapes: TextShape[], on: boolean): TextPatch[] {
  return shapes.filter((s) => s.wordWrap !== on).map((s) => ({ id: s.id, wordWrap: on }));
}
