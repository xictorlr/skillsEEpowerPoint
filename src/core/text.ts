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

/** Edición de un sub-rango de texto. `start`/`length` son índices UTF-16 sobre el texto ANTIGUO (como en los strings de JS). */
export interface TextEdit { start: number; length: number; text: string }

export interface TextPatch {
  id: string;
  /**
   * Ediciones de sub-rango a aplicar sobre `textRange`, p. ej. con
   * `textRange.getSubstring(start, length).text = text`. Vienen ordenadas por
   * `start` DESCENDENTE y deben aplicarse en el ORDEN DEL ARREGLO (de la última
   * a la primera) para no desplazar los índices de las ediciones pendientes.
   */
  edits?: TextEdit[];
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

interface Token { start: number; length: number; text: string }

/** Tokeniza en tramos máximos de no-espacio, con su offset dentro de la línea. */
function wordTokens(line: string): Token[] {
  const out: Token[] = [];
  const re = /\S+/gu;
  let m: RegExpExecArray | null;
  while ((m = re.exec(line)) !== null) {
    out.push({ start: m.index, length: m[0].length, text: m[0] });
  }
  return out;
}

/** Reconstruye una cadena aplicando ediciones locales (offsets propios, no acumulativos). */
function applyLocalEdits(text: string, edits: Token[]): string {
  let out = text;
  for (const e of [...edits].sort((a, b) => b.start - a.start)) {
    out = out.slice(0, e.start) + e.text + out.slice(e.start + e.length);
  }
  return out;
}

/**
 * Ediciones por palabra (offsets locales a la línea) si, y solo si, ambas líneas
 * tokenizan al mismo número de palabras, cada par de tokens tiene el mismo `start`
 * y `length` (solo cambia el contenido), y reconstruir `oldLine` con esas ediciones
 * reproduce `newLine` exactamente (esto último descarta cambios en los espacios
 * ENTRE palabras, p. ej. un espacio que pasa a ser tabulador). Si no se cumple todo
 * eso, devuelve `null` para que la llamadora recurra a una única edición de línea completa.
 */
function alignedWordEdits(oldLine: string, newLine: string): Token[] | null {
  const oldTokens = wordTokens(oldLine);
  const newTokens = wordTokens(newLine);
  if (oldTokens.length !== newTokens.length) return null;
  const out: Token[] = [];
  for (let t = 0; t < oldTokens.length; t++) {
    const ot = oldTokens[t]!;
    const nt = newTokens[t]!;
    if (ot.start !== nt.start || ot.length !== nt.length) return null;
    if (ot.text !== nt.text) out.push({ start: ot.start, length: ot.length, text: nt.text });
  }
  // Si ningún token cambió pero las líneas difieren, la diferencia está en los
  // espacios entre tokens: no es representable a nivel de palabra.
  if (out.length === 0) return null;
  return applyLocalEdits(oldLine, out) === newLine ? out : null;
}

/**
 * Calcula ediciones de sub-rango (índices UTF-16 sobre `oldText`) que transforman
 * `oldText` en `newText`, preservando el formato de las partes sin cambios:
 * - Si el número de párrafos difiere, o si algún separador de párrafo cambia (p. ej.
 *   `\r` por `\v`), emite una única edición de reemplazo de todo el texto (no hay una
 *   correspondencia 1:1 de párrafos en la que basar ediciones parciales).
 * - Si no, compara párrafo a párrafo: si ambas líneas tokenizan al mismo número de
 *   palabras y cada par de tokens coincide en posición y longitud (p. ej. un cambio de
 *   caja), emite ediciones por palabra solo para las palabras que cambian; en cualquier
 *   otro caso (longitud de línea distinta, o palabras desalineadas por un espacio que se
 *   movió) emite una única edición para el párrafo completo.
 * El resultado viene ordenado por `start` DESCENDENTE, listo para aplicarse en ese orden.
 */
export function computeTextEdits(oldText: string, newText: string): TextEdit[] {
  if (oldText === newText) return [];
  const oldLines = splitLines(oldText);
  const newLines = splitLines(newText);
  if (oldLines.length !== newLines.length) {
    return [{ start: 0, length: oldText.length, text: newText }];
  }
  for (let i = 0; i < oldLines.length; i++) {
    if (oldLines[i]!.sep !== newLines[i]!.sep) return [{ start: 0, length: oldText.length, text: newText }];
  }
  const edits: TextEdit[] = [];
  let offset = 0;
  for (let i = 0; i < oldLines.length; i++) {
    const oldLine = oldLines[i]!;
    const newLine = newLines[i]!;
    if (oldLine.text !== newLine.text) {
      const wordEdits = oldLine.text.length === newLine.text.length
        ? alignedWordEdits(oldLine.text, newLine.text)
        : null;
      if (wordEdits) {
        for (const e of wordEdits) edits.push({ start: offset + e.start, length: e.length, text: e.text });
      } else {
        edits.push({ start: offset, length: oldLine.text.length, text: newLine.text });
      }
    }
    offset += oldLine.text.length + oldLine.sep.length;
  }
  edits.sort((a, b) => b.start - a.start);
  return edits;
}

export function casePatches(shapes: TextShape[], mode: CaseMode): TextPatch[] {
  const out: TextPatch[] = [];
  for (const s of shapes) {
    const edits = computeTextEdits(s.text, changeCase(s.text, mode));
    if (edits.length > 0) out.push({ id: s.id, edits });
  }
  return out;
}

export function cleanPatches(shapes: TextShape[]): TextPatch[] {
  const out: TextPatch[] = [];
  for (const s of shapes) {
    const edits = computeTextEdits(s.text, cleanText(s.text));
    if (edits.length > 0) out.push({ id: s.id, edits });
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

/** Compara sobre el valor base (sin el sufijo `Centered`) y, si el valor actual lo tenía, lo conserva en el resultado. */
export function verticalAlignPatches(shapes: TextShape[], align: string): TextPatch[] {
  const out: TextPatch[] = [];
  for (const s of shapes) {
    const v = s.verticalAlignment;
    if (v.replace(/Centered$/, "") === align) continue;
    out.push({ id: s.id, verticalAlignment: v.endsWith("Centered") ? `${align}Centered` : align });
  }
  return out;
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
