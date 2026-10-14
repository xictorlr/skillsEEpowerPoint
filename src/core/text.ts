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
  bold: boolean | null;
  italic: boolean | null;
  underline: string | null;             // PowerPoint.ShapeFontUnderlineStyle; "None" = sin subrayado
  strikethrough: boolean | null;
  bulletsVisible: boolean | null;
  /** `null` cuando el host no soporta PowerPointApi 1.10 (no se puede leer/escribir). */
  indentLevel: number | null;
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
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strikethrough?: boolean;
  bulletsVisible?: boolean;
  indentLevel?: number;
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

/** Ediciones de sub-rango (offsets locales a la línea) para un único par de párrafos: por palabra si se puede, si no de línea completa. */
function lineEdits(oldLineText: string, newLineText: string): Token[] {
  if (oldLineText === newLineText) return [];
  const wordEdits = oldLineText.length === newLineText.length ? alignedWordEdits(oldLineText, newLineText) : null;
  return wordEdits ?? [{ start: 0, length: oldLineText.length, text: newLineText }];
}

/** Offset absoluto (índice UTF-16 sobre el texto completo) de cada línea de `splitLines`. */
function lineOffsets(lines: Line[]): number[] {
  const offsets: number[] = [];
  let acc = 0;
  for (const l of lines) {
    offsets.push(acc);
    acc += l.text.length + l.sep.length;
  }
  return offsets;
}

/** `true` si el párrafo consiste solo en espacios ordinarios (cero o más); es la misma noción de "vacío" que usa `cleanText` (que no recorta tabuladores ni espacios duros, así que un párrafo con solo uno de esos NO es vacío). */
function isEmptyParagraph(text: string): boolean {
  return /^ *$/.test(text);
}

/**
 * Cuando el número de párrafos difiere porque se eliminaron párrafos vacíos (el caso
 * más común de `cleanText`), intenta expresar el cambio como ediciones de sub-rango en
 * vez de una única edición de todo el texto: una edición de borrado (texto + separador)
 * por cada párrafo vacío eliminado, más las ediciones normales de los párrafos que se
 * conservan. Si los párrafos vacíos eliminados estaban al final, el último párrafo
 * conservado pierde su separador; eso se expresa con una edición extra que borra solo
 * ese separador. Devuelve `null` (para que la llamadora recurra a una única edición de
 * todo el texto) si el número de párrafos conservados no coincide con el de `newLines`,
 * o si algún párrafo conservado (salvo el último) tiene un separador que no coincide
 * con el esperado.
 */
function emptyParagraphRemovalEdits(oldText: string, newText: string, oldLines: Line[], newLines: Line[]): TextEdit[] | null {
  const keptIndices: number[] = [];
  for (let i = 0; i < oldLines.length; i++) {
    if (!isEmptyParagraph(oldLines[i]!.text)) keptIndices.push(i);
  }
  if (keptIndices.length !== newLines.length) return null;

  const offsets = lineOffsets(oldLines);
  const edits: TextEdit[] = [];
  const lastKept = keptIndices.length - 1;
  for (let k = 0; k < keptIndices.length; k++) {
    const i = keptIndices[k]!;
    const oldLine = oldLines[i]!;
    const newLine = newLines[k]!;
    const isLast = k === lastKept;
    if (oldLine.sep !== newLine.sep) {
      if (!isLast) return null;
      // `newLine` es la última línea de `splitLines`, así que su separador siempre es "".
      edits.push({ start: offsets[i]! + oldLine.text.length, length: oldLine.sep.length, text: "" });
    }
    for (const e of lineEdits(oldLine.text, newLine.text)) {
      edits.push({ start: offsets[i]! + e.start, length: e.length, text: e.text });
    }
  }

  const keptSet = new Set(keptIndices);
  for (let i = 0; i < oldLines.length; i++) {
    if (keptSet.has(i)) continue;
    const l = oldLines[i]!;
    const length = l.text.length + l.sep.length;
    if (length > 0) edits.push({ start: offsets[i]!, length, text: "" });
  }

  edits.sort((a, b) => b.start - a.start);
  // Red de seguridad: si reconstruir `oldText` con estas ediciones no reproduce `newText`
  // exactamente, no son de fiar; que la llamadora recurra a la edición de todo el texto.
  return applyLocalEdits(oldText, edits) === newText ? edits : null;
}

/**
 * Calcula ediciones de sub-rango (índices UTF-16 sobre `oldText`) que transforman
 * `oldText` en `newText`, preservando el formato de las partes sin cambios:
 * - Si el número de párrafos coincide y ningún separador de párrafo cambia, compara
 *   párrafo a párrafo: si ambas líneas tokenizan al mismo número de palabras y cada par
 *   de tokens coincide en posición y longitud (p. ej. un cambio de caja), emite
 *   ediciones por palabra solo para las palabras que cambian; en cualquier otro caso
 *   (longitud de línea distinta, o palabras desalineadas por un espacio que se movió)
 *   emite una única edición para el párrafo completo.
 * - Si el número de párrafos difiere, intenta primero explicarlo como la eliminación de
 *   párrafos vacíos (ver `emptyParagraphRemovalEdits`); si no se puede, o si algún
 *   separador cambia sin que se pueda justificar así, emite una única edición de
 *   reemplazo de todo el texto (no hay una correspondencia 1:1 de párrafos en la que
 *   basar ediciones parciales).
 * El resultado viene ordenado por `start` DESCENDENTE, listo para aplicarse en ese orden.
 */
export function computeTextEdits(oldText: string, newText: string): TextEdit[] {
  if (oldText === newText) return [];
  const oldLines = splitLines(oldText);
  const newLines = splitLines(newText);
  const wholeText = [{ start: 0, length: oldText.length, text: newText }];
  if (oldLines.length !== newLines.length) {
    return emptyParagraphRemovalEdits(oldText, newText, oldLines, newLines) ?? wholeText;
  }
  for (let i = 0; i < oldLines.length; i++) {
    if (oldLines[i]!.sep !== newLines[i]!.sep) return wholeText;
  }
  const offsets = lineOffsets(oldLines);
  const edits: TextEdit[] = [];
  for (let i = 0; i < oldLines.length; i++) {
    const oldLine = oldLines[i]!;
    const newLine = newLines[i]!;
    for (const e of lineEdits(oldLine.text, newLine.text)) {
      edits.push({ start: offsets[i]! + e.start, length: e.length, text: e.text });
    }
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

/** `true` si hay al menos un valor no nulo en `values` y todos son "on" según `isOn`. */
function nonNullAllOn<T>(values: (T | null)[], isOn: (v: T) => boolean): boolean {
  const known = values.filter((v): v is T => v !== null);
  return known.length > 0 && known.every(isOn);
}

/**
 * Conmutador genérico de un campo booleano: si todas las formas con valor no nulo están en
 * `true`, las pone en `false`; si no (incluido el caso de que todas estén en `null`, o mixto),
 * las pone en `true`. Omite las formas que ya están en el estado destino; una forma con valor
 * `null` siempre recibe parche.
 */
function toggleBooleanPatches(
  shapes: TextShape[],
  get: (s: TextShape) => boolean | null,
  set: (id: string, value: boolean) => TextPatch,
): TextPatch[] {
  const target = !nonNullAllOn(shapes.map(get), (v) => v);
  const out: TextPatch[] = [];
  for (const s of shapes) {
    if (get(s) === target) continue;
    out.push(set(s.id, target));
  }
  return out;
}

export function toggleBoldPatches(shapes: TextShape[]): TextPatch[] {
  return toggleBooleanPatches(shapes, (s) => s.bold, (id, value) => ({ id, bold: value }));
}

export function toggleItalicPatches(shapes: TextShape[]): TextPatch[] {
  return toggleBooleanPatches(shapes, (s) => s.italic, (id, value) => ({ id, italic: value }));
}

export function toggleStrikethroughPatches(shapes: TextShape[]): TextPatch[] {
  return toggleBooleanPatches(shapes, (s) => s.strikethrough, (id, value) => ({ id, strikethrough: value }));
}

export function toggleBulletsPatches(shapes: TextShape[]): TextPatch[] {
  return toggleBooleanPatches(shapes, (s) => s.bulletsVisible, (id, value) => ({ id, bulletsVisible: value }));
}

/**
 * Conmuta el subrayado: "on" es cualquier valor distinto de `"None"`. Si todas las formas con
 * valor no nulo están "on", las pone en `"None"`; si no, las pone en `"Single"`. Una forma ya
 * "on" (p. ej. `"Double"`) que se mantiene "on" no recibe parche, aunque su valor no sea
 * literalmente `"Single"`: el conmutador opera sobre el estado on/off, no sobre el estilo exacto.
 */
export function toggleUnderlinePatches(shapes: TextShape[]): TextPatch[] {
  const allOn = nonNullAllOn(shapes.map((s) => s.underline), (v) => v !== "None");
  const target = allOn ? "None" : "Single";
  const targetOn = target !== "None";
  const out: TextPatch[] = [];
  for (const s of shapes) {
    const v = s.underline;
    if (v !== null && (v !== "None") === targetOn) continue;
    out.push({ id: s.id, underline: target });
  }
  return out;
}

/** Aplica un tamaño de fuente predefinido a todas las formas cuyo tamaño no sea ya ese (incluidas las mixtas/nulas). */
export function fontSizePresetPatches(shapes: TextShape[], size: number): TextPatch[] {
  return shapes.filter((s) => s.fontSize !== size).map((s) => ({ id: s.id, fontSize: size }));
}

export const MIN_INDENT_LEVEL = 0;
export const MAX_INDENT_LEVEL = 8;

/** Suma `delta` niveles de sangría, con clamp 0..8. Omite formas con sangría mixta (null, p. ej. sin soporte de API 1.10). */
export function indentPatches(shapes: TextShape[], delta: number): TextPatch[] {
  if (!Number.isFinite(delta) || delta === 0) return [];
  const out: TextPatch[] = [];
  for (const s of shapes) {
    if (s.indentLevel === null) continue;
    const next = Math.min(MAX_INDENT_LEVEL, Math.max(MIN_INDENT_LEVEL, s.indentLevel + delta));
    if (next !== s.indentLevel) out.push({ id: s.id, indentLevel: next });
  }
  return out;
}
