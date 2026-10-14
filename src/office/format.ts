import { normalizeHex } from "../core/color";
import type { FillFormat, FormatPatch, LineFormat, ShapeFormat } from "../core/format";

/** Propiedades de `fill`/`lineFormat` que se cargan por forma (no a nivel de colección). */
export const FILL_PROPS = "type,foregroundColor,transparency";
export const LINE_PROPS = "visible,color,weight,dashStyle,transparency";
const FONT_PROPS = "name,size,color,bold,italic,underline";
const PLACEHOLDER_PROPS = "containedType";

const TEXT_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"]);
/** Tipos de forma donde Office.js expone `fill` sin lanzar InvalidArgument. */
const FILL_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform", "Image"]);
/** Tipos de forma donde Office.js expone `lineFormat` sin lanzar InvalidArgument (incluye "Line"). */
const LINE_TYPES = new Set([...FILL_TYPES, "Line"]);

/** `shape.textFrame` lanza InvalidArgument en formas sin texto; se decide por tipo. No aplica a los
 *  marcadores de posición: para esos, ver `placeholderHasText`. */
export function hasTextFrameByType(type: string): boolean {
  return TEXT_TYPES.has(type);
}

/** `shape.fill` lanza InvalidArgument en algunos tipos de forma; se decide por tipo. */
export function supportsFill(type: string): boolean {
  return FILL_TYPES.has(type);
}

/** `shape.lineFormat` lanza InvalidArgument en algunos tipos de forma; se decide por tipo. */
export function supportsLine(type: string): boolean {
  return LINE_TYPES.has(type);
}

/**
 * Un marcador de posición ("Placeholder") solo tiene marco de texto cuando su contenido es texto
 * libre o una forma geométrica; `null`/`undefined` (contenido aún no decidido por el usuario) se
 * trata como texto. Imagen, tabla, gráfico, SmartArt, media… no tienen marco de texto aunque el
 * tipo de la forma sea "Placeholder".
 */
export function placeholderHasText(containedType: string | null | undefined): boolean {
  return containedType == null || containedType === "GeometricShape" || containedType === "TextBox";
}

/** Capacidades de formato de una forma, resueltas tras cargar su tipo (y, si aplica, `containedType`). */
export interface FormatCaps {
  fill: boolean;
  line: boolean;
  font: boolean;
}

/**
 * Un marcador de posición no expone `fill`/`lineFormat`/`textFrame` según su propio tipo ("Placeholder"),
 * sino según lo que de verdad contiene: un marcador vacío o de contenido libre aún sin decidir se
 * comporta como una forma geométrica; uno de imagen, tabla, gráfico… como esa forma. `containedType`
 * null/undefined (contenido aún no decidido) se asume forma geométrica.
 */
function placeholderInnerType(containedType: string | null | undefined): string {
  return containedType ?? "GeometricShape";
}

/**
 * Fase 2 de la lectura de formato: para cada forma, según su tipo, encola la carga de `fill`/
 * `lineFormat` y, en formas de texto que no son marcador de posición, de la fuente. Para los
 * marcadores de posición encola ÚNICAMENTE `placeholderFormat.containedType`: su capacidad real de
 * fill/línea/fuente depende de su contenido, que todavía no se conoce (ver `loadPlaceholderFormats`,
 * fase 3). Devuelve los marcadores de posición encontrados, para la fase 3. Requiere un sync posterior.
 */
export function loadFormatProps(shapes: PowerPoint.Shape[]): PowerPoint.Shape[] {
  const placeholders: PowerPoint.Shape[] = [];
  for (const s of shapes) {
    const type = String(s.type);
    if (type === "Placeholder") {
      s.placeholderFormat.load(PLACEHOLDER_PROPS);
      placeholders.push(s);
      continue;
    }
    if (supportsFill(type)) s.fill.load(FILL_PROPS);
    if (supportsLine(type)) s.lineFormat.load(LINE_PROPS);
    if (hasTextFrameByType(type)) s.textFrame.textRange.font.load(FONT_PROPS);
  }
  return placeholders;
}

/**
 * Fase 3, solo si la fase 2 encontró marcadores de posición: una vez conocido `containedType`,
 * encola fill/línea/fuente de cada uno según lo que de verdad contiene (ver `placeholderInnerType`
 * y `placeholderHasText`). Requiere un sync posterior.
 */
export function loadPlaceholderFormats(placeholders: PowerPoint.Shape[]): void {
  for (const p of placeholders) {
    const containedType = p.placeholderFormat.containedType;
    const inner = placeholderInnerType(containedType);
    if (supportsFill(inner)) p.fill.load(FILL_PROPS);
    if (supportsLine(inner)) p.lineFormat.load(LINE_PROPS);
    if (placeholderHasText(containedType)) p.textFrame.textRange.font.load(FONT_PROPS);
  }
}

/** Capacidades de formato de cada forma. Debe llamarse tras las fases 2 (y 3, si hubo). */
export function buildFormatCaps(shapes: PowerPoint.Shape[]): Map<string, FormatCaps> {
  const caps = new Map<string, FormatCaps>();
  for (const s of shapes) {
    const type = String(s.type);
    if (type === "Placeholder") {
      const containedType = s.placeholderFormat.containedType;
      const inner = placeholderInnerType(containedType);
      caps.set(s.id, { fill: supportsFill(inner), line: supportsLine(inner), font: placeholderHasText(containedType) });
    } else {
      caps.set(s.id, { fill: supportsFill(type), line: supportsLine(type), font: hasTextFrameByType(type) });
    }
  }
  return caps;
}

function color(v: string | null | undefined): string | null {
  return v == null ? null : normalizeHex(v);
}

/** Lee el formato de una forma. `caps` decide qué partes se leen (null si la forma no las admite). */
export function toShapeFormat(shape: PowerPoint.Shape, caps: FormatCaps): ShapeFormat {
  let fill: FillFormat | null = null;
  if (caps.fill) {
    const fillType = String(shape.fill.type);
    fill = {
      type: fillType === "Solid" ? "Solid" : fillType === "NoFill" ? "NoFill" : "Other",
      color: color(shape.fill.foregroundColor),
      transparency: shape.fill.transparency ?? null,
    };
  }

  let line: LineFormat | null = null;
  if (caps.line) {
    const lf = shape.lineFormat;
    line = {
      visible: lf.visible ?? null,
      color: color(lf.color),
      weight: lf.weight ?? null,
      dashStyle: lf.dashStyle == null ? null : String(lf.dashStyle),
      transparency: lf.transparency ?? null,
    };
  }

  let font: ShapeFormat["font"] = null;
  if (caps.font) {
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

/** Escribe un parche de formato. `caps` decide qué partes se escriben. No hace sync. */
export function applyFormatPatch(shape: PowerPoint.Shape, patch: FormatPatch, caps: FormatCaps): void {
  if (patch.fill && caps.fill) {
    if (patch.fill.type === "NoFill") shape.fill.clear();
    else if (patch.fill.color) shape.fill.setSolidColor(patch.fill.color);
    if (patch.fill.type !== "NoFill" && patch.fill.transparency != null) shape.fill.transparency = patch.fill.transparency;
  }
  if (patch.line && caps.line) {
    const l = patch.line;
    // Si se muestra la línea, `visible` va primero (los demás campos solo tienen efecto con la
    // línea visible); si se oculta, `visible` va al final (ocultarla antes puede resetear los
    // demás campos en Office.js).
    if (l.visible === true) shape.lineFormat.visible = true;
    if (l.color) shape.lineFormat.color = l.color;
    if (l.weight != null) shape.lineFormat.weight = l.weight;
    if (l.dashStyle) shape.lineFormat.dashStyle = l.dashStyle as PowerPoint.ShapeLineDashStyle;
    if (l.transparency != null) shape.lineFormat.transparency = l.transparency;
    if (l.visible === false) shape.lineFormat.visible = false;
  }
  if (patch.font && caps.font) {
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
