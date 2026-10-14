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
