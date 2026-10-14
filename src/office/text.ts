import type { TextPatch, TextShape } from "../core/text";
import type { FormatCaps } from "./format";

export const TEXT_FRAME_PROPS = "autoSizeSetting,wordWrap,verticalAlignment,leftMargin,rightMargin,topMargin,bottomMargin";
export const TEXT_RANGE_PROPS = "text";
export const PARAGRAPH_PROPS = "horizontalAlignment";

/** Encola la carga de texto en las formas con marco de texto (según caps.font). Requiere un sync posterior. */
export function loadTextProps(shapes: PowerPoint.Shape[], caps: Map<string, FormatCaps>): void {
  for (const s of shapes) {
    if (!caps.get(s.id)?.font) continue;
    s.textFrame.load(TEXT_FRAME_PROPS);
    s.textFrame.textRange.load(TEXT_RANGE_PROPS);
    s.textFrame.textRange.font.load("size");
    s.textFrame.textRange.paragraphFormat.load(PARAGRAPH_PROPS);
  }
}

export function toTextShape(shape: PowerPoint.Shape): TextShape {
  const tf = shape.textFrame;
  const tr = tf.textRange;
  return {
    id: shape.id,
    text: tr.text,
    fontSize: tr.font.size ?? null,
    horizontalAlignment: tr.paragraphFormat.horizontalAlignment == null ? null : String(tr.paragraphFormat.horizontalAlignment),
    verticalAlignment: String(tf.verticalAlignment),
    autoSize: String(tf.autoSizeSetting),
    wordWrap: tf.wordWrap,
    margins: { left: tf.leftMargin, right: tf.rightMargin, top: tf.topMargin, bottom: tf.bottomMargin },
  };
}

/** Escribe un parche de texto. No hace sync. */
export function applyTextPatch(shape: PowerPoint.Shape, p: TextPatch): void {
  const tf = shape.textFrame;
  if (p.text !== undefined) tf.textRange.text = p.text;
  if (p.fontSize !== undefined) tf.textRange.font.size = p.fontSize;
  if (p.horizontalAlignment !== undefined) tf.textRange.paragraphFormat.horizontalAlignment = p.horizontalAlignment as PowerPoint.ParagraphHorizontalAlignment;
  if (p.verticalAlignment !== undefined) tf.verticalAlignment = p.verticalAlignment as PowerPoint.TextVerticalAlignment;
  if (p.autoSize !== undefined) tf.autoSizeSetting = p.autoSize as PowerPoint.ShapeAutoSize;
  if (p.wordWrap !== undefined) tf.wordWrap = p.wordWrap;
  if (p.margins) {
    tf.leftMargin = p.margins.left;
    tf.rightMargin = p.margins.right;
    tf.topMargin = p.margins.top;
    tf.bottomMargin = p.margins.bottom;
  }
}
