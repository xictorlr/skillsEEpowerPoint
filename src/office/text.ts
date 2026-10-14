import type { TextPatch, TextShape } from "../core/text";
import { supportsApi } from "./capabilities";
import type { FormatCaps } from "./format";

export const TEXT_FRAME_PROPS = "autoSizeSetting,wordWrap,verticalAlignment,leftMargin,rightMargin,topMargin,bottomMargin";
export const TEXT_RANGE_PROPS = "text";
export const PARAGRAPH_PROPS = "horizontalAlignment,bulletFormat/visible";
// `strikethrough` requiere PowerPointApi 1.8 (bold/italic/underline/size son 1.4); el manifiesto ya
// exige 1.8 como mínimo, así que se carga sin gating (a diferencia de `indentLevel`, que es 1.10).
export const FONT_PROPS = "size,bold,italic,underline,strikethrough";

/**
 * Encola la carga de texto en las formas con marco de texto (según caps.font). `indentLevel`
 * (PowerPointApi 1.10) solo se encola si el host lo soporta; si no, `toTextShape` devuelve
 * `indentLevel: null`. Requiere un sync posterior.
 */
export function loadTextProps(shapes: PowerPoint.Shape[], caps: Map<string, FormatCaps>): void {
  const indent = supportsApi("1.10");
  for (const s of shapes) {
    if (!caps.get(s.id)?.font) continue;
    s.textFrame.load(TEXT_FRAME_PROPS);
    s.textFrame.textRange.load(TEXT_RANGE_PROPS);
    s.textFrame.textRange.font.load(FONT_PROPS);
    s.textFrame.textRange.paragraphFormat.load(PARAGRAPH_PROPS);
    if (indent) s.textFrame.textRange.paragraphFormat.load("indentLevel");
  }
}

export function toTextShape(shape: PowerPoint.Shape): TextShape {
  const tf = shape.textFrame;
  const tr = tf.textRange;
  const pf = tr.paragraphFormat;
  return {
    id: shape.id,
    text: tr.text,
    fontSize: tr.font.size ?? null,
    horizontalAlignment: pf.horizontalAlignment == null ? null : String(pf.horizontalAlignment),
    verticalAlignment: String(tf.verticalAlignment),
    autoSize: String(tf.autoSizeSetting),
    wordWrap: tf.wordWrap,
    margins: { left: tf.leftMargin, right: tf.rightMargin, top: tf.topMargin, bottom: tf.bottomMargin },
    bold: tr.font.bold ?? null,
    italic: tr.font.italic ?? null,
    underline: tr.font.underline == null ? null : String(tr.font.underline),
    strikethrough: tr.font.strikethrough ?? null,
    bulletsVisible: pf.bulletFormat.visible ?? null,
    indentLevel: supportsApi("1.10") ? pf.indentLevel ?? null : null,
  };
}

/** Escribe un parche de texto. No hace sync. */
export function applyTextPatch(shape: PowerPoint.Shape, p: TextPatch): void {
  const tf = shape.textFrame;
  // Las ediciones vienen ordenadas de la última a la primera (`start` descendente), para que
  // escribir un sub-rango no desplace los índices de las ediciones aún pendientes. El texto se
  // escribe antes que el tamaño de fuente para que este último aplique sobre el texto final.
  for (const e of p.edits ?? []) tf.textRange.getSubstring(e.start, e.length).text = e.text;
  if (p.fontSize !== undefined) tf.textRange.font.size = p.fontSize;
  if (p.bold !== undefined) tf.textRange.font.bold = p.bold;
  if (p.italic !== undefined) tf.textRange.font.italic = p.italic;
  if (p.underline !== undefined) tf.textRange.font.underline = p.underline as PowerPoint.ShapeFontUnderlineStyle;
  if (p.strikethrough !== undefined) tf.textRange.font.strikethrough = p.strikethrough;
  if (p.horizontalAlignment !== undefined) tf.textRange.paragraphFormat.horizontalAlignment = p.horizontalAlignment as PowerPoint.ParagraphHorizontalAlignment;
  if (p.bulletsVisible !== undefined) tf.textRange.paragraphFormat.bulletFormat.visible = p.bulletsVisible;
  if (p.indentLevel !== undefined) tf.textRange.paragraphFormat.indentLevel = p.indentLevel;
  if (p.verticalAlignment !== undefined) tf.verticalAlignment = p.verticalAlignment as PowerPoint.TextVerticalAlignment;
  // autoSize/wordWrap se recalculan en el layout de PowerPoint tras la escritura: no importa su
  // orden relativo al resto de campos dentro del mismo parche.
  if (p.autoSize !== undefined) tf.autoSizeSetting = p.autoSize as PowerPoint.ShapeAutoSize;
  if (p.wordWrap !== undefined) tf.wordWrap = p.wordWrap;
  if (p.margins) {
    tf.leftMargin = p.margins.left;
    tf.rightMargin = p.margins.right;
    tf.topMargin = p.margins.top;
    tf.bottomMargin = p.margins.bottom;
  }
}
