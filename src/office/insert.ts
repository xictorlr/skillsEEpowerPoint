import type { Rect } from "../core/types";

/**
 * Estilo opcional a aplicar a una forma recién creada. Solo se escriben las propiedades provistas
 * (mismo criterio que `addTextBox` en `deck.ts`): `fill`/`line` con `null` borra/oculta en vez de
 * dejar el valor por defecto del tema; `fontColor: null` (o ausente) no toca `font.color`.
 */
export interface ShapeStyle {
  /** Color de relleno (hex); `null` para quitar el relleno (`fill.clear()`). */
  fill?: string | null;
  /** Color de línea (hex); `null` para ocultar la línea (`lineFormat.visible = false`). */
  line?: string | null;
  fontSize?: number;
  bold?: boolean;
  /** Color de fuente (hex); `null`/ausente deja el color por defecto del tema. */
  fontColor?: string | null;
  /** Texto de la forma (solo tiene efecto en `addGeometric`: `addTextBoxes` ya recibe sus textos
   *  aparte, uno por rect). */
  text?: string;
  /** Centra el párrafo (`paragraphFormat.horizontalAlignment = "Center"`); pensado para chevrones,
   *  círculos numerados y la sticker. */
  centered?: boolean;
}

function rectOptions(rect: Rect): { left: number; top: number; width: number; height: number } {
  return { left: rect.left, top: rect.top, width: rect.width, height: rect.height };
}

/** Escribe `fill`/`line` de `style` sobre `shape`, si vienen provistos. */
function applyFillAndLine(shape: PowerPoint.Shape, style?: ShapeStyle): void {
  if (!style) return;
  if (style.fill !== undefined) {
    if (style.fill === null) shape.fill.clear();
    else shape.fill.setSolidColor(style.fill);
  }
  if (style.line !== undefined) {
    if (style.line === null) shape.lineFormat.visible = false;
    else shape.lineFormat.color = style.line;
  }
}

/** Escribe `font.size`/`font.bold`/`font.color` de `style` sobre `shape`, si vienen provistos. */
function applyFont(shape: PowerPoint.Shape, style?: ShapeStyle): void {
  if (!style) return;
  const font = shape.textFrame.textRange.font;
  if (style.fontSize !== undefined) font.size = style.fontSize;
  if (style.bold !== undefined) font.bold = style.bold;
  if (style.fontColor) font.color = style.fontColor;
}

/** Deja el marco de texto listo para una etiqueta corta: centrado verticalmente y con ajuste de
 *  línea; centra el párrafo si `style.centered`. */
function applyTextDefaults(shape: PowerPoint.Shape, style?: ShapeStyle): void {
  if (style?.centered) shape.textFrame.textRange.paragraphFormat.horizontalAlignment = "Center";
  shape.textFrame.verticalAlignment = "Middle";
  shape.textFrame.wordWrap = true;
}

/**
 * Añade un cuadro de texto por cada `rects[i]`/`texts[i]`, con el mismo `style` para todos. No
 * hace `sync()`.
 */
export function addTextBoxes(slide: PowerPoint.Slide, rects: Rect[], texts: string[], style?: ShapeStyle): PowerPoint.Shape[] {
  return rects.map((rect, i) => {
    const shape = slide.shapes.addTextBox(texts[i] ?? "", rectOptions(rect));
    applyFillAndLine(shape, style);
    applyFont(shape, style);
    applyTextDefaults(shape, style);
    return shape;
  });
}

/** Tipos de forma geométrica soportados por `addGeometric`. */
export type GeometricShapeType = "Rectangle" | "RoundRectangle" | "Ellipse" | "Chevron";

/**
 * Añade una forma geométrica (`type`) en `rect`, con el `style` dado (incluido `style.text`, el
 * texto de la etiqueta). No hace `sync()`.
 */
export function addGeometric(slide: PowerPoint.Slide, type: GeometricShapeType, rect: Rect, style?: ShapeStyle): PowerPoint.Shape {
  const shape = slide.shapes.addGeometricShape(type, rectOptions(rect));
  applyFillAndLine(shape, style);
  if (style?.text !== undefined) shape.textFrame.textRange.text = style.text;
  applyFont(shape, style);
  applyTextDefaults(shape, style);
  return shape;
}

/** Tipos de línea soportados por `addLine`. */
export type LineType = "Straight" | "Elbow";

/** Añade una línea (`kind`) en `rect`. No hace `sync()`. */
export function addLine(slide: PowerPoint.Slide, kind: LineType, rect: Rect): PowerPoint.Shape {
  return slide.shapes.addLine(kind, rectOptions(rect));
}

/**
 * Selecciona las formas recién creadas (`shapes`). El llamador debe haber cargado `id` en cada una
 * y sincronizado antes de llamar (igual que cualquier lectura de propiedades de una forma recién
 * creada): aquí solo se usa `shape.id`, ya disponible tras ese sync.
 */
export function selectCreated(slide: PowerPoint.Slide, shapes: PowerPoint.Shape[]): void {
  slide.setSelectedShapes(shapes.map((s) => s.id));
}
