import type { Rect, ShapeModel } from "../core/types";
import { supportsApi } from "./capabilities";
import { hasTextFrameByType } from "./format";
import { SHAPE_PROPS } from "./selection";

export interface DeckShape {
  slideId: string;
  slideIndex: number;
  shape: ShapeModel;
}

/** `SHAPE_PROPS`, anidado bajo `items/shapes/`, para cargar las formas de cada diapositiva del
 *  mazo en el mismo `load()` que las diapositivas (ver `readDeckShapes`). */
const NESTED_SHAPE_PROPS = SHAPE_PROPS.replace(/items\//g, "items/shapes/items/");
const DECK_PROPS = `items/id,${NESTED_SHAPE_PROPS}`;

/**
 * Lee todas las formas del mazo, diapositiva por diapositiva. Un solo `sync()`: la carga de formas
 * (`items/shapes/items/...`) va anidada bajo la carga de diapositivas (ver `DECK_PROPS`).
 */
export async function readDeckShapes(
  context: PowerPoint.RequestContext,
): Promise<{ shapes: DeckShape[]; proxies: Map<string, PowerPoint.Shape> }> {
  const slides = context.presentation.slides;
  slides.load(DECK_PROPS);
  await context.sync();

  const shapes: DeckShape[] = [];
  const proxies = new Map<string, PowerPoint.Shape>();
  slides.items.forEach((slide, slideIndex) => {
    for (const s of slide.shapes.items) {
      proxies.set(s.id, s);
      shapes.push({
        slideId: slide.id,
        slideIndex,
        shape: {
          id: s.id,
          name: s.name,
          type: String(s.type),
          left: s.left,
          top: s.top,
          width: s.width,
          height: s.height,
          rotation: s.rotation,
          zOrder: s.zOrderPosition,
          isGroupChild: (s.level ?? 0) > 0,
        },
      });
    }
  });
  return { shapes, proxies };
}

/** Encola `textFrame.load("hasText")` en las formas con marco de texto (según tipo, como en
 *  `format.ts`). El llamador hace `sync()`. */
export function loadHasText(proxies: Map<string, PowerPoint.Shape>, ids: string[]): void {
  for (const id of ids) {
    const shape = proxies.get(id);
    if (shape && hasTextFrameByType(String(shape.type))) shape.textFrame.load("hasText");
  }
}

/** Borra formas por id (`shape.delete()`). No hace `sync()`. */
export function deleteShapes(proxies: Map<string, PowerPoint.Shape>, ids: string[]): void {
  for (const id of ids) {
    proxies.get(id)?.delete();
  }
}

/** Encola `font.load("name")` en las formas con marco de texto; el llamador hace `sync()`; el
 *  nombre de fuente se lee después con `readFontName`. */
export function loadFontNames(proxies: Map<string, PowerPoint.Shape>, ids: string[]): void {
  for (const id of ids) {
    const shape = proxies.get(id);
    if (shape && hasTextFrameByType(String(shape.type))) shape.textFrame.textRange.font.load("name");
  }
}

/** Lee el nombre de fuente cargado previamente por `loadFontNames`. */
export function readFontName(shape: PowerPoint.Shape): string | null {
  return shape.textFrame.textRange.font.name ?? null;
}

/**
 * Añade una diapositiva al final del mazo y la mueve justo después de la diapositiva en `index`
 * (0-based), es decir, a la posición `index + 1`. `slides.add()` no devuelve la diapositiva creada,
 * así que hace falta un sync para poder contarlas (`getCount()`) y otro para leer ese conteo e
 * indexar la última con `getItemAt` (dos syncs en total). `moveTo(index + 1)` queda encolado sin un
 * sync final: el llamador decide cuándo sincronizar (p. ej. junto con otras escrituras posteriores,
 * como el contenido del nuevo cuadro de texto).
 */
export async function addSlideAfter(context: PowerPoint.RequestContext, index: number): Promise<PowerPoint.Slide> {
  const slides = context.presentation.slides;
  slides.add();
  await context.sync();

  const count = slides.getCount();
  await context.sync();

  const slide = slides.getItemAt(count.value - 1);
  slide.moveTo(index + 1);
  return slide;
}

/**
 * Añade un cuadro de texto a `slide` con el rectángulo y estilo dados. Solo escribe las
 * propiedades de `style` que vienen provistas: `color: null` no se escribe (se deja el color por
 * defecto del tema). No hace `sync()`.
 */
export function addTextBox(
  slide: PowerPoint.Slide,
  text: string,
  rect: Rect,
  style: { size: number; bold: boolean; color: string | null },
): PowerPoint.Shape {
  const shape = slide.shapes.addTextBox(text, { left: rect.left, top: rect.top, width: rect.width, height: rect.height });
  const font = shape.textFrame.textRange.font;
  font.size = style.size;
  font.bold = style.bold;
  if (style.color) font.color = style.color;
  shape.textFrame.wordWrap = true;
  shape.textFrame.autoSizeSetting = "AutoSizeNone" as PowerPoint.ShapeAutoSize;
  shape.textFrame.verticalAlignment = "Middle" as PowerPoint.TextVerticalAlignment;
  return shape;
}

/**
 * Exporta las diapositivas `ids` como una presentación independiente en base64 (requiere
 * PowerPointApi 1.10). Un solo `sync()`. Devuelve `null` si la API no está disponible.
 */
export async function exportSlides(context: PowerPoint.RequestContext, ids: string[]): Promise<string | null> {
  if (!supportsApi("1.10")) return null;
  const result = context.presentation.slides.exportAsBase64Presentation(ids);
  await context.sync();
  return result.value;
}

/** Crea una presentación nueva en una ventana aparte, opcionalmente a partir de un base64 existente. */
export function createPresentation(base64?: string): Promise<void> {
  return PowerPoint.createPresentation(base64);
}
