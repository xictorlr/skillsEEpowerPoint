import type { Rect } from "../core/types";

/**
 * Resumen en vivo de la selección actual: cuántas formas hay seleccionadas, qué diapositiva está
 * activa (y en qué posición, de cuántas) y la caja envolvente de la selección. Lo consume el panel
 * para la sección Navigation (etiqueta "N / M") y el inspector de la sección Expert.
 */
export interface SelectionSummary {
  /** Cantidad de formas seleccionadas. */
  count: number;
  /** Id de la diapositiva activa, o `null` si no se pudo determinar (p. ej. ninguna diapositiva). */
  slideId: string | null;
  /** Posición (0-based) de la diapositiva activa; 0 si no se pudo determinar. */
  slideIndex: number;
  /** Cantidad total de diapositivas de la presentación. */
  slideCount: number;
  /** Caja envolvente (unión) de las formas seleccionadas, o `null` si la selección está vacía. */
  bbox: Rect | null;
}

/** Rectángulo de unión de `items`, o `null` si está vacío. Igual que `boundingBox` en
 *  `core/expert.ts`, pero sobre los campos sueltos que carga esta lectura (sin fabricar un
 *  `ShapeModel` completo para cada forma). */
function union(items: ReadonlyArray<Pick<Rect, "left" | "top" | "width" | "height">>): Rect | null {
  if (items.length === 0) return null;
  const left = Math.min(...items.map((s) => s.left));
  const top = Math.min(...items.map((s) => s.top));
  const right = Math.max(...items.map((s) => s.left + s.width));
  const bottom = Math.max(...items.map((s) => s.top + s.height));
  return { left, top, width: right - left, height: bottom - top };
}

/**
 * Lee el resumen de la selección actual con un único `sync()`: geometría de las formas
 * seleccionadas, la diapositiva activa y el total de diapositivas (para poder derivar el índice).
 */
export async function readSelectionSummary(context: PowerPoint.RequestContext): Promise<SelectionSummary> {
  const selectedShapes = context.presentation.getSelectedShapes();
  selectedShapes.load("items/left,items/top,items/width,items/height");
  const selectedSlides = context.presentation.getSelectedSlides();
  selectedSlides.load("items/id");
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();

  const ids = slides.items.map((s) => s.id);
  const activeId = selectedSlides.items[0]?.id ?? null;
  const slideIndex = activeId !== null ? Math.max(0, ids.indexOf(activeId)) : 0;

  return {
    count: selectedShapes.items.length,
    slideId: activeId,
    slideIndex,
    slideCount: ids.length,
    bbox: union(selectedShapes.items),
  };
}
