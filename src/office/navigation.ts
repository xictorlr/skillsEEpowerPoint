import type { ZOrderOp } from "../core/expert";
import type { ShapeModel } from "../core/types";
import { SHAPE_PROPS } from "./selection";

export interface SlideNav {
  /** Id de la diapositiva activa (primer item de `getSelectedSlides()`). */
  slideId: string;
  /** Posición (0-based) de `slideId` dentro de `ids`. */
  index: number;
  /** Cantidad total de diapositivas. */
  count: number;
  /** Ids de todas las diapositivas, en orden. */
  ids: string[];
}

/**
 * Lee la navegación de diapositivas: todas las diapositivas (para contarlas y poder saltar por
 * índice) y cuál está activa. Un solo `sync()`. Lanza si la presentación no tiene diapositivas.
 *
 * Fallback a la primera diapositiva (`index: 0`): si `getSelectedSlides()` no devuelve ninguna
 * diapositiva (nada seleccionado) o devuelve un id que ya no está en `ids` (p. ej. una diapositiva
 * borrada entre la selección y esta lectura), se asume activa la primera diapositiva del deck en
 * vez de devolver un `index` de -1.
 */
export async function readSlideNav(context: PowerPoint.RequestContext): Promise<SlideNav> {
  const slides = context.presentation.slides;
  slides.load("items/id");
  const selectedSlides = context.presentation.getSelectedSlides();
  selectedSlides.load("items/id");
  await context.sync();

  const ids = slides.items.map((s) => s.id);
  if (ids.length === 0) throw new Error("No hay diapositivas.");

  const selectedId = selectedSlides.items[0]?.id;
  const index = selectedId !== undefined ? ids.indexOf(selectedId) : -1;
  if (index >= 0) return { slideId: selectedId!, index, count: ids.length, ids };

  // `ids[0]` está garantizado por el chequeo de arriba (ids.length > 0), pero TS no lo infiere al
  // indexar un array con `noUncheckedIndexedAccess`.
  return { slideId: ids[0]!, index: 0, count: ids.length, ids };
}

/**
 * Selecciona la diapositiva en `index` (0..count-1) usando `setSelectedSlides([id])`.
 * Devuelve `false` sin seleccionar nada si `index` está fuera de rango. No hace `sync()`: eso
 * corresponde al llamador.
 */
export function selectSlideByIndex(context: PowerPoint.RequestContext, nav: SlideNav, index: number): boolean {
  if (index < 0 || index >= nav.count) return false;
  // El rango ya se validó arriba (0 <= index < nav.count === nav.ids.length).
  context.presentation.setSelectedSlides([nav.ids[index]!]);
  return true;
}

/**
 * Lee las formas de la diapositiva `slideId` como `ShapeModel[]` (geometría, tipo y z-order).
 * Un solo `sync()`.
 */
export async function readSlideShapes(context: PowerPoint.RequestContext, slideId: string): Promise<ShapeModel[]> {
  const shapes = context.presentation.slides.getItem(slideId).shapes;
  shapes.load(SHAPE_PROPS);
  await context.sync();

  return shapes.items.map((s) => ({
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
  }));
}

/** Selecciona formas de la diapositiva `slideId` por id. No hace `sync()`. */
export function selectShapes(context: PowerPoint.RequestContext, slideId: string, ids: string[]): void {
  context.presentation.slides.getItem(slideId).setSelectedShapes(ids);
}

/** Agrupa la selección actual (`getSelectedShapes().group()`). Requiere 2+ formas seleccionadas
 *  (si no, Office.js lanza `GeneralException` al hacer `sync()`). No hace `sync()`. */
export function groupSelected(context: PowerPoint.RequestContext): void {
  context.presentation.getSelectedShapes().group();
}

/** Subconjunto de `PowerPoint.Shape` necesario para desagrupar, renombrar y reordenar (z-order).
 *  Permite mocks en tests (igual que `WritableShape` en `apply.ts`). */
export interface ShapeOpsProxy {
  name: string;
  readonly group: { ungroup(): void };
  setZOrder(op: ZOrderOp): void;
}

/** Desagrupa cada forma de `groupIds` (`shape.type === "Group"`). No hace `sync()`. */
export function ungroupShapes(proxies: Map<string, ShapeOpsProxy>, groupIds: string[]): void {
  for (const id of groupIds) {
    proxies.get(id)?.group.ungroup();
  }
}

/** Renombra formas por id. No hace `sync()`. */
export function renameShapes(proxies: Map<string, ShapeOpsProxy>, names: Array<{ id: string; name: string }>): void {
  for (const { id, name } of names) {
    const shape = proxies.get(id);
    if (shape) shape.name = name;
  }
}

/**
 * Aplica `op` de z-order a cada forma de `orderedIds`, en ese orden (ya decidido por el llamador,
 * ver `zOrderApplyOrder` en `core/expert.ts`: Office.js no tiene operación "por lote"). No hace `sync()`.
 */
export function applyZOrder(proxies: Map<string, ShapeOpsProxy>, orderedIds: string[], op: ZOrderOp): void {
  for (const id of orderedIds) {
    proxies.get(id)?.setZOrder(op);
  }
}
