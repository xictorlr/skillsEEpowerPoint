import type { Rect, ShapeModel } from "../core/types";
import { supportsApi } from "./capabilities";
import { hasTextFrameByType, placeholderHasText } from "./format";
import { SHAPE_PROPS } from "./selection";

export interface DeckShape {
  slideId: string;
  slideIndex: number;
  shape: ShapeModel;
}

/** `SHAPE_PROPS`, anidado bajo `items/shapes/`, para cargar las formas de cada diapositiva del
 *  mazo en el mismo `load()` que las diapositivas (ver `readDeckShapes`). */
const NESTED_SHAPE_PROPS = SHAPE_PROPS.replace(/items\//g, "items/shapes/items/");
export const DECK_PROPS = `items/id,${NESTED_SHAPE_PROPS}`;

/**
 * Lee todas las formas del mazo, diapositiva por diapositiva. Solo ve formas de primer nivel: las
 * formas dentro de un grupo no se visitan (igual que `readSlideShapes`/`readSelection`, que tampoco
 * bajan a los hijos de un grupo).
 *
 * Sincronización:
 * - Sync 1: siempre. Diapositivas + formas (`items/shapes/items/...`, ver `DECK_PROPS`), anidadas
 *   en el mismo `load()` que las diapositivas.
 * - Sync 2: solo si el mazo tiene algún marcador de posición ("Placeholder"), para conocer su
 *   `containedType` (de qué depende si de verdad tiene marco de texto; ver `placeholderHasText`).
 *
 * `textCapable` son los ids de las formas con marco de texto: por tipo (`hasTextFrameByType`) para
 * las que no son marcador de posición, y según `containedType` para las que sí lo son.
 */
export async function readDeckShapes(
  context: PowerPoint.RequestContext,
): Promise<{ shapes: DeckShape[]; proxies: Map<string, PowerPoint.Shape>; textCapable: Set<string> }> {
  const slides = context.presentation.slides;
  slides.load(DECK_PROPS);
  await context.sync();

  const shapes: DeckShape[] = [];
  const proxies = new Map<string, PowerPoint.Shape>();
  const allShapes: PowerPoint.Shape[] = [];
  slides.items.forEach((slide, slideIndex) => {
    for (const s of slide.shapes.items) {
      proxies.set(s.id, s);
      allShapes.push(s);
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

  const placeholders = allShapes.filter((s) => String(s.type) === "Placeholder");
  if (placeholders.length > 0) {
    for (const p of placeholders) p.placeholderFormat.load("containedType");
    await context.sync();
  }

  const textCapable = new Set<string>();
  for (const s of allShapes) {
    if (String(s.type) === "Placeholder") {
      if (placeholderHasText(s.placeholderFormat.containedType)) textCapable.add(s.id);
    } else if (hasTextFrameByType(String(s.type))) {
      textCapable.add(s.id);
    }
  }

  return { shapes, proxies, textCapable };
}

/**
 * Como `readDeckShapes`, pero para una única diapositiva (`slideId`): mismos dos syncs (el segundo
 * solo si esa diapositiva tiene marcadores de posición), pero sin leer ni un sync de más por
 * cargar formas de las demás diapositivas del mazo. Para comandos que solo operan sobre la
 * diapositiva activa (p. ej. `check-empty-text-slide`).
 */
export async function readSlideShapesWithProxies(
  context: PowerPoint.RequestContext,
  slideId: string,
): Promise<{ shapes: ShapeModel[]; proxies: Map<string, PowerPoint.Shape>; textCapable: Set<string> }> {
  const shapesCollection = context.presentation.slides.getItem(slideId).shapes;
  shapesCollection.load(SHAPE_PROPS);
  await context.sync();

  const shapes: ShapeModel[] = [];
  const proxies = new Map<string, PowerPoint.Shape>();
  const allShapes: PowerPoint.Shape[] = [];
  for (const s of shapesCollection.items) {
    proxies.set(s.id, s);
    allShapes.push(s);
    shapes.push({
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
    });
  }

  const placeholders = allShapes.filter((s) => String(s.type) === "Placeholder");
  if (placeholders.length > 0) {
    for (const p of placeholders) p.placeholderFormat.load("containedType");
    await context.sync();
  }

  const textCapable = new Set<string>();
  for (const s of allShapes) {
    if (String(s.type) === "Placeholder") {
      if (placeholderHasText(s.placeholderFormat.containedType)) textCapable.add(s.id);
    } else if (hasTextFrameByType(String(s.type))) {
      textCapable.add(s.id);
    }
  }

  return { shapes, proxies, textCapable };
}

/** Encola `textFrame.load("hasText")` en las formas de `ids` que admiten marco de texto, según
 *  `textCapable` (ver `readDeckShapes`). El llamador hace `sync()`. */
export function loadHasText(proxies: Map<string, PowerPoint.Shape>, ids: string[], textCapable: Set<string>): void {
  for (const id of ids) {
    if (!textCapable.has(id)) continue;
    proxies.get(id)?.textFrame.load("hasText");
  }
}

/** Borra formas por id (`shape.delete()`). No hace `sync()`. */
export function deleteShapes(proxies: Map<string, PowerPoint.Shape>, ids: string[]): void {
  for (const id of ids) {
    proxies.get(id)?.delete();
  }
}

/** Encola `font.load("name")` en las formas de `ids` que admiten marco de texto, según
 *  `textCapable` (ver `readDeckShapes`); el llamador hace `sync()`; el nombre de fuente se lee
 *  después con `readFontName`. */
export function loadFontNames(proxies: Map<string, PowerPoint.Shape>, ids: string[], textCapable: Set<string>): void {
  for (const id of ids) {
    if (!textCapable.has(id)) continue;
    proxies.get(id)?.textFrame.textRange.font.load("name");
  }
}

/** Lee el nombre de fuente cargado previamente por `loadFontNames`. */
export function readFontName(shape: PowerPoint.Shape): string | null {
  return shape.textFrame.textRange.font.name ?? null;
}

/**
 * Añade una diapositiva al final del mazo y la mueve justo después de la diapositiva en `index`
 * (0-based), es decir, a la posición `index + 1`. Si se da `layout` (ver `findBlankLayout`), la
 * nueva diapositiva nace con ese layout (de ese slide master) en vez del layout por defecto. Un
 * solo `sync()`: `slides.add()` no devuelve la diapositiva creada, pero `getCount()` se puede
 * encolar en el mismo lote (ambas operaciones se ejecutan en el servidor en el orden en que se
 * encolaron, así que el conteo ya refleja la diapositiva añadida) y leerse tras ese único sync para
 * indexar la última con `getItemAt`. `moveTo(index + 1)` (requiere PowerPointApi 1.8, el mínimo del
 * manifiesto) queda encolado sin un sync final: el llamador decide cuándo sincronizar (p. ej. junto
 * con otras escrituras posteriores, como el contenido del nuevo cuadro de texto).
 */
export async function addSlideAfter(
  context: PowerPoint.RequestContext,
  index: number,
  layout?: { slideMasterId: string; layoutId: string },
): Promise<PowerPoint.Slide> {
  const slides = context.presentation.slides;
  if (layout) slides.add({ slideMasterId: layout.slideMasterId, layoutId: layout.layoutId });
  else slides.add();
  const count = slides.getCount();
  await context.sync();

  const slide = slides.getItemAt(count.value - 1);
  slide.moveTo(index + 1);
  return slide;
}

/**
 * Como `addSlideAfter`, pero para `count` diapositivas a la vez (p. ej. `agenda-create`, una por
 * sección): encola los `count` `slides.add()` (todas con el mismo `layout`, si se da) y un único
 * `getCount()`, un solo `sync()` de alta, y resuelve cada diapositiva nueva con `getItemAt` sobre
 * las últimas `count` posiciones (el conteo ya las refleja, igual que en `addSlideAfter`). Cada
 * `moveTo(index + 1 + i)` queda encolado en orden (i: 0..count-1), sin sync: el llamador decide
 * cuándo sincronizar (p. ej. junto con el contenido de los nuevos cuadros de texto). Devuelve las
 * diapositivas en el mismo orden en que se insertan (index+1, index+2, …).
 */
export async function addSlidesAfter(
  context: PowerPoint.RequestContext,
  index: number,
  count: number,
  layout?: { slideMasterId: string; layoutId: string },
): Promise<PowerPoint.Slide[]> {
  const slides = context.presentation.slides;
  for (let i = 0; i < count; i++) {
    if (layout) slides.add({ slideMasterId: layout.slideMasterId, layoutId: layout.layoutId });
    else slides.add();
  }
  const total = slides.getCount();
  await context.sync();

  const added: PowerPoint.Slide[] = [];
  for (let i = 0; i < count; i++) {
    const slide = slides.getItemAt(total.value - count + i);
    slide.moveTo(index + 1 + i);
    added.push(slide);
  }
  return added;
}

/**
 * Busca el primer layout "en blanco" entre los layouts del primer slide master, y devuelve también
 * el id de ese slide master: `slides.add({ slideMasterId, layoutId })` necesita ambos para ubicar el
 * layout sin ambigüedad (el mismo `layoutId` podría existir en otro master). Un solo `sync()`.
 * Devuelve `undefined` si ningún layout coincide, o si falla la lectura (p. ej. la presentación no
 * tiene slide masters) — nunca lanza.
 */
export async function findBlankLayout(
  context: PowerPoint.RequestContext,
): Promise<{ slideMasterId: string; layoutId: string } | undefined> {
  try {
    const master = context.presentation.slideMasters.getItemAt(0);
    master.load("id");
    const layouts = master.layouts;
    layouts.load("items/id,items/name");
    await context.sync();

    const blank = layouts.items.find((l) => /blank|en blanco|vac[ií]o/i.test(l.name));
    return blank ? { slideMasterId: master.id, layoutId: blank.id } : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Borra los marcadores de posición ("Placeholder") de `slide`: red de seguridad para cuando
 * `findBlankLayout` no encontró ningún layout en blanco y la diapositiva añadida trae los
 * marcadores del layout por defecto. Un solo `sync()` para conocer el tipo de cada forma; los
 * `delete()` quedan encolados sin un sync final (el llamador decide cuándo sincronizar). Devuelve
 * cuántos se encolaron para borrar.
 */
export async function deletePlaceholders(context: PowerPoint.RequestContext, slide: PowerPoint.Slide): Promise<number> {
  const shapes = slide.shapes;
  shapes.load("items/id,items/type");
  await context.sync();

  let count = 0;
  for (const s of shapes.items) {
    if (String(s.type) === "Placeholder") {
      s.delete();
      count += 1;
    }
  }
  return count;
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
  shape.textFrame.autoSizeSetting = "AutoSizeNone";
  shape.textFrame.verticalAlignment = "Middle";
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
