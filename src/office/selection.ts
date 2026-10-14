import type { ShapeFormat } from "../core/format";
import type { SelectionContext, ShapeModel, SlideSize } from "../core/types";
import { buildFormatCaps, loadFormatProps, loadPlaceholderFonts, toShapeFormat, type FormatCaps } from "./format";
import { DEFAULT_SLIDE_SIZE, getSlideSize } from "./slideSize";

export interface ReadSelectionDeps {
  slideSizeCache: Map<string, SlideSize>;
  /** Tamaño de diapositiva forzado (bypassa el render/caché de getSlideSize). Ver slideSize.ts. */
  slideSizeOverride?: SlideSize;
  /** Si es true, también carga y devuelve el formato (fill/line/font) de cada forma. */
  withFormat?: boolean;
  /**
   * Si es false, omite por completo el cálculo del tamaño de diapositiva (ni render ni caché) y
   * usa `DEFAULT_SLIDE_SIZE` como valor de relleno. Los comandos de formato no leen `selection.slide`,
   * así que no hace falta pagar ese sync. Por defecto true.
   */
  withSlideSize?: boolean;
}

export interface SelectionRead {
  selection: SelectionContext;
  proxies: Map<string, PowerPoint.Shape>;
  /** Vacío si `withFormat` no se pidió. */
  formats: ShapeFormat[];
  /** Capacidades de formato por id de forma; vacío si `withFormat` no se pidió. */
  caps: Map<string, FormatCaps>;
}

const SHAPE_PROPS = "items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/rotation,items/zOrderPosition,items/level";

/**
 * Carga las formas seleccionadas y, si se pide, su formato y el tamaño de diapositiva.
 *
 * Sincronización:
 * - Sync 1: siempre. Geometría + tipo de cada forma.
 * - Sync 2: solo si `withFormat`. Carga fill/lineFormat/fuente según el tipo de cada forma (y, para
 *   los marcadores de posición, `placeholderFormat.containedType`, sin la fuente todavía).
 * - Sync 3: solo si `withFormat` y hubo al menos un marcador de posición. Con `containedType` ya
 *   conocido, carga la fuente de los que tienen texto.
 * - `getSlideSize` puede hacer un sync adicional la primera vez, salvo que `withSlideSize` sea false
 *   o se pase `slideSizeOverride`.
 */
export async function readSelection(context: PowerPoint.RequestContext, deps: ReadSelectionDeps): Promise<SelectionRead> {
  const collection = context.presentation.getSelectedShapes();
  collection.load(SHAPE_PROPS);
  await context.sync();

  let caps = new Map<string, FormatCaps>();
  if (deps.withFormat) {
    const placeholders = loadFormatProps(collection.items);
    await context.sync();
    if (placeholders.length > 0) {
      loadPlaceholderFonts(placeholders);
      await context.sync();
    }
    caps = buildFormatCaps(collection.items);
  }

  const proxies = new Map<string, PowerPoint.Shape>();
  const shapes: ShapeModel[] = collection.items.map((s) => {
    proxies.set(s.id, s);
    return {
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
    };
  });

  const emptyCaps: FormatCaps = { fill: false, line: false, font: false };
  const formats = deps.withFormat ? collection.items.map((s) => toShapeFormat(s, caps.get(s.id) ?? emptyCaps)) : [];

  const slide = deps.withSlideSize === false ? DEFAULT_SLIDE_SIZE : await getSlideSize(context, deps.slideSizeCache, deps.slideSizeOverride);
  return { selection: { shapes, slide }, proxies, formats, caps };
}
