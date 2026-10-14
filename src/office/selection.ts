import type { ShapeFormat } from "../core/format";
import type { TextShape } from "../core/text";
import type { SelectionContext, ShapeModel, SlideSize } from "../core/types";
import { buildFormatCaps, loadFormatProps, loadPlaceholderFormats, toShapeFormat, type FormatCaps } from "./format";
import { DEFAULT_SLIDE_SIZE, getSlideSize } from "./slideSize";
import { loadTextProps, toTextShape } from "./text";

export interface ReadSelectionDeps {
  slideSizeCache: Map<string, SlideSize>;
  /** Tamaño de diapositiva forzado (bypassa el render/caché de getSlideSize). Ver slideSize.ts. */
  slideSizeOverride?: SlideSize;
  /** Si es true, también carga y devuelve el formato (fill/line/font) de cada forma. */
  withFormat?: boolean;
  /**
   * Si es true, también carga y devuelve el texto (texto, alineación, márgenes, autoajuste) de cada
   * forma con marco de texto. Implica cargar `caps` igual que `withFormat` (misma fase 2/3 que las
   * fuentes), aunque no se pida `withFormat`.
   */
  withText?: boolean;
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
  /** Vacío si ni `withFormat` ni `withText` se pidieron. */
  formats: ShapeFormat[];
  /** Capacidades de formato por id de forma; vacío si ni `withFormat` ni `withText` se pidieron. */
  caps: Map<string, FormatCaps>;
  /** Vacío si `withText` no se pidió. */
  texts: TextShape[];
}

const SHAPE_PROPS = "items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/rotation,items/zOrderPosition,items/level";

/**
 * Carga las formas seleccionadas y, si se pide, su formato y el tamaño de diapositiva.
 *
 * Sincronización:
 * - Sync 1: siempre. Geometría + tipo de cada forma.
 * - Sync 2: solo si `withFormat` o `withText`. Carga fill/lineFormat/fuente (y, si `withText`, el
 *   texto: ver `loadTextProps`) según el tipo de cada forma; para los marcadores de posición carga
 *   únicamente `placeholderFormat.containedType` (su fill/línea/fuente/texto dependen de lo que
 *   contienen, que todavía no se conoce).
 * - Sync 3: solo si `withFormat` o `withText`, y hubo al menos un marcador de posición. Con
 *   `containedType` ya conocido, carga fill/línea/fuente (y texto) de cada uno según lo que de
 *   verdad contiene.
 * - `getSlideSize` puede hacer un sync adicional la primera vez, salvo que `withSlideSize` sea false
 *   o se pase `slideSizeOverride`.
 */
export async function readSelection(context: PowerPoint.RequestContext, deps: ReadSelectionDeps): Promise<SelectionRead> {
  const collection = context.presentation.getSelectedShapes();
  collection.load(SHAPE_PROPS);
  await context.sync();

  // `withText` necesita las mismas capacidades (`caps.font`) que `withFormat`, cargadas en la misma
  // fase que las fuentes, para que el conteo de syncs no cambie.
  const needFormat = deps.withFormat || deps.withText;
  let caps = new Map<string, FormatCaps>();
  if (needFormat) {
    const placeholders = loadFormatProps(collection.items);
    if (deps.withText) {
      // `containedType` de los marcadores de posición aún no se cargó (se encola arriba, en
      // `loadFormatProps`): su capacidad de texto se decide en la fase 3. Para el resto, el tipo ya
      // es suficiente y se calcula sin tocar `placeholderFormat`.
      const nonPlaceholders = collection.items.filter((s) => String(s.type) !== "Placeholder");
      loadTextProps(nonPlaceholders, buildFormatCaps(nonPlaceholders));
    }
    await context.sync();
    if (placeholders.length > 0) {
      loadPlaceholderFormats(placeholders);
      if (deps.withText) loadTextProps(placeholders, buildFormatCaps(placeholders));
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
  const formats = needFormat ? collection.items.map((s) => toShapeFormat(s, caps.get(s.id) ?? emptyCaps)) : [];
  const texts = deps.withText ? collection.items.filter((s) => caps.get(s.id)?.font).map((s) => toTextShape(s)) : [];

  const slide = deps.withSlideSize === false ? DEFAULT_SLIDE_SIZE : await getSlideSize(context, deps.slideSizeCache, deps.slideSizeOverride);
  return { selection: { shapes, slide }, proxies, formats, caps, texts };
}
