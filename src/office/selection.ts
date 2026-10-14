import type { ShapeFormat } from "../core/format";
import type { SelectionContext, ShapeModel, SlideSize } from "../core/types";
import { FORMAT_LOAD_PROPS, loadFonts, toShapeFormat } from "./format";
import { getSlideSize } from "./slideSize";

export interface ReadSelectionDeps {
  slideSizeCache: Map<string, SlideSize>;
  /** Tamaño de diapositiva forzado (bypassa el render/caché de getSlideSize). Ver slideSize.ts. */
  slideSizeOverride?: SlideSize;
  /** Si es true, también carga y devuelve el formato (fill/line/font) de cada forma. */
  withFormat?: boolean;
}

export interface SelectionRead {
  selection: SelectionContext;
  proxies: Map<string, PowerPoint.Shape>;
  /** Vacío si `withFormat` no se pidió. */
  formats: ShapeFormat[];
}

const SHAPE_PROPS = "items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/rotation,items/zOrderPosition,items/level";

/**
 * Carga las formas seleccionadas y el tamaño de diapositiva. Sin formato hace un sync para las
 * formas; con formato hace dos syncs de lectura (el segundo para las fuentes de las formas con
 * texto). getSlideSize puede hacer otro sync adicional la primera vez.
 */
export async function readSelection(context: PowerPoint.RequestContext, deps: ReadSelectionDeps): Promise<SelectionRead> {
  const collection = context.presentation.getSelectedShapes();
  collection.load(deps.withFormat ? `${SHAPE_PROPS},${FORMAT_LOAD_PROPS}` : SHAPE_PROPS);
  await context.sync();

  if (deps.withFormat) {
    loadFonts(collection.items);
    await context.sync();
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

  const formats = deps.withFormat ? collection.items.map(toShapeFormat) : [];

  const slide = await getSlideSize(context, deps.slideSizeCache, deps.slideSizeOverride);
  return { selection: { shapes, slide }, proxies, formats };
}
