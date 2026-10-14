import type { SelectionContext, ShapeModel, SlideSize } from "../core/types";
import { getSlideSize } from "./slideSize";

export interface ReadSelectionDeps {
  slideSizeCache: Map<string, SlideSize>;
}

export interface SelectionRead {
  selection: SelectionContext;
  proxies: Map<string, PowerPoint.Shape>;
}

const SHAPE_PROPS = "items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/rotation,items/zOrderPosition,items/level";

/** Carga las formas seleccionadas y el tamaño de diapositiva. Hace un sync. */
export async function readSelection(context: PowerPoint.RequestContext, deps: ReadSelectionDeps): Promise<SelectionRead> {
  const collection = context.presentation.getSelectedShapes();
  collection.load(SHAPE_PROPS);
  await context.sync();

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

  const slide = await getSlideSize(context, deps.slideSizeCache);
  return { selection: { shapes, slide }, proxies };
}
