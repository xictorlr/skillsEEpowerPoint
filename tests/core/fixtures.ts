import type { ShapeModel, SlideSize } from "../../src/core/types";

let counter = 0;

export function shape(partial: Partial<ShapeModel> & Pick<ShapeModel, "left" | "top" | "width" | "height">): ShapeModel {
  counter += 1;
  return {
    id: partial.id ?? `s${counter}`,
    name: partial.name ?? `Shape ${counter}`,
    type: partial.type ?? "GeometricShape",
    rotation: partial.rotation ?? 0,
    zOrder: partial.zOrder ?? counter,
    isGroupChild: partial.isGroupChild ?? false,
    left: partial.left,
    top: partial.top,
    width: partial.width,
    height: partial.height,
  };
}

export const slide16x9: SlideSize = { width: 960, height: 540 };
