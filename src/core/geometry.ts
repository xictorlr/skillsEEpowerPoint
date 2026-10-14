import type { Patch, Rect, ShapeModel, SlideSize } from "./types";

export type AlignEdge = "left" | "centerH" | "right" | "top" | "centerV" | "bottom";

const EPS = 0.01;

function near(a: number, b: number): boolean {
  return Math.abs(a - b) < EPS;
}

function targetLeft(shape: Rect, ref: Rect, edge: "left" | "centerH" | "right"): number {
  if (edge === "left") return ref.left;
  if (edge === "right") return ref.left + ref.width - shape.width;
  return ref.left + ref.width / 2 - shape.width / 2;
}

function targetTop(shape: Rect, ref: Rect, edge: "top" | "centerV" | "bottom"): number {
  if (edge === "top") return ref.top;
  if (edge === "bottom") return ref.top + ref.height - shape.height;
  return ref.top + ref.height / 2 - shape.height / 2;
}

/** Alinea cada forma al rectángulo `ref` por el borde indicado. */
function alignToRect(shapes: ShapeModel[], ref: Rect, edge: AlignEdge): Patch[] {
  const patches: Patch[] = [];
  for (const s of shapes) {
    if (edge === "left" || edge === "centerH" || edge === "right") {
      const left = targetLeft(s, ref, edge);
      if (!near(left, s.left)) patches.push({ id: s.id, left });
    } else {
      const top = targetTop(s, ref, edge);
      if (!near(top, s.top)) patches.push({ id: s.id, top });
    }
  }
  return patches;
}

/** Alinea las formas 2..n a la primera (referencia). */
export function alignShapes(shapes: ShapeModel[], edge: AlignEdge): Patch[] {
  if (shapes.length < 2) return [];
  const [ref, ...rest] = shapes;
  if (!ref) return [];
  return alignToRect(rest, ref, edge);
}

/** Alinea todas las formas a los bordes/centro de la diapositiva. */
export function alignToSlide(shapes: ShapeModel[], slide: SlideSize, edge: AlignEdge): Patch[] {
  if (shapes.length === 0) return [];
  return alignToRect(shapes, { left: 0, top: 0, width: slide.width, height: slide.height }, edge);
}
