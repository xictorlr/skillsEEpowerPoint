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

export type Axis = "h" | "v";

/** Reparte el espacio entre formas de modo uniforme. Las dos extremas (por posición) no se mueven. */
export function distribute(shapes: ShapeModel[], axis: Axis): Patch[] {
  if (shapes.length < 3) return [];
  const pos = axis === "h" ? "left" : "top";
  const size = axis === "h" ? "width" : "height";
  const sorted = [...shapes].sort((a, b) => a[pos] + a[size] / 2 - (b[pos] + b[size] / 2));
  const first = sorted[0];
  const last = sorted[sorted.length - 1];
  if (!first || !last) return [];
  const span = last[pos] + last[size] - first[pos];
  const total = sorted.reduce((acc, s) => acc + s[size], 0);
  const gap = (span - total) / (sorted.length - 1);
  const patches: Patch[] = [];
  let cursor = first[pos] + first[size] + gap;
  for (const s of sorted.slice(1, -1)) {
    if (!near(cursor, s[pos])) patches.push({ id: s.id, [pos]: cursor } as Patch);
    cursor += s[size] + gap;
  }
  return patches;
}

/** Intercambia los centros de dos formas. */
export function swapPositions(shapes: ShapeModel[]): Patch[] {
  if (shapes.length !== 2) return [];
  const [a, b] = shapes;
  if (!a || !b) return [];
  const ca = { x: a.left + a.width / 2, y: a.top + a.height / 2 };
  const cb = { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  return [
    { id: a.id, left: cb.x - a.width / 2, top: cb.y - a.height / 2 },
    { id: b.id, left: ca.x - b.width / 2, top: ca.y - b.height / 2 },
  ];
}
