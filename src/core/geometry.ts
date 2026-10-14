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

export type SizeDim = "width" | "height" | "both";
export type SizeMode = "first" | "max" | "min";
export type Edge = "left" | "right" | "top" | "bottom";

function pickTarget(shapes: ShapeModel[], key: "width" | "height", mode: SizeMode): number {
  if (mode === "first") return shapes[0]?.[key] ?? 0;
  const values = shapes.map((s) => s[key]);
  return mode === "max" ? Math.max(...values) : Math.min(...values);
}

/** Iguala ancho/alto. mode=first: al primero (no se toca). max/min: al mayor/menor de toda la selección. */
export function matchSize(shapes: ShapeModel[], dim: SizeDim, mode: SizeMode): Patch[] {
  if (shapes.length < 2) return [];
  const targets = shapes.slice(mode === "first" ? 1 : 0);
  const w = dim !== "height" ? pickTarget(shapes, "width", mode) : undefined;
  const h = dim !== "width" ? pickTarget(shapes, "height", mode) : undefined;
  const patches: Patch[] = [];
  for (const s of targets) {
    const patch: Patch = { id: s.id };
    if (w !== undefined && !near(w, s.width)) patch.width = w;
    if (h !== undefined && !near(h, s.height)) patch.height = h;
    if (patch.width !== undefined || patch.height !== undefined) patches.push(patch);
  }
  return patches;
}

/** Estira el borde indicado de cada forma hasta el mismo borde de la referencia; el borde opuesto no se mueve. */
export function stretchToEdge(shapes: ShapeModel[], edge: Edge): Patch[] {
  if (shapes.length < 2) return [];
  const [ref, ...rest] = shapes;
  if (!ref) return [];
  const refRight = ref.left + ref.width;
  const refBottom = ref.top + ref.height;
  const patches: Patch[] = [];
  for (const s of rest) {
    let patch: Patch | undefined;
    if (edge === "right") {
      const width = refRight - s.left;
      if (width > 0 && !near(width, s.width)) patch = { id: s.id, width };
    } else if (edge === "left") {
      const width = s.left + s.width - ref.left;
      if (width > 0 && !near(ref.left, s.left)) patch = { id: s.id, left: ref.left, width };
    } else if (edge === "bottom") {
      const height = refBottom - s.top;
      if (height > 0 && !near(height, s.height)) patch = { id: s.id, height };
    } else {
      const height = s.top + s.height - ref.top;
      if (height > 0 && !near(ref.top, s.top)) patch = { id: s.id, top: ref.top, height };
    }
    if (patch) patches.push(patch);
  }
  return patches;
}
