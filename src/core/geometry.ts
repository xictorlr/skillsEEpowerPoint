/**
 * Todas las funciones usan el marco sin rotar (left/top/width/height) tal como lo expone Office.js;
 * la rotación no se tiene en cuenta.
 */
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

/**
 * Reparte el espacio entre formas de modo uniforme. La forma con el borde inicial mínimo y la forma
 * con el borde final máximo (sobre toda la selección) permanecen fijas; el resto se ordena por su
 * centro y se reparte con huecos iguales entre esas dos. Si una sola forma cubre a todas las demás
 * (es la misma por ambos extremos), no hay nada que repartir.
 */
export function distribute(shapes: ShapeModel[], axis: Axis): Patch[] {
  if (shapes.length < 3) return [];
  const pos = axis === "h" ? "left" : "top";
  const size = axis === "h" ? "width" : "height";
  const leftmost = shapes.reduce((m, s) => (s[pos] < m[pos] ? s : m));
  const rightmost = shapes.reduce((m, s) => (s[pos] + s[size] > m[pos] + m[size] ? s : m));
  if (leftmost.id === rightmost.id) return [];
  const span = rightmost[pos] + rightmost[size] - leftmost[pos];
  const total = shapes.reduce((acc, s) => acc + s[size], 0);
  const gap = (span - total) / (shapes.length - 1);
  const middle = shapes
    .filter((s) => s.id !== leftmost.id && s.id !== rightmost.id)
    .sort((a, b) => a[pos] + a[size] / 2 - (b[pos] + b[size] / 2));
  const patches: Patch[] = [];
  let cursor = leftmost[pos] + leftmost[size] + gap;
  for (const s of middle) {
    if (!near(cursor, s[pos])) patches.push(axis === "h" ? { id: s.id, left: cursor } : { id: s.id, top: cursor });
    cursor += s[size] + gap;
  }
  return patches;
}

function swapPatch(shape: ShapeModel, newLeft: number, newTop: number): Patch | undefined {
  const patch: Patch = { id: shape.id };
  if (!near(newLeft, shape.left)) patch.left = newLeft;
  if (!near(newTop, shape.top)) patch.top = newTop;
  return patch.left !== undefined || patch.top !== undefined ? patch : undefined;
}

/** Intercambia los centros de dos formas. Solo se incluyen los campos (left/top) que realmente cambian. */
export function swapPositions(shapes: ShapeModel[]): Patch[] {
  if (shapes.length !== 2) return [];
  const [a, b] = shapes;
  if (!a || !b) return [];
  const ca = { x: a.left + a.width / 2, y: a.top + a.height / 2 };
  const cb = { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  const patches: Patch[] = [];
  const aPatch = swapPatch(a, cb.x - a.width / 2, cb.y - a.height / 2);
  if (aPatch) patches.push(aPatch);
  const bPatch = swapPatch(b, ca.x - b.width / 2, ca.y - b.height / 2);
  if (bPatch) patches.push(bPatch);
  return patches;
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

/**
 * Estira el borde indicado de cada forma hasta el mismo borde de la referencia; el borde opuesto no se mueve.
 * El borde movido se ajusta exactamente al borde de referencia, por lo que la forma también puede encogerse.
 */
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

export interface MatrixOptions {
  columns: number;
  gapX: number;
  gapY: number;
}

/** Tolerancia (en puntos) para considerar que dos formas están en la misma fila pese a pequeñas
 *  diferencias de `top` (p. ej. redondeo). */
const ROW_TOLERANCE_PT = 2;

/**
 * Calcula el orden de lectura de una selección: agrupa las formas en filas (dentro de
 * `ROW_TOLERANCE_PT` del `top` de la primera forma de la fila) y, dentro de cada fila, ordena por `left`.
 */
function readingOrder(shapes: ShapeModel[]): ShapeModel[] {
  const byTop = [...shapes].sort((a, b) => a.top - b.top);
  const rows: ShapeModel[][] = [];
  let currentRow: ShapeModel[] = [];
  let rowTop = 0;
  for (const s of byTop) {
    if (currentRow.length === 0) {
      rowTop = s.top;
    } else if (s.top - rowTop > ROW_TOLERANCE_PT) {
      rows.push(currentRow);
      currentRow = [];
      rowTop = s.top;
    }
    currentRow.push(s);
  }
  if (currentRow.length > 0) rows.push(currentRow);
  return rows.flatMap((row) => [...row].sort((a, b) => a.left - b.left));
}

/**
 * Coloca las formas en una rejilla de `columns` columnas, en orden de lectura (filas por `top` con
 * tolerancia, luego `left` dentro de cada fila), partiendo de la esquina superior izquierda de la
 * selección. Cada celda mide el máximo ancho/alto.
 */
export function arrangeMatrix(shapes: ShapeModel[], opts: MatrixOptions): Patch[] {
  if (
    shapes.length < 2 ||
    !Number.isInteger(opts.columns) ||
    opts.columns < 1 ||
    !Number.isFinite(opts.gapX) ||
    !Number.isFinite(opts.gapY)
  ) {
    return [];
  }
  const ordered = readingOrder(shapes);
  const originX = Math.min(...shapes.map((s) => s.left));
  const originY = Math.min(...shapes.map((s) => s.top));
  const cellW = Math.max(...shapes.map((s) => s.width));
  const cellH = Math.max(...shapes.map((s) => s.height));
  const patches: Patch[] = [];
  ordered.forEach((s, i) => {
    const col = i % opts.columns;
    const row = Math.floor(i / opts.columns);
    const left = originX + col * (cellW + opts.gapX);
    const top = originY + row * (cellH + opts.gapY);
    if (!near(left, s.left) || !near(top, s.top)) patches.push({ id: s.id, left, top });
  });
  return patches;
}
