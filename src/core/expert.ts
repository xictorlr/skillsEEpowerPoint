/**
 * Lógica pura para las herramientas "Expert": filtrado por tipo/nombre, numeración de nombres,
 * orden de aplicación de operaciones de z-order y caja envolvente (mover/escalar un grupo de formas).
 * Sin dependencias de Office/DOM.
 */
import { readingOrder } from "./geometry";
import type { Patch, Rect, ShapeModel } from "./types";

const EPS = 0.01;

function near(a: number, b: number): boolean {
  return Math.abs(a - b) < EPS;
}

export type ShapeKind = "all" | "text" | "shape" | "image" | "line" | "table" | "group" | "other";

/** Mapea el `type` de Office.js (PowerPoint.ShapeType / tipo de forma) a una categoría de alto nivel. */
export function kindOf(type: string): ShapeKind {
  switch (type) {
    case "TextBox":
    case "Placeholder":
    case "Callout":
      return "text";
    case "GeometricShape":
    case "Freeform":
      return "shape";
    case "Image":
      return "image";
    case "Line":
      return "line";
    case "Table":
      return "table";
    case "Group":
      return "group";
    default:
      return "other";
  }
}

/** Filtra por categoría; "all" devuelve la selección sin cambios. */
export function filterByKind(shapes: ShapeModel[], kind: ShapeKind): ShapeModel[] {
  if (kind === "all") return shapes;
  return shapes.filter((s) => kindOf(s.type) === kind);
}

function normalizeName(value: string): string {
  return value
    .normalize("NFD")
    .replace(/\p{M}/gu, "")
    .toLocaleLowerCase("es");
}

/**
 * Filtra por subcadena del nombre, insensible a mayúsculas/minúsculas y a acentos.
 * Un filtro vacío no selecciona nada (evita operar sobre toda la diapositiva por accidente).
 */
export function filterByName(shapes: ShapeModel[], filter: string): ShapeModel[] {
  if (filter === "") return [];
  const target = normalizeName(filter);
  return shapes.filter((s) => normalizeName(s.name).includes(target));
}

/**
 * Genera nombres numerados (`${prefix} ${n}`) en orden de lectura, con el padding de ceros
 * correspondiente a la cantidad de dígitos del total de formas. Las formas que ya tienen
 * exactamente ese nombre se omiten del resultado (no hace falta renombrarlas).
 */
export function numberedNames(shapes: ShapeModel[], prefix: string): { id: string; name: string }[] {
  const ordered = readingOrder(shapes);
  const digits = String(ordered.length).length;
  const result: { id: string; name: string }[] = [];
  ordered.forEach((s, idx) => {
    const name = `${prefix} ${String(idx + 1).padStart(digits, "0")}`;
    if (s.name !== name) result.push({ id: s.id, name });
  });
  return result;
}

export type ZOrderOp = "BringToFront" | "BringForward" | "SendBackward" | "SendToBack";

/**
 * Determina el orden en que hay que aplicar una operación de z-order a cada forma para que,
 * tras aplicarla individualmente (API de Office.js, que no tiene operación "por lote"), el orden
 * relativo final de la selección se preserve (según el patrón del ejemplo de Microsoft):
 * ascendente por zOrder para BringToFront/SendBackward, descendente para SendToBack/BringForward.
 */
export function zOrderApplyOrder(shapes: ShapeModel[], op: ZOrderOp): ShapeModel[] {
  const ascending = op === "BringToFront" || op === "SendBackward";
  return [...shapes].sort((a, b) => (ascending ? a.zOrder - b.zOrder : b.zOrder - a.zOrder));
}

/** Rectángulo de unión (bounding box) de la selección, o `null` si está vacía. */
export function boundingBox(shapes: ShapeModel[]): Rect | null {
  if (shapes.length === 0) return null;
  const left = Math.min(...shapes.map((s) => s.left));
  const top = Math.min(...shapes.map((s) => s.top));
  const right = Math.max(...shapes.map((s) => s.left + s.width));
  const bottom = Math.max(...shapes.map((s) => s.top + s.height));
  return { left, top, width: right - left, height: bottom - top };
}

/**
 * Mueve/escala toda la selección para que su caja envolvente pase a ser `target`. Si solo cambian
 * left/top, es una traslación pura; si cambian width/height, las posiciones y tamaños se escalan
 * proporcionalmente respecto a la esquina superior izquierda de la caja envolvente actual.
 * Devuelve [] si `target` tiene valores no finitos o width/height no positivos, o si la selección
 * está vacía.
 */
export function applyBoundingBox(shapes: ShapeModel[], target: Rect): Patch[] {
  const bbox = boundingBox(shapes);
  if (!bbox) return [];
  if (
    !Number.isFinite(target.left) ||
    !Number.isFinite(target.top) ||
    !Number.isFinite(target.width) ||
    !Number.isFinite(target.height) ||
    target.width <= 0 ||
    target.height <= 0 ||
    bbox.width <= 0 ||
    bbox.height <= 0
  ) {
    return [];
  }
  const sx = target.width / bbox.width;
  const sy = target.height / bbox.height;
  const patches: Patch[] = [];
  for (const s of shapes) {
    const newLeft = target.left + (s.left - bbox.left) * sx;
    const newTop = target.top + (s.top - bbox.top) * sy;
    const newWidth = s.width * sx;
    const newHeight = s.height * sy;
    const patch: Patch = { id: s.id };
    if (!near(newLeft, s.left)) patch.left = newLeft;
    if (!near(newTop, s.top)) patch.top = newTop;
    if (!near(newWidth, s.width)) patch.width = newWidth;
    if (!near(newHeight, s.height)) patch.height = newHeight;
    if (patch.left !== undefined || patch.top !== undefined || patch.width !== undefined || patch.height !== undefined) {
      patches.push(patch);
    }
  }
  return patches;
}
