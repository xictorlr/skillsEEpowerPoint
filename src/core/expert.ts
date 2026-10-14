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

/**
 * Mapea el `type` de Office.js (PowerPoint.ShapeType / tipo de forma) a una categoría de alto nivel.
 * `Placeholder` y `Callout` se tratan como "text": en la práctica casi siempre contienen texto
 * editable (títulos, cuerpos, llamadas) y es lo que el usuario espera al filtrar por "texto",
 * aunque formalmente puedan contener otros marcadores de posición.
 */
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
    case "Graphic":
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
 * El filtro se recorta (trim); un filtro vacío o solo espacios no selecciona nada
 * (evita operar sobre toda la diapositiva por accidente).
 */
export function filterByName(shapes: ShapeModel[], filter: string): ShapeModel[] {
  const trimmed = filter.trim();
  if (trimmed === "") return [];
  const target = normalizeName(trimmed);
  return shapes.filter((s) => normalizeName(s.name).includes(target));
}

/**
 * Genera nombres numerados (`${prefix} ${n}`) en orden de lectura (ver `readingOrder` en
 * geometry.ts: filas con tolerancia de `top`, luego por `left` dentro de cada fila), con el
 * padding de ceros correspondiente a la cantidad de dígitos del total de formas. El prefijo se
 * recorta (trim); si queda vacío, el nombre es el número a secas (sin espacio inicial). Las
 * formas que ya tienen exactamente el nombre calculado (con el padding correcto) se omiten del
 * resultado (no hace falta renombrarlas); un nombre sin el padding correcto (p. ej. "Caja 1"
 * cuando el total exige "Caja 01") no cuenta como ya numerado y sí se renombra.
 */
export function numberedNames(shapes: ShapeModel[], prefix: string): { id: string; name: string }[] {
  const ordered = readingOrder(shapes);
  const digits = String(ordered.length).length;
  const p = prefix.trim();
  const result: { id: string; name: string }[] = [];
  ordered.forEach((s, idx) => {
    const n = String(idx + 1).padStart(digits, "0");
    const name = p ? `${p} ${n}` : n;
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
 *
 * Nota sobre `BringForward`/`SendBackward`: son operaciones de un solo paso (mueven la forma un
 * nivel, no al extremo), así que este orden preserva el orden relativo de la selección entre sí,
 * pero no garantiza su posición final frente a formas *fuera* de la selección que queden
 * justo en el borde contiguo (un caso de límite que BringToFront/SendToBack, al ir al extremo
 * absoluto, no tienen).
 */
export function zOrderApplyOrder(shapes: ShapeModel[], op: ZOrderOp): ShapeModel[] {
  const ascending = op === "BringToFront" || op === "SendBackward";
  return [...shapes].sort((a, b) => (ascending ? a.zOrder - b.zOrder : b.zOrder - a.zOrder));
}

/**
 * Rectángulo de unión (bounding box) de la selección, o `null` si está vacía. Opera sobre el marco
 * sin rotar (left/top/width/height), igual que el resto de geometry.ts; la rotación no se tiene
 * en cuenta.
 */
export function boundingBox(shapes: ShapeModel[]): Rect | null {
  if (shapes.length === 0) return null;
  const left = Math.min(...shapes.map((s) => s.left));
  const top = Math.min(...shapes.map((s) => s.top));
  const right = Math.max(...shapes.map((s) => s.left + s.width));
  const bottom = Math.max(...shapes.map((s) => s.top + s.height));
  return { left, top, width: right - left, height: bottom - top };
}

/**
 * Factor de escala para un eje al llevar la caja envolvente de `bboxSize` a `targetSize`.
 * Devuelve `null` si el cambio no es válido: tamaño destino no finito o negativo, o un tamaño
 * destino positivo sobre un eje degenerado (bboxSize === 0, p. ej. una línea vertical u
 * horizontal) al que no se le puede aplicar un factor de escala finito. Un eje degenerado con
 * destino también 0 es válido: se trata como traslación pura en ese eje (factor 1).
 */
function axisScale(bboxSize: number, targetSize: number): number | null {
  if (!Number.isFinite(targetSize) || targetSize < 0) return null;
  if (bboxSize === 0) return targetSize === 0 ? 1 : null;
  return targetSize / bboxSize;
}

/**
 * Mueve/escala toda la selección para que su caja envolvente pase a ser `target`. Opera sobre el
 * marco sin rotar (left/top/width/height), igual que el resto de geometry.ts; la rotación no se
 * tiene en cuenta. Cada eje (X/Y) se trata por separado: si `target.width`/`height` coincide con
 * el de la caja envolvente actual, ese eje solo se traslada; si difiere, las posiciones y tamaños
 * de ese eje se escalan proporcionalmente respecto a la esquina superior izquierda de la caja
 * envolvente actual. Un eje degenerado (ancho o alto de la caja envolvente igual a 0, p. ej. una
 * única línea vertical u horizontal) solo admite traslación: un destino de tamaño 0 en ese eje es
 * válido, uno positivo no (no hay un factor de escala finito que lleve 0 a un valor positivo).
 * Devuelve [] si `target` tiene valores no finitos, tamaños negativos, un tamaño positivo sobre un
 * eje degenerado, o si la selección está vacía.
 */
export function applyBoundingBox(shapes: ShapeModel[], target: Rect): Patch[] {
  const bbox = boundingBox(shapes);
  if (!bbox) return [];
  if (!Number.isFinite(target.left) || !Number.isFinite(target.top)) return [];
  const sx = axisScale(bbox.width, target.width);
  const sy = axisScale(bbox.height, target.height);
  if (sx === null || sy === null) return [];
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
