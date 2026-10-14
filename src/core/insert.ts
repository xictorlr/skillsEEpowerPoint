import type { Rect, SlideSize } from "./types";

/** Margen lateral (izquierda/derecha) del área de contenido, como fracción del ancho. */
const MARGIN_SIDE = 0.08;
/** Margen superior del área de contenido, como fracción del alto. */
const MARGIN_TOP = 0.2;
/** Margen inferior del área de contenido, como fracción del alto. */
const MARGIN_BOTTOM = 0.08;

export type ContentArea = Rect;

/** Límites por defecto para los parámetros de "cantidad" (columnas, filas, chevrones, círculos…). */
const MIN_COUNT = 1;
const MAX_COUNT = 12;

/** Redondea y acota `n` al rango [min, max]. `NaN`/`Infinity` caen al mínimo. */
function clampCount(n: number, min = MIN_COUNT, max = MAX_COUNT): number {
  if (!Number.isFinite(n)) return min;
  return Math.min(max, Math.max(min, Math.round(n)));
}

/**
 * Hueco entre `n` elementos consecutivos repartidos en una dimensión de `areaDimension` puntos:
 * `slideDimension * gapRatio`, acotado para que nunca deje un tamaño negativo. Un `gapRatio` no
 * finito o negativo se trata como 0 (sin hueco), y el hueco nunca supera `areaDimension/(n-1)` (ese
 * límite deja el elemento con tamaño 0, nunca negativo). Con `n <= 1` no hay huecos que calcular.
 */
function resolveGap(slideDimension: number, gapRatio: number, areaDimension: number, n: number): number {
  if (n <= 1) return 0;
  const ratio = Number.isFinite(gapRatio) && gapRatio > 0 ? gapRatio : 0;
  return Math.min(slideDimension * ratio, areaDimension / (n - 1));
}

/**
 * Área de contenido de la diapositiva: descuenta un margen lateral de 8 % (izquierda y derecha),
 * uno superior de 20 % (deja sitio al título) y uno inferior de 8 %.
 */
export function contentArea(slide: SlideSize): ContentArea {
  return {
    left: slide.width * MARGIN_SIDE,
    top: slide.height * MARGIN_TOP,
    width: slide.width * (1 - 2 * MARGIN_SIDE),
    height: slide.height * (1 - MARGIN_TOP - MARGIN_BOTTOM),
  };
}

/**
 * `count` (acotado a 1..12) columnas de igual ancho, repartidas en el área de contenido con un
 * hueco de `gapRatio` (fracción del ancho de la diapositiva) entre columnas consecutivas. Cada
 * columna ocupa todo el alto del área. `gapRatio` no finito o negativo se trata como 0, y el hueco
 * nunca deja una columna con ancho negativo (ver `resolveGap`): con `count === 1` no hay hueco y la
 * única columna ocupa toda el área de contenido.
 */
export function columnsLayout(slide: SlideSize, count: number, gapRatio = 0.02): Rect[] {
  const n = clampCount(count);
  const area = contentArea(slide);
  const gap = resolveGap(slide.width, gapRatio, area.width, n);
  const width = (area.width - (n - 1) * gap) / n;

  const rects: Rect[] = [];
  for (let i = 0; i < n; i++) {
    rects.push({ left: area.left + i * (width + gap), top: area.top, width, height: area.height });
  }
  return rects;
}

/**
 * `rows` × `cols` (cada uno acotado a 1..12) celdas de igual tamaño, repartidas fila a fila (de
 * izquierda a derecha, de arriba a abajo) en el área de contenido, con huecos de `gapRatio`
 * (fracción del ancho/alto de la diapositiva) entre celdas. Igual que en `columnsLayout`, cada
 * hueco se acota de forma independiente (horizontal y vertical) para no dejar celdas con tamaño
 * negativo (ver `resolveGap`).
 */
export function gridLayout(slide: SlideSize, rows: number, cols: number, gapRatio = 0.02): Rect[] {
  const r = clampCount(rows);
  const c = clampCount(cols);
  const area = contentArea(slide);
  const hGap = resolveGap(slide.width, gapRatio, area.width, c);
  const vGap = resolveGap(slide.height, gapRatio, area.height, r);
  const width = (area.width - (c - 1) * hGap) / c;
  const height = (area.height - (r - 1) * vGap) / r;

  const rects: Rect[] = [];
  for (let row = 0; row < r; row++) {
    for (let col = 0; col < c; col++) {
      rects.push({
        left: area.left + col * (width + hGap),
        top: area.top + row * (height + vGap),
        width,
        height,
      });
    }
  }
  return rects;
}

/** Alto de cada chevron: 12 % del alto de la diapositiva. */
const CHEVRON_HEIGHT_RATIO = 0.12;

/**
 * `count` (acotado a 1..12) chevrones de igual ancho, en una sola fila que ocupa todo el ancho del
 * área de contenido (sin solape: lado a lado), centrada verticalmente dentro de esa área. Cada
 * chevron mide `CHEVRON_HEIGHT_RATIO` del alto de la diapositiva.
 */
export function chevronChainLayout(slide: SlideSize, count: number): Rect[] {
  const n = clampCount(count);
  const area = contentArea(slide);
  const height = slide.height * CHEVRON_HEIGHT_RATIO;
  const width = area.width / n;
  const top = area.top + (area.height - height) / 2;

  const rects: Rect[] = [];
  for (let i = 0; i < n; i++) {
    rects.push({ left: area.left + i * width, top, width, height });
  }
  return rects;
}

/** Diámetro de cada círculo numerado: 7 % del ancho de la diapositiva. */
const CIRCLE_DIAMETER_RATIO = 0.07;

/**
 * `count` (acotado a 1..12) círculos de diámetro `CIRCLE_DIAMETER_RATIO` del ancho de la
 * diapositiva, repartidos en una única fila que reparte todo el ancho del área de contenido en
 * `count` franjas iguales (cada círculo centrado en la suya), centrada verticalmente en esa área.
 * El diámetro se acota al ancho de la franja (`Math.min(diameter, slot)`) para que los círculos
 * nunca se solapen, aunque `count` sea alto en una diapositiva estrecha.
 */
export function numberedCirclesLayout(slide: SlideSize, count: number): Rect[] {
  const n = clampCount(count);
  const area = contentArea(slide);
  const slot = area.width / n;
  const diameter = Math.min(slide.width * CIRCLE_DIAMETER_RATIO, slot);
  const top = area.top + (area.height - diameter) / 2;

  const rects: Rect[] = [];
  for (let i = 0; i < n; i++) {
    rects.push({ left: area.left + i * slot + (slot - diameter) / 2, top, width: diameter, height: diameter });
  }
  return rects;
}

/** Ratio de margen de la "sticker" respecto al lado menor de la diapositiva. */
const STICKER_MARGIN = 0.03;
const STICKER_WIDTH_RATIO = 0.18;
const STICKER_HEIGHT_RATIO = 0.07;

/**
 * Rectángulo de una "sticker" (etiqueta/cinta) en la esquina superior derecha de la diapositiva:
 * 18 % × 7 % del tamaño de la diapositiva, con un margen uniforme (igual en ambos ejes) de
 * `Math.min(slide.width, slide.height) * STICKER_MARGIN` respecto a los bordes — así el margen no
 * se deforma en diapositivas muy apaisadas o muy verticales.
 */
export function stickerRect(slide: SlideSize): Rect {
  const margin = Math.min(slide.width, slide.height) * STICKER_MARGIN;
  const width = slide.width * STICKER_WIDTH_RATIO;
  const height = slide.height * STICKER_HEIGHT_RATIO;
  return {
    left: slide.width - margin - width,
    top: margin,
    width,
    height,
  };
}

/** Lado (como fracción del lado menor de la diapositiva) de una forma básica centrada. */
const BASIC_SHAPE_SIDE_RATIO = 0.25;

/** Rectángulo cuadrado centrado en la diapositiva para una forma básica: el lado mide 25 % del
 *  lado menor (ancho o alto) de la diapositiva. */
export const BASIC_SHAPE_RECT = (slide: SlideSize): Rect => {
  const side = Math.min(slide.width, slide.height) * BASIC_SHAPE_SIDE_RATIO;
  return { left: (slide.width - side) / 2, top: (slide.height - side) / 2, width: side, height: side };
};
