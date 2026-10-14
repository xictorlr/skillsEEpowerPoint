import type { SlideSize } from "./types";

export interface PixelSize {
  width: number;
  height: number;
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
export const STANDARD_SLIDE_HEIGHT_PT = 540;

/** Lee ancho/alto del chunk IHDR de un PNG en base64. */
export function pngDimensions(base64: string): PixelSize {
  const bytes = Uint8Array.from(atob(base64.slice(0, 64)), (c) => c.charCodeAt(0));
  const isPng = PNG_SIGNATURE.every((b, i) => bytes[i] === b);
  if (!isPng || bytes.length < 24) throw new Error("Los datos no son un PNG");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return { width: view.getUint32(16), height: view.getUint32(20) };
}

/**
 * Convierte la proporción de una imagen de la diapositiva a puntos, asumiendo alto estándar 540 pt
 * (el usado por PowerPoint en 16:9, 4:3, A4 y Letter modernos).
 *
 * Limitación conocida: diapositivas con un alto distinto de 540 pt —p. ej. el 16:9 heredado de
 * 720x405 pt, o un tamaño personalizado— producirán un ancho incorrecto, porque esta función no
 * tiene forma de saber el alto real a partir solo del PNG. Para esos casos, quien llama a
 * `getSlideSize` puede pasar un `override` con el tamaño real y evitar este cálculo.
 */
export function slideSizeFromAspect(px: PixelSize): SlideSize {
  const width = Math.round((STANDARD_SLIDE_HEIGHT_PT * px.width) / px.height);
  return { width, height: STANDARD_SLIDE_HEIGHT_PT };
}
