import { pngDimensions, slideSizeFromAspect, STANDARD_SLIDE_HEIGHT_PT } from "../core/png";
import type { SlideSize } from "../core/types";

export const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };
const CACHE_KEY = "current";

/** Evita inundar la consola: el aviso de fallback se emite una sola vez por sesión. */
let warned = false;

/**
 * Deduce el tamaño de la diapositiva seleccionada renderizándola a `STANDARD_SLIDE_HEIGHT_PT` px
 * de alto (1 px = 1 pt bajo la asunción de 540 pt, lo que evita el error de redondeo de pedir un
 * render pequeño y escalarlo). Cachea el resultado exitoso (el tamaño no cambia dentro de una
 * presentación).
 *
 * El fallback 16:9 (960x540) NO se cachea: un fallo puede ser transitorio (p. ej. ningún slide
 * seleccionado en ese momento), y cachearlo congelaría un tamaño incorrecto para el resto de la
 * sesión. El aviso por consola de ese fallback sí se limita a una sola vez por sesión para no
 * inundar la consola en comandos repetidos.
 *
 * Limitación conocida: se asume alto de diapositiva de 540 pt (ver `slideSizeFromAspect`). Para
 * presentaciones con otro alto (16:9 heredado a 720x405 pt, tamaños personalizados, etc.), pasa
 * `override` con el tamaño real; se devuelve tal cual, sin renderizar ni tocar la caché.
 */
export async function getSlideSize(
  context: PowerPoint.RequestContext,
  cache: Map<string, SlideSize>,
  override?: SlideSize,
): Promise<SlideSize> {
  if (override) return override;

  const cached = cache.get(CACHE_KEY);
  if (cached) return cached;
  try {
    // Si no hay ninguna diapositiva seleccionada, getItemAt(0) lanza al hacer sync();
    // el catch de abajo lo convierte en el fallback 16:9 (sin cachear: puede ser transitorio).
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const image = slide.getImageAsBase64({ height: STANDARD_SLIDE_HEIGHT_PT });
    await context.sync();
    const size = slideSizeFromAspect(pngDimensions(image.value));
    cache.set(CACHE_KEY, size);
    return size;
  } catch (err) {
    if (!warned) {
      warned = true;
      console.warn(
        "[slide-tools] No se pudo deducir el tamaño de la diapositiva; se asume 16:9 (960x540).",
        err,
      );
    }
    return DEFAULT_SLIDE_SIZE;
  }
}
