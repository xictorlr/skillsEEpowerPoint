import { pngDimensions, slideSizeFromAspect, STANDARD_SLIDE_HEIGHT_PT } from "../core/png";
import type { SlideSize } from "../core/types";

export const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };
const CACHE_KEY = "current";

/**
 * Deduce el tamaño de la diapositiva seleccionada renderizándola a `STANDARD_SLIDE_HEIGHT_PT` px
 * de alto (1 px = 1 pt bajo la asunción de 540 pt, lo que evita el error de redondeo de pedir un
 * render pequeño y escalarlo). Cachea el resultado (el tamaño no cambia dentro de una
 * presentación), incluido el fallback, para no volver a renderizar en cada comando si la
 * presentación falla sistemáticamente. Fallback: 16:9 (960x540).
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
    // el catch de abajo lo convierte en el fallback 16:9.
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const image = slide.getImageAsBase64({ height: STANDARD_SLIDE_HEIGHT_PT });
    await context.sync();
    const size = slideSizeFromAspect(pngDimensions(image.value));
    cache.set(CACHE_KEY, size);
    return size;
  } catch (err) {
    console.warn(
      "[slide-tools] No se pudo deducir el tamaño de la diapositiva; se asume 16:9 (960x540).",
      err,
    );
    cache.set(CACHE_KEY, DEFAULT_SLIDE_SIZE);
    return DEFAULT_SLIDE_SIZE;
  }
}
