import { pngDimensions, slideSizeFromAspect } from "../core/png";
import type { SlideSize } from "../core/types";

export const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };
const CACHE_KEY = "current";

/**
 * Deduce el tamaño de la diapositiva seleccionada renderizándola a 54 px de alto.
 * Cachea el resultado (el tamaño no cambia dentro de una presentación). Fallback: 16:9.
 */
export async function getSlideSize(
  context: PowerPoint.RequestContext,
  cache: Map<string, SlideSize>,
): Promise<SlideSize> {
  const cached = cache.get(CACHE_KEY);
  if (cached) return cached;
  try {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const image = slide.getImageAsBase64({ height: 54 });
    await context.sync();
    const size = slideSizeFromAspect(pngDimensions(image.value));
    cache.set(CACHE_KEY, size);
    return size;
  } catch {
    return DEFAULT_SLIDE_SIZE;
  }
}
