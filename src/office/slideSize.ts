import { pngDimensions, slideSizeFromAspect, STANDARD_SLIDE_HEIGHT_PT } from "../core/png";
import type { SlideSize } from "../core/types";
import { supportsApi } from "./capabilities";

export const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };
const CACHE_KEY = "current";

/** Evita inundar la consola: el aviso de fallback se emite una sola vez por sesión. */
let warned = false;

/**
 * Deduce el tamaño de la diapositiva seleccionada. Sin PowerPointApi 1.10, se renderiza la
 * diapositiva a `STANDARD_SLIDE_HEIGHT_PT` px de alto (1 px = 1 pt bajo la asunción de 540 pt, lo
 * que evita el error de redondeo de pedir un render pequeño y escalarlo) y se cachea el resultado
 * exitoso (el tamaño no cambia dentro de una presentación, y esta heurística es cara de repetir).
 *
 * El fallback 16:9 (960x540) NO se cachea: un fallo puede ser transitorio (p. ej. ningún slide
 * seleccionado en ese momento), y cachearlo congelaría un tamaño incorrecto para el resto de la
 * sesión. El aviso por consola de ese fallback sí se limita a una sola vez por sesión para no
 * inundar la consola en comandos repetidos.
 *
 * Sin PowerPointApi 1.10, limitación conocida: se asume alto de diapositiva de 540 pt (ver
 * `slideSizeFromAspect`). Para presentaciones con otro alto (16:9 heredado a 720x405 pt, tamaños
 * personalizados, etc.), pasa `override` con el tamaño real; se devuelve tal cual, sin renderizar
 * ni tocar la caché.
 *
 * Si el host soporta PowerPointApi 1.10 (ver `supportsApi`), se usa `presentation.pageSetup` para
 * obtener el tamaño exacto (en puntos) de un solo sync, sin recurrir al render PNG ni a la caché:
 * se lee en cada llamada (un sync barato) para reflejar de inmediato un cambio de Diseño > Tamaño
 * de diapositiva hecho mientras el panel está abierto. Si esos valores resultan inválidos (no
 * finitos o <= 0) o la lectura falla, se cae a la heurística PNG de abajo (que sí usa la caché).
 */
export async function getSlideSize(
  context: PowerPoint.RequestContext,
  cache: Map<string, SlideSize>,
  override?: SlideSize,
): Promise<SlideSize> {
  if (override) return override;

  if (supportsApi("1.10")) {
    try {
      const pageSetup = context.presentation.pageSetup;
      pageSetup.load("slideWidth,slideHeight");
      await context.sync();
      const { slideWidth, slideHeight } = pageSetup;
      if (Number.isFinite(slideWidth) && slideWidth > 0 && Number.isFinite(slideHeight) && slideHeight > 0) {
        return { width: slideWidth, height: slideHeight };
      }
    } catch {
      // cae a la heurística PNG de abajo (p. ej. host que anuncia 1.10 pero falla al leerlo).
    }
  }

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
