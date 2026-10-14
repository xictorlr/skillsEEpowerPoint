import { getThemeColors } from "../office/theme";

/** Color de acento por defecto cuando `getThemeColors` no está disponible (sin PowerPointApi 1.10)
 *  o no encuentra "Accent1". */
export const DEFAULT_ACCENT = "#1F6FEB";

/**
 * Color de acento a usar en formas creadas por el add-in (agenda, chevrones, círculos numerados):
 * el "Accent1" del tema de la diapositiva activa si se puede leer (un sync, solo con PowerPointApi
 * 1.10; ver `office/theme.ts`), o `DEFAULT_ACCENT` en caso contrario. Nunca lanza.
 */
export async function resolveAccent(context: PowerPoint.RequestContext): Promise<string> {
  const themeColors = await getThemeColors(context);
  return themeColors?.find((c) => c.name === "Accent1")?.hex ?? DEFAULT_ACCENT;
}
