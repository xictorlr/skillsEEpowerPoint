import { normalizeHex } from "../core/color";
import { supportsApi } from "./capabilities";

export const THEME_COLOR_NAMES = [
  "Dark1",
  "Light1",
  "Dark2",
  "Light2",
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
  "Hyperlink",
  "FollowedHyperlink",
] as const;

export type ThemeColorName = (typeof THEME_COLOR_NAMES)[number];

export interface ThemeColor {
  name: ThemeColorName;
  hex: string;
}

/**
 * Colores del tema de la diapositiva seleccionada (requiere PowerPointApi 1.10). Devuelve `null`
 * si la API no está disponible, o si falla la lectura (p. ej. ninguna diapositiva seleccionada).
 * Las entradas cuyo valor no sea una cadena normalizable a "#RRGGBB" (incluyendo `null`/`undefined`)
 * se descartan. Hace un único sync.
 */
export async function getThemeColors(context: PowerPoint.RequestContext): Promise<ThemeColor[] | null> {
  if (!supportsApi("1.10")) return null;
  try {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const scheme = slide.themeColorScheme;
    const results = THEME_COLOR_NAMES.map((name) => scheme.getThemeColor(name));
    await context.sync();

    const colors: ThemeColor[] = [];
    THEME_COLOR_NAMES.forEach((name, i) => {
      const value = results[i]?.value;
      const hex = typeof value === "string" ? normalizeHex(value) : null;
      if (hex) colors.push({ name, hex });
    });
    return colors;
  } catch {
    return null;
  }
}
