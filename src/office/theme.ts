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
 * Las entradas cuyo valor no se puede normalizar a "#RRGGBB" se descartan. Hace un único sync.
 */
export async function getThemeColors(context: PowerPoint.RequestContext): Promise<ThemeColor[] | null> {
  if (!supportsApi("1.10")) return null;
  try {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const scheme = slide.themeColorScheme;
    const results = new Map(THEME_COLOR_NAMES.map((name) => [name, scheme.getThemeColor(name)] as const));
    await context.sync();

    const colors: ThemeColor[] = [];
    for (const name of THEME_COLOR_NAMES) {
      const hex = normalizeHex(results.get(name)!.value);
      if (hex) colors.push({ name, hex });
    }
    return colors;
  } catch {
    return null;
  }
}
