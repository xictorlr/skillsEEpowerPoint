import {
  DEFAULT_PERSISTED,
  sanitizeColorTarget,
  sanitizeMatrix,
  sanitizePasteParts,
  sanitizeRecentColors,
  sanitizeSlideSizePreset,
  type PersistedOptions,
} from "./sanitize";

export type { PersistedOptions };

const KEY = "slide-tools.options";

/** Lee y sanea las opciones persistidas (forma cruda, para alimentar los controles del panel). */
export function loadPersistedOptions(): PersistedOptions {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return DEFAULT_PERSISTED;
    // Compatible con el formato antiguo `{ matrix }`: los campos ausentes se sanean a sus valores por defecto.
    const parsed = JSON.parse(raw) as {
      matrix?: unknown;
      slideSizePreset?: unknown;
      recentColors?: unknown;
      colorTarget?: unknown;
      pasteParts?: unknown;
    };
    return {
      matrix: sanitizeMatrix(parsed.matrix),
      slideSizePreset: sanitizeSlideSizePreset(parsed.slideSizePreset),
      recentColors: sanitizeRecentColors(parsed.recentColors),
      colorTarget: sanitizeColorTarget(parsed.colorTarget),
      pasteParts: sanitizePasteParts(parsed.pasteParts),
    };
  } catch {
    return DEFAULT_PERSISTED;
  }
}

export function saveOptions(persisted: PersistedOptions): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(persisted));
  } catch {
    /* almacenamiento no disponible: se ignora */
  }
}
