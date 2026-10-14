import type { CommandOptions } from "../commands/types";
import { DEFAULT_MATRIX, sanitizeMatrix } from "./sanitize";

const KEY = "slide-tools.options";

export const DEFAULT_OPTIONS: CommandOptions = { matrix: DEFAULT_MATRIX };

export function loadOptions(): CommandOptions {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return DEFAULT_OPTIONS;
    const parsed = JSON.parse(raw) as { matrix?: unknown };
    return { matrix: sanitizeMatrix(parsed.matrix) };
  } catch {
    return DEFAULT_OPTIONS;
  }
}

export function saveOptions(options: CommandOptions): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(options));
  } catch {
    /* almacenamiento no disponible: se ignora */
  }
}
