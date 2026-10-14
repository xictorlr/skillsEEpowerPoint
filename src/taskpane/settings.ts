import type { CommandOptions } from "../commands/types";

const KEY = "slide-tools.options";

export const DEFAULT_OPTIONS: CommandOptions = { matrix: { columns: 3, gapX: 8, gapY: 8 } };

export function loadOptions(): CommandOptions {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return DEFAULT_OPTIONS;
    const parsed = JSON.parse(raw) as Partial<CommandOptions>;
    return { matrix: { ...DEFAULT_OPTIONS.matrix, ...parsed.matrix } };
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
