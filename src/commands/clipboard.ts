import type { ShapeFormat } from "../core/format";

let clipboard: ShapeFormat | null = null;

export function getClipboard(): ShapeFormat | null {
  return clipboard;
}

export function setClipboard(format: ShapeFormat | null): void {
  clipboard = format;
}

/** Solo para tests. */
export function resetClipboard(): void {
  clipboard = null;
}
