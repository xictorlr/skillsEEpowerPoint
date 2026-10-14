import type { MatrixOptions } from "../core/geometry";

/**
 * Valores por defecto de la matriz. Se usan cuando falta un campo o su valor no es válido.
 * Archivo puro (sin DOM ni localStorage) para poder testearlo bajo Node.
 */
export const DEFAULT_MATRIX: MatrixOptions = { columns: 3, gapX: 8, gapY: 8 };

function pick(value: unknown, min: number, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value >= min ? value : fallback;
}

/** Sanea un valor arbitrario (p.ej. JSON.parse de localStorage, o un <input>) a una MatrixOptions válida. */
export function sanitizeMatrix(input: unknown): MatrixOptions {
  const obj = input !== null && typeof input === "object" ? (input as Record<string, unknown>) : {};
  const columns = pick(obj.columns, 1, DEFAULT_MATRIX.columns);
  return {
    columns: Math.round(columns),
    gapX: pick(obj.gapX, 0, DEFAULT_MATRIX.gapX),
    gapY: pick(obj.gapY, 0, DEFAULT_MATRIX.gapY),
  };
}
