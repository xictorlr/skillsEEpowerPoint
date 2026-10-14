import type { Patch } from "../core/types";

/** Subconjunto de PowerPoint.Shape que escribimos. Permite mocks en tests. */
export interface WritableShape {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** Escribe en los proxies solo los campos presentes en cada parche. No hace sync. Descarta valores no finitos (NaN, Infinity) para nunca escribir geometría inválida. */
export function applyPatches(proxies: Map<string, WritableShape>, patches: Patch[]): void {
  for (const p of patches) {
    const shape = proxies.get(p.id);
    if (!shape) continue;
    if (p.left !== undefined && Number.isFinite(p.left)) shape.left = p.left;
    if (p.top !== undefined && Number.isFinite(p.top)) shape.top = p.top;
    if (p.width !== undefined && Number.isFinite(p.width)) shape.width = p.width;
    if (p.height !== undefined && Number.isFinite(p.height)) shape.height = p.height;
  }
}
