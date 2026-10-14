import { describe, expect, it } from "vitest";
import { arrangeMatrix } from "../../src/core/geometry";
import { shape } from "./fixtures";

describe("arrangeMatrix", () => {
  it("returns [] with fewer than 2 shapes or invalid columns", () => {
    const a = shape({ left: 0, top: 0, width: 10, height: 10 });
    const b = shape({ left: 50, top: 0, width: 10, height: 10 });
    expect(arrangeMatrix([a], { columns: 2, gapX: 5, gapY: 5 })).toEqual([]);
    expect(arrangeMatrix([a, b], { columns: 0, gapX: 5, gapY: 5 })).toEqual([]);
  });

  it("lays shapes out in reading order from the selection's top-left, with equal column/row sizes", () => {
    // Orden de lectura (por top, luego left): c (0,0), a (0,100), b (60,20), d (60,200)
    const a = shape({ id: "a", left: 100, top: 0, width: 40, height: 20 });
    const b = shape({ id: "b", left: 20, top: 60, width: 30, height: 30 });
    const c = shape({ id: "c", left: 0, top: 0, width: 20, height: 10 });
    const d = shape({ id: "d", left: 200, top: 60, width: 10, height: 10 });
    // columna = ancho máximo 40, fila = alto máximo 30, gap 10 → cols x: 0, 50 ; rows y: 0, 40
    expect(arrangeMatrix([a, b, c, d], { columns: 2, gapX: 10, gapY: 10 })).toEqual([
      { id: "a", left: 50, top: 0 },
      { id: "b", left: 0, top: 40 },
      { id: "d", left: 50, top: 40 },
    ]);
  });
});
