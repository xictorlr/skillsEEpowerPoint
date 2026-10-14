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

  it("returns [] when gapX or gapY is not finite", () => {
    const a = shape({ left: 0, top: 0, width: 10, height: 10 });
    const b = shape({ left: 50, top: 0, width: 10, height: 10 });
    expect(arrangeMatrix([a, b], { columns: 2, gapX: NaN, gapY: 5 })).toEqual([]);
    expect(arrangeMatrix([a, b], { columns: 2, gapX: 5, gapY: Infinity })).toEqual([]);
  });

  it("groups shapes into rows using a tolerance before sorting by left within the row", () => {
    // Tops 0.3 / 0 / 0.1 are all within ROW_TOLERANCE_PT (2pt) of each other → one row, ordered by left.
    const p = shape({ id: "p", left: 0, top: 0.3, width: 10, height: 10 });
    const q = shape({ id: "q", left: 20, top: 0, width: 10, height: 10 });
    const r = shape({ id: "r", left: 40, top: 0.1, width: 10, height: 10 });
    // columns: 1 → reading order determines the row (and thus top) of each shape.
    // originX = 0, originY = 0, cellW = cellH = 10, gap 5 → rows at top 0, 15, 30.
    expect(arrangeMatrix([p, q, r], { columns: 1, gapX: 5, gapY: 5 })).toEqual([
      { id: "p", left: 0, top: 0 },
      { id: "q", left: 0, top: 15 },
      { id: "r", left: 0, top: 30 },
    ]);
  });

  it("places every shape in a single row when columns >= count", () => {
    const x = shape({ id: "x", left: 50, top: 0, width: 10, height: 10 });
    const y = shape({ id: "y", left: 0, top: 0, width: 10, height: 10 });
    const z = shape({ id: "z", left: 25, top: 0, width: 10, height: 10 });
    // reading order (single row, by left): y(0), z(25), x(50); cols at 0, 15, 30; y stays in place.
    expect(arrangeMatrix([x, y, z], { columns: 5, gapX: 5, gapY: 5 })).toEqual([
      { id: "z", left: 15, top: 0 },
      { id: "x", left: 30, top: 0 },
    ]);
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
