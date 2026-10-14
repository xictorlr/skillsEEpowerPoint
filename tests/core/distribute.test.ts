import { describe, expect, it } from "vitest";
import { distribute, swapPositions } from "../../src/core/geometry";
import { shape } from "./fixtures";

describe("distribute", () => {
  it("returns [] with fewer than 3 shapes", () => {
    const a = shape({ left: 0, top: 0, width: 10, height: 10 });
    const b = shape({ left: 100, top: 0, width: 10, height: 10 });
    expect(distribute([a, b], "h")).toEqual([]);
  });

  it("spaces shapes evenly between the outermost ones (horizontal)", () => {
    const a = shape({ id: "a", left: 0, top: 0, width: 100, height: 10 });
    const b = shape({ id: "b", left: 120, top: 0, width: 50, height: 10 });
    const c = shape({ id: "c", left: 400, top: 0, width: 100, height: 10 });
    // span 0..500, widths 250, gaps = (500-250)/2 = 125 → b.left = 225
    expect(distribute([c, a, b], "h")).toEqual([{ id: "b", left: 225 }]);
  });

  it("spaces shapes evenly (vertical) regardless of input order", () => {
    const a = shape({ id: "a", left: 0, top: 0, width: 10, height: 40 });
    const b = shape({ id: "b", left: 0, top: 50, width: 10, height: 20 });
    const c = shape({ id: "c", left: 0, top: 200, width: 10, height: 40 });
    const d = shape({ id: "d", left: 0, top: 500, width: 10, height: 100 });
    // span 0..600, heights 200, gaps = 400/3 ≈ 133.33
    const result = distribute([d, b, c, a], "v");
    expect(result).toHaveLength(2);
    expect(result[0]?.id).toBe("b");
    expect(result[0]?.top).toBeCloseTo(173.333, 2);
    expect(result[1]?.id).toBe("c");
    expect(result[1]?.top).toBeCloseTo(326.667, 2);
  });
});

describe("swapPositions", () => {
  it("swaps the centers of exactly two shapes", () => {
    const a = shape({ id: "a", left: 0, top: 0, width: 100, height: 100 });   // center 50,50
    const b = shape({ id: "b", left: 300, top: 200, width: 20, height: 40 }); // center 310,220
    expect(swapPositions([a, b])).toEqual([
      { id: "a", left: 260, top: 170 },
      { id: "b", left: 40, top: 30 },
    ]);
  });

  it("returns [] unless exactly two shapes", () => {
    const a = shape({ left: 0, top: 0, width: 1, height: 1 });
    expect(swapPositions([a])).toEqual([]);
    expect(swapPositions([a, a, a])).toEqual([]);
  });
});
