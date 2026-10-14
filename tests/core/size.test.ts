import { describe, expect, it } from "vitest";
import { matchSize, stretchToEdge } from "../../src/core/geometry";
import { shape } from "./fixtures";

describe("matchSize", () => {
  const ref = shape({ id: "ref", left: 0, top: 0, width: 200, height: 100 });
  const a = shape({ id: "a", left: 300, top: 0, width: 50, height: 40 });
  const b = shape({ id: "b", left: 600, top: 0, width: 300, height: 150 });

  it("returns [] with fewer than 2 shapes", () => {
    expect(matchSize([ref], "width", "first")).toEqual([]);
  });

  it("matches width to the first shape, keeping top-left", () => {
    expect(matchSize([ref, a, b], "width", "first")).toEqual([
      { id: "a", width: 200 },
      { id: "b", width: 200 },
    ]);
  });

  it("matches height and both", () => {
    expect(matchSize([ref, a], "height", "first")).toEqual([{ id: "a", height: 100 }]);
    expect(matchSize([ref, a], "both", "first")).toEqual([{ id: "a", width: 200, height: 100 }]);
  });

  it("matches to the largest in the selection, including the first", () => {
    expect(matchSize([ref, a, b], "width", "max")).toEqual([
      { id: "ref", width: 300 },
      { id: "a", width: 300 },
    ]);
  });

  it("matches to the smallest in the selection", () => {
    expect(matchSize([ref, a, b], "height", "min")).toEqual([
      { id: "ref", height: 40 },
      { id: "b", height: 40 },
    ]);
  });
});

describe("stretchToEdge", () => {
  const ref = shape({ id: "ref", left: 100, top: 100, width: 400, height: 200 }); // right 500, bottom 300
  const a = shape({ id: "a", left: 200, top: 150, width: 50, height: 30 });

  it("stretches right edge to the reference right edge, keeping left fixed", () => {
    expect(stretchToEdge([ref, a], "right")).toEqual([{ id: "a", width: 300 }]);
  });

  it("stretches left edge to the reference left edge, keeping right fixed", () => {
    expect(stretchToEdge([ref, a], "left")).toEqual([{ id: "a", left: 100, width: 150 }]);
  });

  it("stretches top and bottom", () => {
    expect(stretchToEdge([ref, a], "top")).toEqual([{ id: "a", top: 100, height: 80 }]);
    expect(stretchToEdge([ref, a], "bottom")).toEqual([{ id: "a", height: 150 }]);
  });

  it("returns [] when the result would have non-positive size", () => {
    const outside = shape({ id: "o", left: 600, top: 0, width: 50, height: 30 }); // left beyond ref.right
    expect(stretchToEdge([ref, outside], "right")).toEqual([]);
  });

  it("returns [] for left/top when the result would have non-positive size", () => {
    // right edge (-30) is already to the left of ref.left (100) → width would be <= 0
    const leftOfRef = shape({ id: "lo", left: -50, top: 150, width: 20, height: 30 });
    expect(stretchToEdge([ref, leftOfRef], "left")).toEqual([]);
    // bottom edge (-30) is already above ref.top (100) → height would be <= 0
    const aboveRef = shape({ id: "ao", left: 200, top: -50, width: 20, height: 20 });
    expect(stretchToEdge([ref, aboveRef], "top")).toEqual([]);
  });

  it("returns [] for top when the shape is already flush with the reference's top edge", () => {
    const flush = shape({ id: "flush", left: 250, top: 100, width: 60, height: 80 });
    expect(stretchToEdge([ref, flush], "top")).toEqual([]);
  });
});
