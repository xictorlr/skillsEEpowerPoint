import { describe, expect, it } from "vitest";
import { alignShapes, alignToSlide } from "../../src/core/geometry";
import { shape, slide16x9 } from "./fixtures";

describe("alignShapes", () => {
  const ref = shape({ id: "ref", left: 100, top: 100, width: 200, height: 100 });
  const a = shape({ id: "a", left: 400, top: 300, width: 50, height: 40 });
  const b = shape({ id: "b", left: 10, top: 20, width: 80, height: 60 });

  it("returns [] with fewer than 2 shapes", () => {
    expect(alignShapes([ref], "left")).toEqual([]);
  });

  it("aligns left edges to the reference (first shape)", () => {
    expect(alignShapes([ref, a, b], "left")).toEqual([
      { id: "a", left: 100 },
      { id: "b", left: 100 },
    ]);
  });

  it("aligns right edges", () => {
    expect(alignShapes([ref, a, b], "right")).toEqual([
      { id: "a", left: 250 },
      { id: "b", left: 220 },
    ]);
  });

  it("aligns horizontal centers", () => {
    expect(alignShapes([ref, a, b], "centerH")).toEqual([
      { id: "a", left: 175 },
      { id: "b", left: 160 },
    ]);
  });

  it("aligns top, bottom and vertical centers", () => {
    expect(alignShapes([ref, a], "top")).toEqual([{ id: "a", top: 100 }]);
    expect(alignShapes([ref, a], "bottom")).toEqual([{ id: "a", top: 160 }]);
    expect(alignShapes([ref, a], "centerV")).toEqual([{ id: "a", top: 130 }]);
  });

  it("omits shapes already aligned", () => {
    const same = shape({ id: "same", left: 100, top: 0, width: 10, height: 10 });
    expect(alignShapes([ref, same], "left")).toEqual([]);
  });
});

describe("alignToSlide", () => {
  const a = shape({ id: "a", left: 10, top: 10, width: 100, height: 50 });

  it("centers horizontally and vertically on the slide", () => {
    expect(alignToSlide([a], slide16x9, "centerH")).toEqual([{ id: "a", left: 430 }]);
    expect(alignToSlide([a], slide16x9, "centerV")).toEqual([{ id: "a", top: 245 }]);
  });

  it("aligns to slide edges", () => {
    expect(alignToSlide([a], slide16x9, "right")).toEqual([{ id: "a", left: 860 }]);
    expect(alignToSlide([a], slide16x9, "bottom")).toEqual([{ id: "a", top: 490 }]);
  });

  it("works with a single shape and returns [] for none", () => {
    expect(alignToSlide([], slide16x9, "left")).toEqual([]);
  });
});
