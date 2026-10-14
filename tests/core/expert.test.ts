import { describe, expect, it } from "vitest";
import { applyBoundingBox, boundingBox, filterByKind, filterByName, kindOf, numberedNames, zOrderApplyOrder } from "../../src/core/expert";
import { shape } from "./fixtures";

describe("kindOf / filterByKind", () => {
  it("maps PowerPoint types to kinds", () => {
    expect(kindOf("TextBox")).toBe("text"); expect(kindOf("GeometricShape")).toBe("shape"); expect(kindOf("Image")).toBe("image");
    expect(kindOf("Line")).toBe("line"); expect(kindOf("Table")).toBe("table"); expect(kindOf("Group")).toBe("group");
    expect(kindOf("Chart")).toBe("other"); expect(kindOf("Placeholder")).toBe("text");
  });
  it("filters by kind or returns all", () => {
    const a = shape({ id: "a", type: "TextBox", left: 0, top: 0, width: 1, height: 1 });
    const b = shape({ id: "b", type: "Image", left: 0, top: 0, width: 1, height: 1 });
    expect(filterByKind([a, b], "image").map((s) => s.id)).toEqual(["b"]);
    expect(filterByKind([a, b], "all").map((s) => s.id)).toEqual(["a", "b"]);
  });
});

describe("filterByName", () => {
  it("is case- and accent-insensitive substring match; empty filter matches nothing", () => {
    const a = shape({ id: "a", name: "Título principal", left: 0, top: 0, width: 1, height: 1 });
    expect(filterByName([a], "titulo").map((s) => s.id)).toEqual(["a"]);
    expect(filterByName([a], "")).toEqual([]);
  });
});

describe("numberedNames", () => {
  it("numbers in reading order with zero padding by count", () => {
    const a = shape({ id: "a", left: 100, top: 0, width: 1, height: 1 });
    const b = shape({ id: "b", left: 0, top: 0, width: 1, height: 1 });
    expect(numberedNames([a, b], "Caja")).toEqual([{ id: "b", name: "Caja 1" }, { id: "a", name: "Caja 2" }]);
    const many = Array.from({ length: 10 }, (_, i) => shape({ id: `s${i}`, left: i, top: 0, width: 1, height: 1 }));
    expect(numberedNames(many, "X")[0]?.name).toBe("X 01");
  });
  it("skips shapes that already have the target name", () => {
    const a = shape({ id: "a", name: "Caja 1", left: 0, top: 0, width: 1, height: 1 });
    expect(numberedNames([a], "Caja")).toEqual([]);
  });
});

describe("zOrderApplyOrder", () => {
  it("applies bottom-most first for BringToFront/SendBackward and top-most first for SendToBack/BringForward", () => {
    const low = shape({ id: "low", zOrder: 1, left: 0, top: 0, width: 1, height: 1 });
    const high = shape({ id: "high", zOrder: 5, left: 0, top: 0, width: 1, height: 1 });
    expect(zOrderApplyOrder([high, low], "BringToFront").map((s) => s.id)).toEqual(["low", "high"]);
    expect(zOrderApplyOrder([low, high], "SendToBack").map((s) => s.id)).toEqual(["high", "low"]);
  });
});

describe("boundingBox / applyBoundingBox", () => {
  const a = shape({ id: "a", left: 10, top: 10, width: 100, height: 50 });
  const b = shape({ id: "b", left: 60, top: 40, width: 100, height: 50 });
  it("computes the union rect", () => {
    expect(boundingBox([a, b])).toEqual({ left: 10, top: 10, width: 150, height: 80 });
    expect(boundingBox([])).toBeNull();
  });
  it("moves all shapes when only left/top change", () => {
    expect(applyBoundingBox([a, b], { left: 20, top: 10, width: 150, height: 80 })).toEqual([{ id: "a", left: 20 }, { id: "b", left: 70 }]);
  });
  it("scales positions and sizes proportionally when width/height change", () => {
    expect(applyBoundingBox([a, b], { left: 10, top: 10, width: 300, height: 80 })).toEqual([
      { id: "a", width: 200 }, { id: "b", left: 110, width: 200 },
    ]);
  });
  it("ignores non-finite or non-positive targets", () => {
    expect(applyBoundingBox([a], { left: NaN, top: 10, width: 0, height: 50 })).toEqual([]);
  });
});
