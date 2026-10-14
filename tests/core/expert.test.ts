import { describe, expect, it } from "vitest";
import { applyBoundingBox, boundingBox, filterByKind, filterByName, kindOf, numberedNames, zOrderApplyOrder } from "../../src/core/expert";
import { shape } from "./fixtures";

describe("kindOf / filterByKind", () => {
  it("maps PowerPoint types to kinds", () => {
    expect(kindOf("TextBox")).toBe("text"); expect(kindOf("GeometricShape")).toBe("shape"); expect(kindOf("Image")).toBe("image");
    expect(kindOf("Line")).toBe("line"); expect(kindOf("Table")).toBe("table"); expect(kindOf("Group")).toBe("group");
    expect(kindOf("Chart")).toBe("other"); expect(kindOf("Placeholder")).toBe("text");
  });
  it("maps Graphic to image and leaves unknown/unsupported types as other", () => {
    expect(kindOf("Graphic")).toBe("image");
    expect(kindOf("SmartArt")).toBe("other");
    expect(kindOf("Unsupported")).toBe("other");
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
  it("matches a fully uppercase accented filter", () => {
    const a = shape({ id: "a", name: "título principal", left: 0, top: 0, width: 1, height: 1 });
    expect(filterByName([a], "TÍTULO").map((s) => s.id)).toEqual(["a"]);
  });
  it("trims the filter; whitespace-only matches nothing", () => {
    const a = shape({ id: "a", name: "Título principal", left: 0, top: 0, width: 1, height: 1 });
    expect(filterByName([a], "  titulo  ").map((s) => s.id)).toEqual(["a"]);
    expect(filterByName([a], "   ")).toEqual([]);
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
  it("uses the bare zero-padded number (no leading space) when the prefix trims to empty", () => {
    const a = shape({ id: "a", left: 0, top: 0, width: 1, height: 1 });
    expect(numberedNames([a], "   ")).toEqual([{ id: "a", name: "1" }]);
  });
  it("pads to the digit count of 10+ shapes and skips those already correctly named", () => {
    const total = 12;
    const shapes = Array.from({ length: total }, (_, i) =>
      shape({ id: `s${i}`, name: `X ${String(i + 1).padStart(2, "0")}`, left: i, top: 0, width: 1, height: 1 }),
    );
    expect(numberedNames(shapes, "X")).toEqual([]);
  });
  it("does not confuse an unpadded existing name with the padded computed name (collision)", () => {
    const total = 10;
    const shapes = Array.from({ length: total }, (_, i) =>
      shape({
        id: `s${i}`,
        name: i === 0 ? "Caja 1" : i === 9 ? "Caja 10" : `Shape ${i}`,
        left: i,
        top: 0,
        width: 1,
        height: 1,
      }),
    );
    const result = numberedNames(shapes, "Caja");
    // reading-order index 0 is named "Caja 1" but the computed name is "Caja 01" (2 digits for 10 shapes) -> must rename
    expect(result.find((r) => r.id === "s0")).toEqual({ id: "s0", name: "Caja 01" });
    // reading-order index 9 already has the exact padded computed name "Caja 10" -> skipped
    expect(result.find((r) => r.id === "s9")).toBeUndefined();
  });
  it("treats shapes within the row tolerance as the same row (ordered by left), beyond it as separate rows (ordered by top)", () => {
    const sameRow = [
      shape({ id: "a", left: 10, top: 0, width: 1, height: 1 }),
      shape({ id: "b", left: 0, top: 1.5, width: 1, height: 1 }),
    ];
    expect(numberedNames(sameRow, "R").map((r) => r.id)).toEqual(["b", "a"]);

    const differentRows = [
      shape({ id: "c", left: 10, top: 0, width: 1, height: 1 }),
      shape({ id: "d", left: 0, top: 3, width: 1, height: 1 }),
    ];
    expect(numberedNames(differentRows, "R").map((r) => r.id)).toEqual(["c", "d"]);
  });
});

describe("zOrderApplyOrder", () => {
  it("applies bottom-most first for BringToFront/SendBackward and top-most first for SendToBack/BringForward", () => {
    const low = shape({ id: "low", zOrder: 1, left: 0, top: 0, width: 1, height: 1 });
    const high = shape({ id: "high", zOrder: 5, left: 0, top: 0, width: 1, height: 1 });
    expect(zOrderApplyOrder([high, low], "BringToFront").map((s) => s.id)).toEqual(["low", "high"]);
    expect(zOrderApplyOrder([low, high], "SendToBack").map((s) => s.id)).toEqual(["high", "low"]);
  });
  it("orders descending by zOrder for BringForward and ascending for SendBackward", () => {
    const low = shape({ id: "low", zOrder: 1, left: 0, top: 0, width: 1, height: 1 });
    const mid = shape({ id: "mid", zOrder: 3, left: 0, top: 0, width: 1, height: 1 });
    const high = shape({ id: "high", zOrder: 5, left: 0, top: 0, width: 1, height: 1 });
    expect(zOrderApplyOrder([low, mid, high], "BringForward").map((s) => s.id)).toEqual(["high", "mid", "low"]);
    expect(zOrderApplyOrder([high, mid, low], "SendBackward").map((s) => s.id)).toEqual(["low", "mid", "high"]);
  });
});

describe("boundingBox / applyBoundingBox", () => {
  const a = shape({ id: "a", left: 10, top: 10, width: 100, height: 50 });
  const b = shape({ id: "b", left: 60, top: 40, width: 100, height: 50 });
  it("computes the union rect", () => {
    expect(boundingBox([a, b])).toEqual({ left: 10, top: 10, width: 150, height: 80 });
    expect(boundingBox([])).toBeNull();
  });
  it("returns the shape's own rect for a single shape", () => {
    expect(boundingBox([a])).toEqual({ left: 10, top: 10, width: 100, height: 50 });
  });
  it("moves all shapes when only left/top change", () => {
    expect(applyBoundingBox([a, b], { left: 20, top: 10, width: 150, height: 80 })).toEqual([{ id: "a", left: 20 }, { id: "b", left: 70 }]);
  });
  it("scales positions and sizes proportionally when width/height change", () => {
    expect(applyBoundingBox([a, b], { left: 10, top: 10, width: 300, height: 80 })).toEqual([
      { id: "a", width: 200 }, { id: "b", left: 110, width: 200 },
    ]);
  });
  it("ignores non-finite or negative targets", () => {
    expect(applyBoundingBox([a], { left: NaN, top: 10, width: 0, height: 50 })).toEqual([]);
    expect(applyBoundingBox([a], { left: 10, top: 10, width: -5, height: 50 })).toEqual([]);
  });
  it("scales a single shape in place", () => {
    expect(applyBoundingBox([a], { left: 10, top: 10, width: 200, height: 100 })).toEqual([
      { id: "a", width: 200, height: 100 },
    ]);
  });
  it("translates a degenerate axis (zero-width line) when the target size matches it", () => {
    const line = shape({ id: "line", type: "Line", left: 10, top: 10, width: 0, height: 50 });
    expect(applyBoundingBox([line], { left: 20, top: 10, width: 0, height: 50 })).toEqual([{ id: "line", left: 20 }]);
  });
  it("rejects a positive target size on a degenerate axis", () => {
    const line = shape({ id: "line", type: "Line", left: 10, top: 10, width: 0, height: 50 });
    expect(applyBoundingBox([line], { left: 10, top: 10, width: 50, height: 50 })).toEqual([]);
  });
});
