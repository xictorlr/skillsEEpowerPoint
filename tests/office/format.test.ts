import { describe, expect, it } from "vitest";
import { applyFormatPatch, hasTextFrameByType, loadFonts, toShapeFormat } from "../../src/office/format";
import { MockShape } from "./mockPowerPoint";

describe("hasTextFrameByType", () => {
  it("is true for shapes that can hold text and false otherwise", () => {
    for (const t of ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"]) expect(hasTextFrameByType(t)).toBe(true);
    for (const t of ["Image", "Line", "Table", "Chart", "Group", "SmartArt", "Media", "Unsupported"]) expect(hasTextFrameByType(t)).toBe(false);
  });
});

describe("loadFonts / toShapeFormat", () => {
  it("loads fonts only on text shapes and maps the format, normalizing colors", () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20,
      fill: { type: "Solid", foregroundColor: "ffa500", transparency: 0.25 },
      font: { name: "Arial", size: 14, color: "#112233", bold: true, italic: null, underline: "None" } });
    const pic = new MockShape({ id: "p", type: "Image", left: 0, top: 0, width: 10, height: 20, font: null });
    loadFonts([box, pic] as never);
    expect(box.fontLoaded).toBe(true);
    expect(toShapeFormat(box as never)).toEqual({
      id: "a",
      fill: { type: "Solid", color: "#FFA500", transparency: 0.25 },
      line: { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 },
      font: { name: "Arial", size: 14, color: "#112233", bold: true, italic: null, underline: "None" },
      size: { width: 10, height: 20 },
    });
    expect(toShapeFormat(pic as never).font).toBeNull();
  });

  it("maps unknown fill types to Other and NoFill to NoFill", () => {
    const g = new MockShape({ id: "g", left: 0, top: 0, width: 1, height: 1, fill: { type: "Gradient", foregroundColor: "#000000", transparency: null } });
    const n = new MockShape({ id: "n", left: 0, top: 0, width: 1, height: 1, fill: { type: "NoFill", foregroundColor: "#000000", transparency: null } });
    expect(toShapeFormat(g as never).fill.type).toBe("Other");
    expect(toShapeFormat(n as never).fill.type).toBe("NoFill");
  });
});

describe("applyFormatPatch", () => {
  it("writes solid color via setSolidColor, NoFill via clear, and other fields directly", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20 });
    applyFormatPatch(s as never, { id: "a", fill: { type: "Solid", color: "#00FF00", transparency: 0.5 }, line: { visible: true, color: "#0000FF", weight: 2 }, font: { size: 18, bold: true }, width: 30 });
    expect(s.formatWrites).toEqual([
      { path: "fill.setSolidColor", value: "#00FF00" },
      { path: "fill.transparency", value: 0.5 },
      { path: "lineFormat.color", value: "#0000FF" },
      { path: "lineFormat.weight", value: 2 },
      { path: "lineFormat.visible", value: true },
      { path: "font.size", value: 18 },
      { path: "font.bold", value: true },
    ]);
    expect(s.writes).toEqual([{ prop: "width", value: 30 }]);
  });

  it("clears the fill for NoFill and never touches font on shapes without text", () => {
    const pic = new MockShape({ id: "p", type: "Image", left: 0, top: 0, width: 1, height: 1, font: null });
    applyFormatPatch(pic as never, { id: "p", fill: { type: "NoFill" }, font: { size: 10 } });
    expect(pic.formatWrites).toEqual([{ path: "fill.clear", value: null }]);
  });

  it("writes only lineFormat.visible when the patch only toggles visibility", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20 });
    applyFormatPatch(s as never, { id: "a", line: { visible: false } });
    expect(s.formatWrites).toEqual([{ path: "lineFormat.visible", value: false }]);
  });

  it("does not write fill.transparency when clearing to NoFill, even if transparency is also in the patch", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20 });
    applyFormatPatch(s as never, { id: "a", fill: { type: "NoFill", transparency: 0.5 } });
    expect(s.formatWrites).toEqual([{ path: "fill.clear", value: null }]);
  });
});
