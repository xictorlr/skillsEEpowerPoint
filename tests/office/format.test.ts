import { describe, expect, it } from "vitest";
import {
  applyFormatPatch,
  buildFormatCaps,
  hasTextFrameByType,
  loadFormatProps,
  loadPlaceholderFonts,
  placeholderHasText,
  supportsFill,
  supportsLine,
  toShapeFormat,
  type FormatCaps,
} from "../../src/office/format";
import { MockContext, MockShape } from "./mockPowerPoint";

describe("hasTextFrameByType", () => {
  it("is true for shapes that can hold text and false otherwise", () => {
    for (const t of ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"]) expect(hasTextFrameByType(t)).toBe(true);
    for (const t of ["Image", "Line", "Table", "Chart", "Group", "SmartArt", "Media", "Unsupported"]) expect(hasTextFrameByType(t)).toBe(false);
  });
});

describe("supportsFill / supportsLine", () => {
  it("includes Image for fill, and additionally Line for line", () => {
    for (const t of ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform", "Image"]) {
      expect(supportsFill(t)).toBe(true);
      expect(supportsLine(t)).toBe(true);
    }
    expect(supportsLine("Line")).toBe(true);
    expect(supportsFill("Line")).toBe(false);
    for (const t of ["Table", "Chart", "Group", "SmartArt", "Media", "Unsupported"]) {
      expect(supportsFill(t)).toBe(false);
      expect(supportsLine(t)).toBe(false);
    }
  });
});

describe("placeholderHasText", () => {
  it("treats null/undefined and GeometricShape/TextBox as text-capable, everything else as not", () => {
    expect(placeholderHasText(null)).toBe(true);
    expect(placeholderHasText(undefined)).toBe(true);
    expect(placeholderHasText("GeometricShape")).toBe(true);
    expect(placeholderHasText("TextBox")).toBe(true);
    expect(placeholderHasText("Image")).toBe(false);
    expect(placeholderHasText("Table")).toBe(false);
  });
});

describe("loadFormatProps / loadPlaceholderFonts / buildFormatCaps", () => {
  it("loads fill/line/font for a text shape and reports its capabilities", () => {
    const box = new MockShape({
      id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20,
      fill: { type: "Solid", foregroundColor: "ffa500", transparency: 0.25 },
      font: { name: "Arial", size: 14, color: "#112233", bold: true, italic: null, underline: "None" },
    });
    const placeholders = loadFormatProps([box] as never);
    expect(placeholders).toEqual([]);
    expect(box.fillLoaded).toBe(true);
    expect(box.lineLoaded).toBe(true);
    expect(box.fontLoaded).toBe(true);

    const caps = buildFormatCaps([box] as never);
    expect(caps.get("a")).toEqual({ fill: true, line: true, font: true });
    expect(toShapeFormat(box as never, caps.get("a") as FormatCaps)).toEqual({
      id: "a",
      fill: { type: "Solid", color: "#FFA500", transparency: 0.25 },
      line: { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 },
      font: { name: "Arial", size: 14, color: "#112233", bold: true, italic: null, underline: "None" },
      size: { width: 10, height: 20 },
    });
  });

  it("never touches fill/line/font/placeholderFormat for a type that supports none of them", async () => {
    const group = new MockShape({ id: "g", type: "Group", left: 0, top: 0, width: 5, height: 5, supportsFormat: false });
    const placeholders = loadFormatProps([group] as never);
    expect(placeholders).toEqual([]);
    expect(group.fillLoaded).toBe(false);
    expect(group.lineLoaded).toBe(false);
    expect(group.fontLoaded).toBe(false);

    const caps = buildFormatCaps([group] as never);
    expect(caps.get("g")).toEqual({ fill: false, line: false, font: false });
    expect(toShapeFormat(group as never, caps.get("g") as FormatCaps)).toEqual({
      id: "g",
      fill: null,
      line: null,
      font: null,
      size: { width: 5, height: 5 },
    });

    // Nada tocó fill/lineFormat, así que el mock nunca marcó el contexto como inválido.
    const ctx = new MockContext([group]);
    await ctx.sync();
  });

  it("defers the font for a placeholder until phase 3, and skips it when its content isn't text", () => {
    const imagePh = new MockShape({ id: "p1", type: "Placeholder", left: 0, top: 0, width: 1, height: 1, containedType: "Image" });
    const textPh = new MockShape({ id: "p2", type: "Placeholder", left: 0, top: 0, width: 1, height: 1, containedType: null,
      font: { name: "Arial", size: 10, color: "#000000", bold: false, italic: false, underline: "None" } });

    const placeholders = loadFormatProps([imagePh, textPh] as never);
    expect(placeholders).toEqual([imagePh, textPh]);
    // La fase 2 nunca carga la fuente de un marcador de posición directamente.
    expect(imagePh.fontLoaded).toBe(false);
    expect(textPh.fontLoaded).toBe(false);

    loadPlaceholderFonts(placeholders);
    expect(imagePh.fontLoaded).toBe(false); // contenido Image: sin marco de texto
    expect(textPh.fontLoaded).toBe(true); // contenido null: se asume texto

    const caps = buildFormatCaps([imagePh, textPh] as never);
    expect(caps.get("p1")?.font).toBe(false);
    expect(caps.get("p2")?.font).toBe(true);
    expect(toShapeFormat(imagePh as never, caps.get("p1") as FormatCaps).font).toBeNull();
    expect(toShapeFormat(textPh as never, caps.get("p2") as FormatCaps).font?.name).toBe("Arial");
  });

  it("maps unknown fill types to Other and NoFill to NoFill", () => {
    const g = new MockShape({ id: "g", left: 0, top: 0, width: 1, height: 1, fill: { type: "Gradient", foregroundColor: "#000000", transparency: null } });
    const n = new MockShape({ id: "n", left: 0, top: 0, width: 1, height: 1, fill: { type: "NoFill", foregroundColor: "#000000", transparency: null } });
    const caps: FormatCaps = { fill: true, line: true, font: true };
    expect(toShapeFormat(g as never, caps).fill?.type).toBe("Other");
    expect(toShapeFormat(n as never, caps).fill?.type).toBe("NoFill");
  });
});

describe("applyFormatPatch", () => {
  const ALL_CAPS: FormatCaps = { fill: true, line: true, font: true };

  it("writes solid color via setSolidColor, NoFill via clear, visible first when shown, and other fields directly", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20 });
    applyFormatPatch(
      s as never,
      { id: "a", fill: { type: "Solid", color: "#00FF00", transparency: 0.5 }, line: { visible: true, color: "#0000FF", weight: 2 }, font: { size: 18, bold: true }, width: 30 },
      ALL_CAPS,
    );
    expect(s.formatWrites).toEqual([
      { path: "fill.setSolidColor", value: "#00FF00" },
      { path: "fill.transparency", value: 0.5 },
      { path: "lineFormat.visible", value: true },
      { path: "lineFormat.color", value: "#0000FF" },
      { path: "lineFormat.weight", value: 2 },
      { path: "font.size", value: 18 },
      { path: "font.bold", value: true },
    ]);
    expect(s.writes).toEqual([{ prop: "width", value: 30 }]);
  });

  it("writes visible last when hiding the line, after color/weight/dashStyle/transparency", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20 });
    applyFormatPatch(s as never, { id: "a", line: { visible: false, color: "#ABCDEF", weight: 3 } }, ALL_CAPS);
    expect(s.formatWrites).toEqual([
      { path: "lineFormat.color", value: "#ABCDEF" },
      { path: "lineFormat.weight", value: 3 },
      { path: "lineFormat.visible", value: false },
    ]);
  });

  it("clears the fill for NoFill and never touches font on shapes without text", () => {
    const pic = new MockShape({ id: "p", type: "Image", left: 0, top: 0, width: 1, height: 1, font: null });
    applyFormatPatch(pic as never, { id: "p", fill: { type: "NoFill" }, font: { size: 10 } }, { fill: true, line: true, font: false });
    expect(pic.formatWrites).toEqual([{ path: "fill.clear", value: null }]);
  });

  it("writes only lineFormat.visible when the patch only toggles visibility", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20 });
    applyFormatPatch(s as never, { id: "a", line: { visible: false } }, ALL_CAPS);
    expect(s.formatWrites).toEqual([{ path: "lineFormat.visible", value: false }]);
  });

  it("does not write fill.transparency when clearing to NoFill, even if transparency is also in the patch", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20 });
    applyFormatPatch(s as never, { id: "a", fill: { type: "NoFill", transparency: 0.5 } }, ALL_CAPS);
    expect(s.formatWrites).toEqual([{ path: "fill.clear", value: null }]);
  });

  it("ignores a fill/line/font part the caps say the shape doesn't support", () => {
    const s = new MockShape({ id: "a", type: "Line", left: 0, top: 0, width: 10, height: 20, font: null });
    applyFormatPatch(s as never, { id: "a", fill: { type: "Solid", color: "#00FF00" }, line: { visible: true }, font: { size: 10 } }, { fill: false, line: true, font: false });
    expect(s.formatWrites).toEqual([{ path: "lineFormat.visible", value: true }]);
  });

  it("taints the mock context so the next sync rejects when writing to a shape PowerPoint refuses", async () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10, supportsFormat: false });
    const ctx = new MockContext([s]);
    applyFormatPatch(s as never, { id: "a", fill: { type: "Solid", color: "#00FF00" } }, { fill: true, line: false, font: false });
    await expect(ctx.sync()).rejects.toMatchObject({ code: "InvalidArgument" });
  });
});
