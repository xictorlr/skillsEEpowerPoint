import { afterEach, describe, expect, it, vi } from "vitest";
import { applyTextPatch, FONT_PROPS, loadTextProps, PARAGRAPH_PROPS, toTextShape } from "../../src/office/text";
import { MockContext, MockShape } from "./mockPowerPoint";

function stubOffice1_10(): void {
  vi.stubGlobal("Office", {
    context: {
      requirements: {
        isSetSupported: (name: string, version: string) => name === "PowerPointApi" && version === "1.10",
      },
    },
  });
}

describe("loadTextProps / toTextShape", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("loads frame, text range and paragraph props only on shapes with font caps and maps them", () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
    const pic = new MockShape({ id: "p", type: "Image", left: 0, top: 0, width: 10, height: 10, font: null });
    new MockContext([box, pic], "");
    loadTextProps([box, pic] as never, new Map([["a", { fill: true, line: true, font: true }], ["p", { fill: true, line: true, font: false }]]));
    expect(box.textFrameLoaded && box.textRangeLoaded).toBe(true);
    expect(box.loads).toContainEqual({ path: "font", props: FONT_PROPS });
    expect(box.loads).toContainEqual({ path: "paragraphFormat", props: PARAGRAPH_PROPS });
    // Sin soporte de API 1.10 (no hay stub de Office), no se encola `indentLevel`.
    expect(box.loads.filter((l) => l.path === "paragraphFormat")).toHaveLength(1);
    expect(toTextShape(box as never)).toEqual({
      id: "a", text: "hola mundo", fontSize: 18, horizontalAlignment: "Left", verticalAlignment: "Top",
      autoSize: "AutoSizeNone", wordWrap: true, margins: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
      bold: false, italic: false, underline: "None", strikethrough: null, bulletsVisible: true, indentLevel: null,
    });
  });

  it("maps bold/italic/underline/bulletsVisible from non-default mock values", () => {
    const box = new MockShape({
      id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10,
      font: { name: "Calibri", size: 18, color: "#000000", bold: true, italic: true, underline: "Single" },
      textProps: { bulletsVisible: false },
    });
    new MockContext([box], "");
    expect(toTextShape(box as never)).toMatchObject({ bold: true, italic: true, underline: "Single", bulletsVisible: false });
  });

  it("loads indentLevel only when the host supports PowerPointApi 1.10, and maps it", () => {
    stubOffice1_10();
    const box = new MockShape({
      id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10, textProps: { indentLevel: 2 },
    });
    new MockContext([box], "");
    loadTextProps([box] as never, new Map([["a", { fill: true, line: true, font: true }]]));
    expect(box.loads).toContainEqual({ path: "paragraphFormat", props: "indentLevel" });
    expect(toTextShape(box as never).indentLevel).toBe(2);
  });
});

describe("applyTextPatch", () => {
  it("writes sub-range edits (in array order) before font.size, then the rest", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
    // "hola mundo" -> "HOLA mundo": dos ediciones de palabra, en orden descendente por start.
    applyTextPatch(s as never, {
      id: "a",
      edits: [{ start: 0, length: 4, text: "HOLA" }],
      fontSize: 20,
      bold: true,
      italic: true,
      underline: "Single",
      horizontalAlignment: "Center",
      bulletsVisible: false,
      indentLevel: 2,
      verticalAlignment: "Middle",
      autoSize: "AutoSizeShapeToFitText",
      wordWrap: false,
      margins: { left: 0, right: 0, top: 0, bottom: 0 },
    });
    expect(s.formatWrites).toEqual([
      { path: "textRange.substring", value: { start: 0, length: 4, text: "HOLA" } },
      { path: "font.size", value: 20 },
      { path: "font.bold", value: true },
      { path: "font.italic", value: true },
      { path: "font.underline", value: "Single" },
      { path: "paragraphFormat.horizontalAlignment", value: "Center" },
      { path: "bulletFormat.visible", value: false },
      { path: "paragraphFormat.indentLevel", value: 2 },
      { path: "textFrame.verticalAlignment", value: "Middle" },
      { path: "textFrame.autoSizeSetting", value: "AutoSizeShapeToFitText" },
      { path: "textFrame.wordWrap", value: false },
      { path: "textFrame.leftMargin", value: 0 },
      { path: "textFrame.rightMargin", value: 0 },
      { path: "textFrame.topMargin", value: 0 },
      { path: "textFrame.bottomMargin", value: 0 },
    ]);
  });

  it("writes strikethrough on the font proxy (the mock doesn't track it in formatWrites, only bold/italic/underline)", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
    applyTextPatch(s as never, { id: "a", strikethrough: true });
    expect(toTextShape(s as never).strikethrough).toBe(true);
  });

  it("applies several sub-range edits in array (descending start) order", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
    // "hola mundo" -> "HOLA MUNDO": dos palabras, edición de la última primero para no desplazar
    // el offset de la primera.
    applyTextPatch(s as never, {
      id: "a",
      edits: [
        { start: 5, length: 5, text: "MUNDO" },
        { start: 0, length: 4, text: "HOLA" },
      ],
    });
    expect(s.formatWrites).toEqual([
      { path: "textRange.substring", value: { start: 5, length: 5, text: "MUNDO" } },
      { path: "textRange.substring", value: { start: 0, length: 4, text: "HOLA" } },
    ]);
  });

  it("applies a whole-text replacement edit ({start:0, length:oldText.length})", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10 }); // "hola mundo", 10 caracteres
    applyTextPatch(s as never, { id: "a", edits: [{ start: 0, length: 10, text: "otro texto" }] });
    expect(s.formatWrites).toEqual([
      { path: "textRange.substring", value: { start: 0, length: 10, text: "otro texto" } },
    ]);
    expect(toTextShape(s as never).text).toBe("otro texto");
  });
});
