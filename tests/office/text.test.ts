import { describe, expect, it } from "vitest";
import { applyTextPatch, loadTextProps, toTextShape } from "../../src/office/text";
import { MockContext, MockShape } from "./mockPowerPoint";

describe("loadTextProps / toTextShape", () => {
  it("loads frame, text range and paragraph props only on shapes with font caps and maps them", () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
    const pic = new MockShape({ id: "p", type: "Image", left: 0, top: 0, width: 10, height: 10, font: null });
    new MockContext([box, pic], "");
    loadTextProps([box, pic] as never, new Map([["a", { fill: true, line: true, font: true }], ["p", { fill: true, line: true, font: false }]]));
    expect(box.textFrameLoaded && box.textRangeLoaded).toBe(true);
    expect(toTextShape(box as never)).toEqual({
      id: "a", text: "hola mundo", fontSize: 18, horizontalAlignment: "Left", verticalAlignment: "Top",
      autoSize: "AutoSizeNone", wordWrap: true, margins: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
    });
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
      horizontalAlignment: "Center",
      verticalAlignment: "Middle",
      autoSize: "AutoSizeShapeToFitText",
      wordWrap: false,
      margins: { left: 0, right: 0, top: 0, bottom: 0 },
    });
    expect(s.formatWrites).toEqual([
      { path: "textRange.substring", value: { start: 0, length: 4, text: "HOLA" } },
      { path: "font.size", value: 20 },
      { path: "paragraphFormat.horizontalAlignment", value: "Center" },
      { path: "textFrame.verticalAlignment", value: "Middle" },
      { path: "textFrame.autoSizeSetting", value: "AutoSizeShapeToFitText" },
      { path: "textFrame.wordWrap", value: false },
      { path: "textFrame.leftMargin", value: 0 },
      { path: "textFrame.rightMargin", value: 0 },
      { path: "textFrame.topMargin", value: 0 },
      { path: "textFrame.bottomMargin", value: 0 },
    ]);
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
