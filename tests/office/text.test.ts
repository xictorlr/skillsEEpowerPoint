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
  it("writes each present field", () => {
    const s = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
    applyTextPatch(s as never, { id: "a", text: "X", fontSize: 20, horizontalAlignment: "Center", verticalAlignment: "Middle", autoSize: "AutoSizeShapeToFitText", wordWrap: false, margins: { left: 0, right: 0, top: 0, bottom: 0 } });
    expect(s.formatWrites).toEqual([
      { path: "textRange.text", value: "X" },
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
});
