import { describe, expect, it } from "vitest";
import { addGeometric, addLine, addTextBoxes, selectCreated } from "../../src/office/insert";
import { MockContext, MockShape, MockSlide } from "./mockPowerPoint";

const rect = { left: 10, top: 20, width: 100, height: 50 };

/** Diapositiva vacía, lista para `addTextBoxes`/`addGeometric`/`addLine`/`selectCreated`. */
function newSlide(): { ctx: MockContext; slide: MockSlide } {
  const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
  return { ctx, slide: ctx.slidesList[0]! };
}

describe("addTextBoxes", () => {
  it("adds one text box per rect/text, recording each call and the shared style", () => {
    const { ctx, slide } = newSlide();
    const rects = [rect, { left: 200, top: 20, width: 100, height: 50 }];

    const shapes = addTextBoxes(slide as never, rects, ["Uno", "Dos"], { fontSize: 20, bold: true, fontColor: "#112233" });

    expect(shapes).toHaveLength(2);
    const addCalls = ctx.calls.filter((c) => c.path === "addTextBox");
    expect(addCalls).toEqual([
      { id: shapes[0]!.id, path: "addTextBox", value: { text: "Uno", ...rects[0] } },
      { id: shapes[1]!.id, path: "addTextBox", value: { text: "Dos", ...rects[1] } },
    ]);

    for (const shape of shapes) {
      expect((shape as unknown as MockShape).formatWrites).toEqual([
        { path: "font.size", value: 20 },
        { path: "font.bold", value: true },
        { path: "font.color", value: "#112233" },
        { path: "textFrame.wordWrap", value: true },
        { path: "textFrame.autoSizeSetting", value: "AutoSizeNone" },
        { path: "textFrame.verticalAlignment", value: "Middle" },
      ]);
    }
  });

  it("clears the fill and hides the line when style.fill/style.line are null", () => {
    const { slide } = newSlide();

    const [shape] = addTextBoxes(slide as never, [rect], ["Hola"], { fill: null, line: null });

    const paths = (shape as unknown as MockShape).formatWrites.map((w) => w.path);
    expect(paths).toEqual(["fill.clear", "lineFormat.visible", "textFrame.wordWrap", "textFrame.autoSizeSetting", "textFrame.verticalAlignment"]);
  });

  it("sets a solid fill color, and forces the line visible before setting its color", () => {
    const { slide } = newSlide();

    const [shape] = addTextBoxes(slide as never, [rect], ["Hola"], { fill: "#FF0000", line: "#00FF00" });

    expect((shape as unknown as MockShape).formatWrites).toEqual([
      { path: "fill.setSolidColor", value: "#FF0000" },
      { path: "lineFormat.visible", value: true },
      { path: "lineFormat.color", value: "#00FF00" },
      { path: "textFrame.wordWrap", value: true },
      { path: "textFrame.autoSizeSetting", value: "AutoSizeNone" },
      { path: "textFrame.verticalAlignment", value: "Middle" },
    ]);
  });

  it("writes only the wordWrap/autoSizeSetting/verticalAlignment defaults without a style", () => {
    const { slide } = newSlide();

    const [shape] = addTextBoxes(slide as never, [rect], ["Hola"]);

    expect((shape as unknown as MockShape).formatWrites).toEqual([
      { path: "textFrame.wordWrap", value: true },
      { path: "textFrame.autoSizeSetting", value: "AutoSizeNone" },
      { path: "textFrame.verticalAlignment", value: "Middle" },
    ]);
  });

  it("falls back to an empty string when a rect has no matching text", () => {
    const { ctx, slide } = newSlide();

    addTextBoxes(slide as never, [rect], []);

    expect(ctx.calls[0]).toEqual({ id: expect.any(String), path: "addTextBox", value: { text: "", ...rect } });
  });
});

describe("addGeometric", () => {
  it("adds a geometric shape of the given type, recording the call", () => {
    const { ctx, slide } = newSlide();

    const shape = addGeometric(slide as never, "Chevron", rect);

    expect(ctx.calls[0]).toEqual({ id: shape.id, path: "addGeometricShape", value: { type: "Chevron", ...rect } });
    expect((shape as unknown as MockShape).formatWrites).toEqual([
      { path: "textFrame.wordWrap", value: true },
      { path: "textFrame.autoSizeSetting", value: "AutoSizeNone" },
      { path: "textFrame.verticalAlignment", value: "Middle" },
    ]);
  });

  it("writes text, font, fill/line and centers the paragraph when style.centered", () => {
    const { slide } = newSlide();

    const shape = addGeometric(slide as never, "Ellipse", rect, {
      text: "1",
      fontSize: 24,
      bold: true,
      fontColor: "#FFFFFF",
      fill: "#0000FF",
      line: null,
      centered: true,
    });

    expect((shape as unknown as MockShape).formatWrites).toEqual([
      { path: "fill.setSolidColor", value: "#0000FF" },
      { path: "lineFormat.visible", value: false },
      { path: "textRange.text", value: "1" },
      { path: "font.size", value: 24 },
      { path: "font.bold", value: true },
      { path: "font.color", value: "#FFFFFF" },
      { path: "paragraphFormat.horizontalAlignment", value: "Center" },
      { path: "textFrame.wordWrap", value: true },
      { path: "textFrame.autoSizeSetting", value: "AutoSizeNone" },
      { path: "textFrame.verticalAlignment", value: "Middle" },
    ]);
  });

  it("forces the line visible before setting its color, same as addTextBoxes", () => {
    const { slide } = newSlide();

    const shape = addGeometric(slide as never, "Rectangle", rect, { line: "#00FF00" });

    expect((shape as unknown as MockShape).formatWrites).toEqual([
      { path: "lineFormat.visible", value: true },
      { path: "lineFormat.color", value: "#00FF00" },
      { path: "textFrame.wordWrap", value: true },
      { path: "textFrame.autoSizeSetting", value: "AutoSizeNone" },
      { path: "textFrame.verticalAlignment", value: "Middle" },
    ]);
  });

  it("does not center the paragraph without style.centered", () => {
    const { slide } = newSlide();

    const shape = addGeometric(slide as never, "Rectangle", rect, { text: "A" });

    const paths = (shape as unknown as MockShape).formatWrites.map((w) => w.path);
    expect(paths).not.toContain("paragraphFormat.horizontalAlignment");
  });

  it("skips font.color when fontColor is an empty string", () => {
    const { slide } = newSlide();

    const shape = addGeometric(slide as never, "Rectangle", rect, { fontColor: "" });

    const paths = (shape as unknown as MockShape).formatWrites.map((w) => w.path);
    expect(paths).not.toContain("font.color");
  });
});

describe("applyTextDefaults (autoSizeSetting)", () => {
  it("sets autoSizeSetting to AutoSizeNone for both addTextBoxes and addGeometric with text, so columns do not shrink", () => {
    const { slide } = newSlide();

    const [textBox] = addTextBoxes(slide as never, [rect], ["Hola"]);
    const geometric = addGeometric(slide as never, "Rectangle", rect, { text: "A" });

    for (const shape of [textBox!, geometric]) {
      const write = (shape as unknown as MockShape).formatWrites.find((w) => w.path === "textFrame.autoSizeSetting");
      expect(write).toEqual({ path: "textFrame.autoSizeSetting", value: "AutoSizeNone" });
    }
  });
});

describe("addLine", () => {
  it("adds a line of the given kind, recording the call, without touching any style", () => {
    const { ctx, slide } = newSlide();

    const shape = addLine(slide as never, "Straight", rect);

    expect(ctx.calls).toEqual([{ id: shape.id, path: "addLine", value: { kind: "Straight", ...rect } }]);
    expect((shape as unknown as MockShape).formatWrites).toEqual([]);
  });

  it("supports the Elbow kind too", () => {
    const { ctx, slide } = newSlide();

    const shape = addLine(slide as never, "Elbow", rect);

    expect(ctx.calls[0]).toEqual({ id: shape.id, path: "addLine", value: { kind: "Elbow", ...rect } });
  });
});

describe("selectCreated", () => {
  it("selects the shapes created by addTextBoxes, by id (as the caller would after loading id and syncing)", () => {
    const { ctx, slide } = newSlide();
    const shapes = addTextBoxes(slide as never, [rect, rect], ["Uno", "Dos"]);

    selectCreated(slide as never, shapes);

    expect(ctx.selectedShapeIds).toEqual(shapes.map((s) => s.id));
  });

  it("selects an empty list without throwing", () => {
    const { ctx, slide } = newSlide();

    selectCreated(slide as never, []);

    expect(ctx.selectedShapeIds).toEqual([]);
  });
});
