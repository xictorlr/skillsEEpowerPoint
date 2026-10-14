import { describe, expect, it } from "vitest";
import { addGeometric, addLine, addTextBoxes, selectCreated } from "../../src/office/insert";
import { MockContext, MockShape } from "./mockPowerPoint";

const rect = { left: 10, top: 20, width: 100, height: 50 };

describe("addTextBoxes", () => {
  it("adds one text box per rect/text, recording each call and the shared style", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;
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
        { path: "textFrame.verticalAlignment", value: "Middle" },
        { path: "textFrame.wordWrap", value: true },
      ]);
    }
  });

  it("clears the fill and hides the line when style.fill/style.line are null", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

    const [shape] = addTextBoxes(slide as never, [rect], ["Hola"], { fill: null, line: null });

    const paths = (shape as unknown as MockShape).formatWrites.map((w) => w.path);
    expect(paths).toEqual(["fill.clear", "lineFormat.visible", "textFrame.verticalAlignment", "textFrame.wordWrap"]);
  });

  it("sets a solid fill color and line color when given as hex strings", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

    const [shape] = addTextBoxes(slide as never, [rect], ["Hola"], { fill: "#FF0000", line: "#00FF00" });

    expect((shape as unknown as MockShape).formatWrites).toEqual([
      { path: "fill.setSolidColor", value: "#FF0000" },
      { path: "lineFormat.color", value: "#00FF00" },
      { path: "textFrame.verticalAlignment", value: "Middle" },
      { path: "textFrame.wordWrap", value: true },
    ]);
  });

  it("writes only the verticalAlignment/wordWrap defaults without a style", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

    const [shape] = addTextBoxes(slide as never, [rect], ["Hola"]);

    expect((shape as unknown as MockShape).formatWrites).toEqual([
      { path: "textFrame.verticalAlignment", value: "Middle" },
      { path: "textFrame.wordWrap", value: true },
    ]);
  });

  it("falls back to an empty string when a rect has no matching text", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

    addTextBoxes(slide as never, [rect], []);

    expect(ctx.calls[0]).toEqual({ id: expect.any(String), path: "addTextBox", value: { text: "", ...rect } });
  });
});

describe("addGeometric", () => {
  it("adds a geometric shape of the given type, recording the call", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

    const shape = addGeometric(slide as never, "Chevron", rect);

    expect(ctx.calls[0]).toEqual({ id: shape.id, path: "addGeometricShape", value: { type: "Chevron", ...rect } });
    expect((shape as unknown as MockShape).formatWrites).toEqual([
      { path: "textFrame.verticalAlignment", value: "Middle" },
      { path: "textFrame.wordWrap", value: true },
    ]);
  });

  it("writes text, font, fill/line and centers the paragraph when style.centered", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

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
      { path: "textFrame.verticalAlignment", value: "Middle" },
      { path: "textFrame.wordWrap", value: true },
    ]);
  });

  it("does not center the paragraph without style.centered", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

    const shape = addGeometric(slide as never, "Rectangle", rect, { text: "A" });

    const paths = (shape as unknown as MockShape).formatWrites.map((w) => w.path);
    expect(paths).not.toContain("paragraphFormat.horizontalAlignment");
  });
});

describe("addLine", () => {
  it("adds a line of the given kind, recording the call, without touching any style", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

    const shape = addLine(slide as never, "Straight", rect);

    expect(ctx.calls).toEqual([{ id: shape.id, path: "addLine", value: { kind: "Straight", ...rect } }]);
    expect((shape as unknown as MockShape).formatWrites).toEqual([]);
  });

  it("supports the Elbow kind too", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

    const shape = addLine(slide as never, "Elbow", rect);

    expect(ctx.calls[0]).toEqual({ id: shape.id, path: "addLine", value: { kind: "Elbow", ...rect } });
  });
});

describe("selectCreated", () => {
  it("selects the given shapes by id", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;
    const a = new MockShape({ id: "a", left: 0, top: 0, width: 1, height: 1 });
    const b = new MockShape({ id: "b", left: 0, top: 0, width: 1, height: 1 });

    selectCreated(slide as never, [a, b] as never);

    expect(ctx.selectedShapeIds).toEqual(["a", "b"]);
  });

  it("selects an empty list without throwing", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

    selectCreated(slide as never, []);

    expect(ctx.selectedShapeIds).toEqual([]);
  });
});
