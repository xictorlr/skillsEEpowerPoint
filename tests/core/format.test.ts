import { describe, expect, it } from "vitest";
import {
  buildPastePatches,
  clearPatches,
  colorPatches,
  transparencyPatches,
  type ShapeFormat,
} from "../../src/core/format";

function fmt(id: string, over: Partial<ShapeFormat> = {}): ShapeFormat {
  return {
    id,
    fill: { type: "Solid", color: "#FF0000", transparency: 0 },
    line: { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 },
    font: { name: "Calibri", size: 12, color: "#111111", bold: false, italic: false, underline: "None" },
    size: { width: 100, height: 50 },
    ...over,
  };
}

const ALL = { fill: true, line: true, font: true, size: true };

describe("buildPastePatches", () => {
  it("copies every part from the source to each target, skipping the source itself", () => {
    const src = fmt("src");
    const t = fmt("t", {
      fill: { type: "NoFill", color: null, transparency: null },
      size: { width: 10, height: 10 },
    });
    expect(buildPastePatches(src, [src, t], ALL)).toEqual([
      {
        id: "t",
        fill: { type: "Solid", color: "#FF0000", transparency: 0 },
        line: { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 },
        font: { name: "Calibri", size: 12, color: "#111111", bold: false, italic: false, underline: "None" },
        width: 100,
        height: 50,
      },
    ]);
  });

  it("only includes the requested parts", () => {
    const src = fmt("src");
    const t = fmt("t");
    expect(buildPastePatches(src, [t], { fill: true, line: false, font: false, size: false })).toEqual([
      { id: "t", fill: { type: "Solid", color: "#FF0000", transparency: 0 } },
    ]);
  });

  it("omits null source values and never pastes font onto shapes without text", () => {
    const src = fmt("src", { fill: { type: "Solid", color: null, transparency: 0.5 } });
    const picture = fmt("pic", { font: null, size: { width: 10, height: 10 } });
    expect(buildPastePatches(src, [picture], ALL)).toEqual([
      {
        id: "pic",
        fill: { type: "Solid", transparency: 0.5 },
        line: { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 },
        width: 100,
        height: 50,
      },
    ]);
  });

  it("returns [] when there is no target", () => {
    expect(buildPastePatches(fmt("src"), [fmt("src")], ALL)).toEqual([]);
  });

  it("returns [] for an empty target list", () => {
    expect(buildPastePatches(fmt("src"), [], ALL)).toEqual([]);
  });

  it("pastes a NoFill source as only { type: 'NoFill' }, ignoring any stale color", () => {
    const src = fmt("src", { fill: { type: "NoFill", color: "#ABCDEF", transparency: null } });
    const t = fmt("t");
    expect(buildPastePatches(src, [t], { fill: true, line: false, font: false, size: false })).toEqual([
      { id: "t", fill: { type: "NoFill" } },
    ]);
  });

  it("skips fill entirely when the source fill is 'Other' (gradient/pattern/picture)", () => {
    const src = fmt("src", { fill: { type: "Other", color: null, transparency: 0.2 } });
    const t = fmt("t", { size: { width: 1, height: 1 } });
    expect(buildPastePatches(src, [t], { fill: true, line: false, font: false, size: true })).toEqual([
      { id: "t", width: 100, height: 50 },
    ]);
  });

  it("pastes an invisible-line source as only { visible: false }, ignoring stale line fields", () => {
    const src = fmt("src", {
      line: { visible: false, color: "#ABCDEF", weight: 3, dashStyle: "Dash", transparency: 0.4 },
    });
    const t = fmt("t");
    expect(buildPastePatches(src, [t], { fill: false, line: true, font: false, size: false })).toEqual([
      { id: "t", line: { visible: false } },
    ]);
  });

  it("omits width/height when the target already matches the source size", () => {
    const src = fmt("src");
    const t = fmt("t", { size: { width: 100.005, height: 50 } });
    expect(buildPastePatches(src, [t], { fill: false, line: false, font: false, size: true })).toEqual([]);
  });
});

describe("colorPatches", () => {
  it("sets fill, line or font color on every shape", () => {
    const a = fmt("a");
    expect(colorPatches([a], "fill", "#00FF00")).toEqual([{ id: "a", fill: { type: "Solid", color: "#00FF00" } }]);
    expect(colorPatches([a], "line", "#00FF00")).toEqual([{ id: "a", line: { visible: true, color: "#00FF00" } }]);
    expect(colorPatches([a], "font", "#00FF00")).toEqual([{ id: "a", font: { color: "#00FF00" } }]);
  });

  it("skips font color for shapes without text", () => {
    expect(colorPatches([fmt("pic", { font: null })], "font", "#00FF00")).toEqual([]);
  });

  it("returns [] for an invalid color", () => {
    expect(colorPatches([fmt("a")], "fill", "not-a-color")).toEqual([]);
  });

  it("normalizes a lowercase color in the output", () => {
    expect(colorPatches([fmt("a")], "fill", "#00ff00")).toEqual([{ id: "a", fill: { type: "Solid", color: "#00FF00" } }]);
  });

  it("returns [] for an empty target list", () => {
    expect(colorPatches([], "fill", "#00FF00")).toEqual([]);
  });
});

describe("transparencyPatches", () => {
  it("steps transparency and clamps to [0, 1]", () => {
    const a = fmt("a", { fill: { type: "Solid", color: "#FF0000", transparency: 0.95 } });
    expect(transparencyPatches([a], "fill", 0.1)).toEqual([{ id: "a", fill: { transparency: 1 } }]);
    expect(transparencyPatches([a], "fill", -0.1)).toEqual([{ id: "a", fill: { transparency: 0.85 } }]);
  });

  it("treats null transparency as 0 and skips invisible lines", () => {
    const a = fmt("a", {
      fill: { type: "Solid", color: "#FF0000", transparency: null },
      line: { visible: false, color: null, weight: null, dashStyle: null, transparency: null },
    });
    expect(transparencyPatches([a], "fill", 0.2)).toEqual([{ id: "a", fill: { transparency: 0.2 } }]);
    expect(transparencyPatches([a], "line", 0.2)).toEqual([]);
  });

  it("omits shapes whose value would not change", () => {
    const a = fmt("a", { fill: { type: "Solid", color: "#FF0000", transparency: 1 } });
    expect(transparencyPatches([a], "fill", 0.1)).toEqual([]);
  });

  it("skips shapes without a solid fill (NoFill/Other) when target is fill", () => {
    const noFill = fmt("a", { fill: { type: "NoFill", color: null, transparency: null } });
    expect(transparencyPatches([noFill], "fill", 0.1)).toEqual([]);
    const other = fmt("b", { fill: { type: "Other", color: null, transparency: 0.2 } });
    expect(transparencyPatches([other], "fill", 0.1)).toEqual([]);
  });

  it("returns [] for a negative delta that would not move transparency 0 below 0", () => {
    const a = fmt("a", { fill: { type: "Solid", color: "#FF0000", transparency: 0 } });
    expect(transparencyPatches([a], "fill", -0.1)).toEqual([]);
  });

  it("returns [] when delta is 0, even with a fractional current transparency", () => {
    const a = fmt("a", { fill: { type: "Solid", color: "#FF0000", transparency: 0.298 } });
    expect(transparencyPatches([a], "fill", 0)).toEqual([]);
  });

  it("returns [] for an empty target list", () => {
    expect(transparencyPatches([], "fill", 0.1)).toEqual([]);
  });
});

describe("clearPatches", () => {
  it("removes fill or hides line, skipping shapes already cleared", () => {
    const a = fmt("a");
    const b = fmt("b", {
      fill: { type: "NoFill", color: null, transparency: null },
      line: { visible: false, color: null, weight: null, dashStyle: null, transparency: null },
    });
    expect(clearPatches([a, b], "fill")).toEqual([{ id: "a", fill: { type: "NoFill" } }]);
    expect(clearPatches([a, b], "line")).toEqual([{ id: "a", line: { visible: false } }]);
  });

  it("hides the line when visible is null (mixed/not applicable)", () => {
    const a = fmt("a", {
      line: { visible: null, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 },
    });
    expect(clearPatches([a], "line")).toEqual([{ id: "a", line: { visible: false } }]);
  });

  it("returns [] for an empty target list", () => {
    expect(clearPatches([], "fill")).toEqual([]);
    expect(clearPatches([], "line")).toEqual([]);
  });
});
