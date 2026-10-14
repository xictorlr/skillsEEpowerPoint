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
    const t = fmt("t", { fill: { type: "NoFill", color: null, transparency: null } });
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
    const picture = fmt("pic", { font: null });
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
});
