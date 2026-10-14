import { afterEach, describe, expect, it, vi } from "vitest";
import {
  addSlideAfter,
  addTextBox,
  createPresentation,
  deleteShapes,
  exportSlides,
  loadFontNames,
  loadHasText,
  readDeckShapes,
  readFontName,
} from "../../src/office/deck";
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

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("readDeckShapes", () => {
  it("reads the shapes of every slide, with their slide id/index and proxies, in a single sync", async () => {
    const s1a = new MockShape({ id: "s1-a", name: "Caja 1", left: 0, top: 0, width: 10, height: 10 });
    const s2a = new MockShape({
      id: "s2-a",
      name: "Caja 2",
      type: "Image",
      left: 5,
      top: 5,
      width: 20,
      height: 20,
      rotation: 90,
      zOrderPosition: 1,
      level: 1,
    });
    const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [s1a] }, { id: "sl2", shapes: [s2a] }] });

    const { shapes, proxies } = await readDeckShapes(ctx as never);

    expect(ctx.syncCount).toBe(1);
    expect(shapes).toEqual([
      {
        slideId: "sl1",
        slideIndex: 0,
        shape: { id: "s1-a", name: "Caja 1", type: "GeometricShape", left: 0, top: 0, width: 10, height: 10, rotation: 0, zOrder: 0, isGroupChild: false },
      },
      {
        slideId: "sl2",
        slideIndex: 1,
        shape: { id: "s2-a", name: "Caja 2", type: "Image", left: 5, top: 5, width: 20, height: 20, rotation: 90, zOrder: 1, isGroupChild: true },
      },
    ]);
    expect(proxies.get("s1-a")).toBe(s1a);
    expect(proxies.get("s2-a")).toBe(s2a);
  });

  it("returns empty results for a deck with no slides", async () => {
    const ctx = new MockContext([], "", { slides: [] });
    const { shapes, proxies } = await readDeckShapes(ctx as never);
    expect(shapes).toEqual([]);
    expect(proxies.size).toBe(0);
  });
});

describe("loadHasText", () => {
  it("queues textFrame.load(\"hasText\") only on shapes with a text frame by type", () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 1, height: 1 });
    const img = new MockShape({ id: "b", type: "Image", left: 0, top: 0, width: 1, height: 1, font: null });
    const ctx = new MockContext([box, img]);

    loadHasText(new Map([["a", box], ["b", img]]) as never, ["a", "b"]);

    expect(box.loads).toEqual([{ path: "textFrame", props: "hasText" }]);
    expect(img.loads).toEqual([]);
    // No debe tocar `textFrame` de la imagen (eso marcaría invalidTouched en el mock).
    expect(ctx.invalidTouched).toBe(false);
  });

  it("ignores ids missing from the proxies map", () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 1, height: 1 });
    expect(() => loadHasText(new Map([["a", box]]) as never, ["missing"])).not.toThrow();
    expect(box.loads).toEqual([]);
  });
});

describe("deleteShapes", () => {
  it("calls shape.delete() on each id, in order, and records it on ctx.calls", () => {
    const a = new MockShape({ id: "a", left: 0, top: 0, width: 1, height: 1 });
    const b = new MockShape({ id: "b", left: 0, top: 0, width: 1, height: 1 });
    const ctx = new MockContext([a, b]);

    deleteShapes(new Map([["a", a], ["b", b]]) as never, ["a", "b"]);

    expect(a.formatWrites).toEqual([{ path: "delete", value: null }]);
    expect(b.formatWrites).toEqual([{ path: "delete", value: null }]);
    expect(ctx.calls).toEqual([
      { id: "a", path: "delete", value: null },
      { id: "b", path: "delete", value: null },
    ]);
  });

  it("ignores ids missing from the proxies map", () => {
    expect(() => deleteShapes(new Map() as never, ["missing"])).not.toThrow();
  });
});

describe("loadFontNames / readFontName", () => {
  it("queues font.load(\"name\") only on shapes with a text frame, and readFontName reads it back", () => {
    const box = new MockShape({
      id: "a",
      type: "TextBox",
      left: 0,
      top: 0,
      width: 1,
      height: 1,
      font: { name: "Georgia", size: 18, color: "#000000", bold: false, italic: false, underline: "None" },
    });
    const img = new MockShape({ id: "b", type: "Image", left: 0, top: 0, width: 1, height: 1, font: null });

    loadFontNames(new Map([["a", box], ["b", img]]) as never, ["a", "b"]);

    expect(box.loads).toEqual([{ path: "font", props: "name" }]);
    expect(img.loads).toEqual([]);
    expect(readFontName(box as never)).toBe("Georgia");
  });
});

describe("addSlideAfter", () => {
  it("adds a slide at the end, two syncs, then moves it right after `index` (index + 1)", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }, { id: "s2" }] });

    const slide = await addSlideAfter(ctx as never, 1);

    expect(ctx.syncCount).toBe(2);
    expect(ctx.slidesList).toHaveLength(3);
    expect(slide.id).toBe("new-slide-1");
    expect(ctx.calls).toEqual([{ id: "new-slide-1", path: "moveTo", value: 2 }]);
  });
});

describe("addTextBox", () => {
  const rect = { left: 10, top: 20, width: 100, height: 50 };

  it("adds the text box, records the call, and writes the given style", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

    const shape = addTextBox(slide as never, "Hola", rect, { size: 24, bold: true, color: "#FF0000" });

    expect(ctx.calls[0]).toEqual({ id: shape.id, path: "addTextBox", value: { text: "Hola", ...rect } });
    expect((shape as unknown as MockShape).formatWrites).toEqual([
      { path: "font.size", value: 24 },
      { path: "font.bold", value: true },
      { path: "font.color", value: "#FF0000" },
      { path: "textFrame.wordWrap", value: true },
      { path: "textFrame.autoSizeSetting", value: "AutoSizeNone" },
      { path: "textFrame.verticalAlignment", value: "Middle" },
    ]);
  });

  it("does not write font.color when style.color is null", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });
    const slide = ctx.slidesList[0]!;

    const shape = addTextBox(slide as never, "Sin color", rect, { size: 12, bold: false, color: null });

    const paths = (shape as unknown as MockShape).formatWrites.map((w) => w.path);
    expect(paths).not.toContain("font.color");
    expect(paths).toEqual(["font.size", "font.bold", "textFrame.wordWrap", "textFrame.autoSizeSetting", "textFrame.verticalAlignment"]);
  });
});

describe("exportSlides", () => {
  it("returns null without PowerPointApi 1.10 support, without syncing", async () => {
    const ctx = new MockContext([]);
    const result = await exportSlides(ctx as never, ["s1", "s2"]);
    expect(result).toBeNull();
    expect(ctx.syncCount).toBe(0);
  });

  it("returns the base64 string with 1.10 support, with a single sync", async () => {
    stubOffice1_10();
    const ctx = new MockContext([]);
    const result = await exportSlides(ctx as never, ["s1", "s2"]);
    expect(result).toBe("BASE64:s1,s2");
    expect(ctx.syncCount).toBe(1);
  });
});

describe("createPresentation", () => {
  it("delegates to the global PowerPoint.createPresentation with the given base64", async () => {
    const createPresentationMock = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("PowerPoint", { createPresentation: createPresentationMock });

    await createPresentation("base64==");

    expect(createPresentationMock).toHaveBeenCalledWith("base64==");
  });

  it("works without a base64 argument", async () => {
    const createPresentationMock = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("PowerPoint", { createPresentation: createPresentationMock });

    await createPresentation();

    expect(createPresentationMock).toHaveBeenCalledWith(undefined);
  });
});
