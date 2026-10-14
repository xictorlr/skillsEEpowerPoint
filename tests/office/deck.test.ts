import { afterEach, describe, expect, it, vi } from "vitest";
import {
  addSlideAfter,
  addSlidesAfter,
  addTextBox,
  createPresentation,
  DECK_PROPS,
  deletePlaceholders,
  deleteShapes,
  exportSlides,
  findBlankLayout,
  loadFontNames,
  loadHasText,
  readDeckShapes,
  readFontName,
  readSlideShapesWithProxies,
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
  it("reads the shapes of every slide, with their slide id/index, proxies and textCapable, in a single sync", async () => {
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

    const { shapes, proxies, textCapable } = await readDeckShapes(ctx as never);

    expect(ctx.syncCount).toBe(1);
    expect(ctx.slideLoads).toEqual([DECK_PROPS]);
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
    // s1-a (GeometricShape) admite marco de texto; s2-a (Image) no.
    expect(textCapable).toEqual(new Set(["s1-a"]));
  });

  it("returns empty results for a deck with no slides", async () => {
    const ctx = new MockContext([], "", { slides: [] });
    const { shapes, proxies, textCapable } = await readDeckShapes(ctx as never);
    expect(shapes).toEqual([]);
    expect(proxies.size).toBe(0);
    expect(textCapable.size).toBe(0);
  });

  it("does a single sync when the deck has no placeholders", async () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 1, height: 1 });
    const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [box] }] });
    await readDeckShapes(ctx as never);
    expect(ctx.syncCount).toBe(1);
  });

  it("queues placeholderFormat.load(\"containedType\") and does a 2nd sync when the deck has placeholders; textCapable follows containedType", async () => {
    const imgPh = new MockShape({ id: "ph-img", type: "Placeholder", containedType: "Image", left: 0, top: 0, width: 1, height: 1 });
    const freePh = new MockShape({ id: "ph-free", type: "Placeholder", containedType: null, left: 0, top: 0, width: 1, height: 1 });
    const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [imgPh, freePh] }] });

    const { proxies, textCapable } = await readDeckShapes(ctx as never);

    expect(ctx.syncCount).toBe(2);
    expect(textCapable.has("ph-img")).toBe(false);
    expect(textCapable.has("ph-free")).toBe(true);

    // Un marcador de posición de imagen no debe tocar textFrame; uno de contenido libre sí.
    loadHasText(proxies, ["ph-img", "ph-free"], textCapable);
    expect(imgPh.loads).toEqual([]);
    expect(freePh.loads).toEqual([{ path: "textFrame", props: "hasText" }]);
    expect(ctx.invalidTouched).toBe(false);
  });
});

describe("readSlideShapesWithProxies", () => {
  it("reads only the shapes of the given slide, with proxies and textCapable, in a single sync", async () => {
    const onActive = new MockShape({ id: "a-1", name: "Caja", left: 0, top: 0, width: 10, height: 10 });
    const onOther = new MockShape({ id: "b-1", type: "Image", left: 0, top: 0, width: 5, height: 5 });
    const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [onActive] }, { id: "sl2", shapes: [onOther] }] });

    const { shapes, proxies, textCapable } = await readSlideShapesWithProxies(ctx as never, "sl1");

    expect(ctx.syncCount).toBe(1);
    expect(shapes).toEqual([
      { id: "a-1", name: "Caja", type: "GeometricShape", left: 0, top: 0, width: 10, height: 10, rotation: 0, zOrder: 0, isGroupChild: false },
    ]);
    expect(proxies.size).toBe(1);
    expect(proxies.get("a-1")).toBe(onActive);
    expect(proxies.has("b-1")).toBe(false); // la diapositiva "sl2" no se toca
    expect(textCapable).toEqual(new Set(["a-1"]));
  });

  it("queues placeholderFormat.load(\"containedType\") and does a 2nd sync only when THIS slide has placeholders", async () => {
    const freePh = new MockShape({ id: "ph-free", type: "Placeholder", containedType: null, left: 0, top: 0, width: 1, height: 1 });
    const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [freePh] }, { id: "sl2" }] });

    const { textCapable } = await readSlideShapesWithProxies(ctx as never, "sl1");

    expect(ctx.syncCount).toBe(2);
    expect(textCapable.has("ph-free")).toBe(true);
  });

  it("does a single sync when the slide has no placeholders", async () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 1, height: 1 });
    const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [box] }] });

    await readSlideShapesWithProxies(ctx as never, "sl1");

    expect(ctx.syncCount).toBe(1);
  });
});

describe("loadHasText", () => {
  it("queues textFrame.load(\"hasText\") only for ids in textCapable", () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 1, height: 1 });
    const img = new MockShape({ id: "b", type: "Image", left: 0, top: 0, width: 1, height: 1, font: null });
    const ctx = new MockContext([box, img]);

    loadHasText(new Map([["a", box], ["b", img]]) as never, ["a", "b"], new Set(["a"]));

    expect(box.loads).toEqual([{ path: "textFrame", props: "hasText" }]);
    expect(img.loads).toEqual([]);
    // No debe tocar `textFrame` de la imagen (eso marcaría invalidTouched en el mock).
    expect(ctx.invalidTouched).toBe(false);
  });

  it("ignores ids missing from the proxies map, even if listed in textCapable", () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 1, height: 1 });
    expect(() => loadHasText(new Map([["a", box]]) as never, ["missing"], new Set(["missing"]))).not.toThrow();
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
  it("queues font.load(\"name\") only for ids in textCapable, and readFontName reads it back", () => {
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

    loadFontNames(new Map([["a", box], ["b", img]]) as never, ["a", "b"], new Set(["a"]));

    expect(box.loads).toEqual([{ path: "font", props: "name" }]);
    expect(img.loads).toEqual([]);
    expect(readFontName(box as never)).toBe("Georgia");
  });
});

describe("addSlideAfter", () => {
  it("adds a slide at the end and moves it right after `index` (index + 1), in a single sync", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }, { id: "s2" }] });

    const slide = await addSlideAfter(ctx as never, 1);

    expect(ctx.syncCount).toBe(1);
    expect(ctx.slidesList).toHaveLength(3);
    expect(slide.id).toBe("new-slide-1");
    expect(ctx.calls).toEqual([
      { id: "presentation", path: "slides.add", value: undefined },
      { id: "new-slide-1", path: "moveTo", value: 2 },
    ]);
  });

  it("passes { slideMasterId, layoutId } to slides.add when given", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });

    await addSlideAfter(ctx as never, 0, { slideMasterId: "master-1", layoutId: "lay-blank" });

    expect(ctx.calls[0]).toEqual({
      id: "presentation",
      path: "slides.add",
      value: { slideMasterId: "master-1", layoutId: "lay-blank" },
    });
  });
});

describe("addSlidesAfter", () => {
  it("adds `count` slides and moves each right after `index`, in insertion order, with a single sync", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });

    const added = await addSlidesAfter(ctx as never, 0, 3);

    expect(ctx.syncCount).toBe(1);
    expect(ctx.slidesList).toHaveLength(4);
    expect(added.map((s) => (s as unknown as { id: string }).id)).toEqual(["new-slide-1", "new-slide-2", "new-slide-3"]);

    const addCalls = ctx.calls.filter((c) => c.path === "slides.add");
    expect(addCalls).toHaveLength(3);
    const moveToCalls = ctx.calls.filter((c) => c.path === "moveTo");
    expect(moveToCalls).toEqual([
      { id: "new-slide-1", path: "moveTo", value: 1 },
      { id: "new-slide-2", path: "moveTo", value: 2 },
      { id: "new-slide-3", path: "moveTo", value: 3 },
    ]);
    // Los 3 `slides.add` quedan encolados antes del único sync (el mock los ejecuta de forma
    // síncrona, pero el código bajo prueba no espera entre ellos).
    const addIndexes = addCalls.map((c) => ctx.calls.indexOf(c));
    const firstMoveToIndex = ctx.calls.findIndex((c) => c.path === "moveTo");
    expect(Math.max(...addIndexes)).toBeLessThan(firstMoveToIndex);
  });

  it("passes { slideMasterId, layoutId } to every slides.add call when given", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }] });

    await addSlidesAfter(ctx as never, 0, 2, { slideMasterId: "master-1", layoutId: "lay-blank" });

    const addCalls = ctx.calls.filter((c) => c.path === "slides.add");
    expect(addCalls).toHaveLength(2);
    for (const call of addCalls) {
      expect(call.value).toEqual({ slideMasterId: "master-1", layoutId: "lay-blank" });
    }
  });

  it("inserts sequentially after a non-zero index (moveTo index+1, index+2, …)", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }, { id: "s2" }, { id: "s3" }] });

    await addSlidesAfter(ctx as never, 1, 2);

    const moveToCalls = ctx.calls.filter((c) => c.path === "moveTo").map((c) => c.value);
    expect(moveToCalls).toEqual([2, 3]);
  });
});

describe("findBlankLayout", () => {
  it("returns the slide master id and the id of the first layout whose name matches blank/en blanco/vacío, in a single sync", async () => {
    const ctx = new MockContext([]);
    const layout = await findBlankLayout(ctx as never);
    expect(layout).toEqual({ slideMasterId: "master-1", layoutId: "lay-blank" });
    expect(ctx.syncCount).toBe(1);
  });

  it("returns undefined when no layout matches", async () => {
    const ctx = new MockContext([], "", { layouts: [{ id: "lay-title", name: "Title Slide" }, { id: "lay-content", name: "Title and Content" }] });
    const layout = await findBlankLayout(ctx as never);
    expect(layout).toBeUndefined();
  });

  it("returns undefined instead of throwing when reading layouts fails", async () => {
    const ctx = new MockContext([]);
    ctx.invalidTouched = true; // fuerza que el próximo sync() lance
    const layout = await findBlankLayout(ctx as never);
    expect(layout).toBeUndefined();
  });
});

describe("deletePlaceholders", () => {
  it("deletes every placeholder shape of the slide and returns the count, in a single sync", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }], newSlidePlaceholders: 2 });
    const slide = await addSlideAfter(ctx as never, 0);
    const syncsBefore = ctx.syncCount;

    const count = await deletePlaceholders(ctx as never, slide);

    expect(count).toBe(2);
    expect(ctx.syncCount).toBe(syncsBefore + 1);
    const placeholderDeletes = ctx.calls.filter((c) => c.path === "delete");
    expect(placeholderDeletes).toHaveLength(2);
  });

  it("returns 0 for a slide with no placeholders", async () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 1, height: 1 });
    const ctx = new MockContext([], "", { slides: [{ id: "s1", shapes: [box] }] });
    const slide = ctx.slidesList[0]!;

    const count = await deletePlaceholders(ctx as never, slide as never);

    expect(count).toBe(0);
    expect(box.formatWrites).toEqual([]);
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
