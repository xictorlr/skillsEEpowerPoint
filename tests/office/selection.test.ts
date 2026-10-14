import { describe, expect, it, vi } from "vitest";
import { readSelection } from "../../src/office/selection";
import { MockContext, MockShape } from "./mockPowerPoint";

const PNG_1x1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

const SHAPE_PROP_NAMES = [
  "id",
  "name",
  "type",
  "left",
  "top",
  "width",
  "height",
  "rotation",
  "zOrderPosition",
  "level",
];

describe("readSelection", () => {
  it("maps proxies to ShapeModel and keeps a proxy map by id", async () => {
    const ctx = new MockContext(
      [new MockShape({ id: "a", left: 1, top: 2, width: 3, height: 4, rotation: 10, zOrderPosition: 2, level: 1 })],
      PNG_1x1,
    );
    const { selection, proxies } = await readSelection(ctx as never, { slideSizeCache: new Map() });
    expect(selection.shapes).toEqual([
      { id: "a", name: "a", type: "GeometricShape", left: 1, top: 2, width: 3, height: 4, rotation: 10, zOrder: 2, isGroupChild: true },
    ]);
    expect(proxies.get("a")).toBeDefined();
    expect(selection.slide).toEqual({ width: 540, height: 540 }); // PNG 1x1 → aspecto 1:1
  });

  it("falls back to 960x540 without caching it, after trying to render the slide", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => undefined);
    const cache = new Map();
    const ctx = new MockContext([], "not-base64-png");
    const { selection } = await readSelection(ctx as never, { slideSizeCache: cache });
    expect(selection.shapes).toEqual([]);
    expect(selection.slide).toEqual({ width: 960, height: 540 });
    expect(ctx.imageOptions).toEqual({ height: 540 }); // intentó renderizar antes de caer al fallback
    expect(cache.size).toBe(0); // el fallback no se cachea: puede ser un fallo transitorio
    warn.mockRestore();
  });

  it("loads every shape property it reads", async () => {
    const ctx = new MockContext(
      [new MockShape({ id: "a", left: 1, top: 2, width: 3, height: 4 })],
      PNG_1x1,
    );
    await readSelection(ctx as never, { slideSizeCache: new Map() });
    for (const prop of SHAPE_PROP_NAMES) {
      expect(ctx.loadedProps).toContain(`items/${prop}`);
    }
  });

  it("renders the slide at 540px high", async () => {
    const ctx = new MockContext([], PNG_1x1);
    await readSelection(ctx as never, { slideSizeCache: new Map() });
    expect(ctx.imageOptions).toEqual({ height: 540 });
  });

  it("syncs twice the first time (shapes + slide render) and once when the slide size is already cached", async () => {
    const cache = new Map();

    const ctx1 = new MockContext([], PNG_1x1);
    await readSelection(ctx1 as never, { slideSizeCache: cache });
    expect(ctx1.syncCount).toBe(2);

    const ctx2 = new MockContext([], PNG_1x1);
    await readSelection(ctx2 as never, { slideSizeCache: cache });
    expect(ctx2.syncCount).toBe(1);
  });

  it("uses the slideSizeOverride verbatim and never renders the slide", async () => {
    const ctx = new MockContext([], PNG_1x1);
    const { selection } = await readSelection(ctx as never, {
      slideSizeCache: new Map(),
      slideSizeOverride: { width: 720, height: 405 },
    });
    expect(selection.slide).toEqual({ width: 720, height: 405 });
    expect(ctx.imageOptions).toBeUndefined();
  });

  it("loads format props and fonts when withFormat is set (two read syncs) and returns formats", async () => {
    const box = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 20,
      font: { name: "Arial", size: 14, color: "#112233", bold: true, italic: null, underline: "None" } });
    const ctx = new MockContext([box], "bad-png");
    const { formats } = await readSelection(ctx as never, { slideSizeCache: new Map(), slideSizeOverride: { width: 960, height: 540 }, withFormat: true });
    expect(ctx.loadedProps).toContain("items/fill/type");
    expect(ctx.loadedProps).toContain("items/lineFormat/weight");
    expect(box.fontLoaded).toBe(true);
    expect(ctx.syncCount).toBe(2);
    expect(formats[0]?.font?.name).toBe("Arial");
  });

  it("does not load format props by default", async () => {
    const ctx = new MockContext([new MockShape({ id: "a", left: 0, top: 0, width: 1, height: 1 })], "bad-png");
    const { formats } = await readSelection(ctx as never, { slideSizeCache: new Map(), slideSizeOverride: { width: 960, height: 540 } });
    expect(ctx.loadedProps).not.toContain("items/fill/type");
    expect(formats).toEqual([]);
  });
});
