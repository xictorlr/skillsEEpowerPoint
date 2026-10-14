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

  it("falls back to 960x540 when the slide image cannot be read", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => undefined);
    const ctx = new MockContext([], "not-base64-png");
    const { selection } = await readSelection(ctx as never, { slideSizeCache: new Map() });
    expect(selection.shapes).toEqual([]);
    expect(selection.slide).toEqual({ width: 960, height: 540 });
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
});
