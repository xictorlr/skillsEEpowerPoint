import { describe, expect, it } from "vitest";
import {
  applyZOrder,
  groupSelected,
  readSlideNav,
  readSlideShapes,
  renameShapes,
  selectShapes,
  selectSlideByIndex,
  ungroupShapes,
  type ShapeOpsProxy,
} from "../../src/office/navigation";
import { MockContext, MockShape } from "./mockPowerPoint";

describe("readSlideNav", () => {
  it("reports the active slide (2nd of 3) with its index, count and ids", async () => {
    const ctx = new MockContext([], "", {
      slides: [{ id: "s1" }, { id: "s2" }, { id: "s3" }],
      activeSlideIndex: 1,
    });
    const nav = await readSlideNav(ctx as never);
    expect(nav).toEqual({ slideId: "s2", index: 1, count: 3, ids: ["s1", "s2", "s3"] });
    expect(ctx.syncCount).toBe(1);
  });

  it("throws a plain Error when the deck has no slides", async () => {
    const ctx = new MockContext([], "", { slides: [] });
    await expect(readSlideNav(ctx as never)).rejects.toThrow("No hay diapositivas.");
  });
});

describe("selectSlideByIndex", () => {
  it("selects the slide at index via setSelectedSlides([id])", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }, { id: "s2" }, { id: "s3" }] });
    const nav = { slideId: "s1", index: 0, count: 3, ids: ["s1", "s2", "s3"] };
    const ok = selectSlideByIndex(ctx as never, nav, 2);
    expect(ok).toBe(true);
    expect(ctx.selectedSlideIds).toEqual(["s3"]);
  });

  it("returns false and selects nothing when index is out of range", () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }, { id: "s2" }] });
    const nav = { slideId: "s1", index: 0, count: 2, ids: ["s1", "s2"] };
    expect(selectSlideByIndex(ctx as never, nav, -1)).toBe(false);
    expect(selectSlideByIndex(ctx as never, nav, 2)).toBe(false);
    expect(ctx.selectedSlideIds).toBeUndefined();
  });
});

describe("readSlideShapes", () => {
  it("maps the shapes of the given slide to ShapeModel[]", async () => {
    const shape = new MockShape({ id: "a", name: "Box 1", left: 1, top: 2, width: 3, height: 4, rotation: 10, zOrderPosition: 2, level: 1 });
    const ctx = new MockContext([], "", { slides: [{ id: "s1", shapes: [shape] }] });
    const shapes = await readSlideShapes(ctx as never, "s1");
    expect(shapes).toEqual([
      { id: "a", name: "Box 1", type: "GeometricShape", left: 1, top: 2, width: 3, height: 4, rotation: 10, zOrder: 2, isGroupChild: true },
    ]);
    expect(ctx.syncCount).toBe(1);
  });
});

describe("selectShapes / groupSelected / ungroupShapes / renameShapes / applyZOrder", () => {
  it("register the calls, in order, on the mock", () => {
    const a = new MockShape({ id: "a", left: 0, top: 0, width: 1, height: 1 });
    const b = new MockShape({ id: "b", left: 0, top: 0, width: 1, height: 1 });
    const group = new MockShape({ id: "g", type: "Group", left: 0, top: 0, width: 1, height: 1 });
    const ctx = new MockContext([a, b, group], "", { slides: [{ id: "s1", shapes: [a, b, group] }] });

    selectShapes(ctx as never, "s1", ["a", "b"]);
    expect(ctx.selectedShapeIds).toEqual(["a", "b"]);

    groupSelected(ctx as never);
    expect(ctx.grouped).toBe(true);

    const proxies = new Map<string, ShapeOpsProxy>([
      ["a", a],
      ["b", b],
      ["g", group],
    ]);

    ungroupShapes(proxies, ["g"]);
    renameShapes(proxies, [{ id: "a", name: "Uno" }, { id: "b", name: "Dos" }]);
    applyZOrder(proxies, ["a", "b"], "BringToFront");

    expect(group.formatWrites).toEqual([{ path: "ungroup", value: null }]);
    expect(a.formatWrites).toEqual([
      { path: "name", value: "Uno" },
      { path: "setZOrder", value: "BringToFront" },
    ]);
    expect(b.formatWrites).toEqual([
      { path: "name", value: "Dos" },
      { path: "setZOrder", value: "BringToFront" },
    ]);
  });
});
