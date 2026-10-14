import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { executeCommand, resetSlideSizeCache } from "../../src/commands/execute";
import { findCommand } from "../../src/commands/registry";
import type { CommandOptions } from "../../src/commands/types";
import { MockContext, MockShape } from "../office/mockPowerPoint";

const DEFAULT_OPTIONS: CommandOptions = {
  matrix: { columns: 2, gapX: 8, gapY: 8 },
  colorTarget: "fill",
  pasteParts: { fill: true, line: true, font: true, size: true },
  fontSizeStep: 1,
  nameFilter: "",
  renamePrefix: "Forma",
  shapeKind: "all",
};

function stubPowerPoint(ctx: MockContext): void {
  vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => Promise<unknown>) => fn(ctx) });
}

describe("expert commands", () => {
  beforeEach(() => {
    resetSlideSizeCache();
    vi.spyOn(console, "error").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("select-kind 'image' selects only the images on the active slide", async () => {
    const img = new MockShape({ id: "img1", type: "Image", left: 0, top: 0, width: 10, height: 10 });
    const txt = new MockShape({ id: "txt1", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [img, txt] }] });
    stubPowerPoint(ctx);

    const cmd = findCommand("select-kind")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, shapeKind: "image" });

    expect(result).toEqual({ ok: true, changed: 1 });
    expect(ctx.selectedShapeIds).toEqual(["img1"]);
  });

  it("select-kind with no matches reports a message and selects nothing", async () => {
    const txt = new MockShape({ id: "txt1", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [txt] }] });
    stubPowerPoint(ctx);

    const cmd = findCommand("select-kind")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, shapeKind: "image" });

    expect(result).toEqual({ ok: true, changed: 0, message: "No hay formas de ese tipo." });
    expect(ctx.selectedShapeIds).toEqual([]);
  });

  it("select-by-name with an empty filter reports a message without touching the selection", async () => {
    const ctx = new MockContext([], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("select-by-name")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, nameFilter: "  " });

    expect(result).toEqual({ ok: true, changed: 0, message: "Escribe parte del nombre." });
    expect(ctx.selectedShapeIds).toBeUndefined();
  });

  it("select-by-name selects the shapes matching the filter", async () => {
    const a = new MockShape({ id: "a", name: "Logo Acme", left: 0, top: 0, width: 10, height: 10 });
    const b = new MockShape({ id: "b", name: "Título", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [a, b] }] });
    stubPowerPoint(ctx);

    const cmd = findCommand("select-by-name")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, nameFilter: "logo" });

    expect(result).toEqual({ ok: true, changed: 1 });
    expect(ctx.selectedShapeIds).toEqual(["a"]);
  });

  it("rename-numbered writes names in reading order using the given prefix", async () => {
    const a = new MockShape({ id: "a", name: "X", left: 0, top: 0, width: 10, height: 10 });
    const b = new MockShape({ id: "b", name: "Y", left: 20, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([a, b], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("rename-numbered")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, renamePrefix: "Caja" });

    expect(result).toEqual({ ok: true, changed: 2 });
    expect(a.name).toBe("Caja 1");
    expect(b.name).toBe("Caja 2");
  });

  it("z-front applies setZOrder in ascending zOrder order", async () => {
    const callOrder: string[] = [];
    function tracked(init: ConstructorParameters<typeof MockShape>[0]): MockShape {
      const s = new MockShape(init);
      const original = s.setZOrder.bind(s);
      s.setZOrder = (op: string) => {
        callOrder.push(s.id);
        original(op);
      };
      return s;
    }
    const a = tracked({ id: "a", zOrderPosition: 5, left: 0, top: 0, width: 10, height: 10 });
    const b = tracked({ id: "b", zOrderPosition: 1, left: 0, top: 0, width: 10, height: 10 });
    const c = tracked({ id: "c", zOrderPosition: 3, left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([a, b, c], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("z-front")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 3 });
    expect(callOrder).toEqual(["b", "c", "a"]);
    expect(a.formatWrites).toContainEqual({ path: "setZOrder", value: "BringToFront" });
  });

  it("group with a single shape fails the minShapes precondition without grouping", async () => {
    const s = new MockShape({ id: "s1", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([s], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("group")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: false, message: "Selecciona al menos 2 formas." });
    expect(ctx.grouped).toBe(false);
  });

  it("group with two shapes groups them", async () => {
    const a = new MockShape({ id: "a", left: 0, top: 0, width: 10, height: 10 });
    const b = new MockShape({ id: "b", left: 20, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([a, b], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("group")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 1, message: "Agrupado." });
    expect(ctx.grouped).toBe(true);
  });

  it("ungroup with no groups in the selection reports a message", async () => {
    const s = new MockShape({ id: "s1", type: "GeometricShape", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([s], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("ungroup")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 0, message: "No hay grupos en la selección." });
  });

  it("ungroup dissolves every group in the selection", async () => {
    const g = new MockShape({ id: "g1", type: "Group", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([g], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("ungroup")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 1 });
    expect(g.formatWrites).toContainEqual({ path: "ungroup", value: null });
  });

  it("inspector-apply moves the selection to the requested rectangle", async () => {
    const s = new MockShape({ id: "s1", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([s], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("inspector-apply")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, inspector: { left: 100, top: 50 } });

    expect(result).toEqual({ ok: true, changed: 1 });
    expect(s.writes).toEqual([
      { prop: "left", value: 100 },
      { prop: "top", value: 50 },
    ]);
  });
});
