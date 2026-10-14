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
    expect(ctx.syncCount).toBe(4); // selección + readSlideNav + readSlideShapes + escritura
  });

  it("select-kind with no matches reports a message and leaves the current selection untouched", async () => {
    const txt = new MockShape({ id: "txt1", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [txt] }] });
    stubPowerPoint(ctx);

    const cmd = findCommand("select-kind")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, shapeKind: "image" });

    expect(result).toEqual({ ok: true, changed: 0, message: "No hay formas de ese tipo." });
    expect(ctx.selectedShapeIds).toBeUndefined();
    expect(ctx.syncCount).toBe(3); // selección + readSlideNav + readSlideShapes, sin sync de escritura
  });

  it("select-by-name with an empty filter reports a message without touching the selection", async () => {
    const ctx = new MockContext([], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("select-by-name")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, nameFilter: "  " });

    expect(result).toEqual({ ok: true, changed: 0, message: "Escribe parte del nombre." });
    expect(ctx.selectedShapeIds).toBeUndefined();
    expect(ctx.syncCount).toBe(1); // solo la lectura de selección; ni siquiera lee la diapositiva
  });

  it("select-by-name with no matches reports its own message and leaves the current selection untouched", async () => {
    const a = new MockShape({ id: "a", name: "Título", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [a] }] });
    stubPowerPoint(ctx);

    const cmd = findCommand("select-by-name")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, nameFilter: "logo" });

    expect(result).toEqual({ ok: true, changed: 0, message: "Ninguna forma coincide con ese nombre." });
    expect(ctx.selectedShapeIds).toBeUndefined();
    expect(ctx.syncCount).toBe(3); // selección + readSlideNav + readSlideShapes, sin sync de escritura
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

  it("rename-numbered is a no-op (and skips the write sync) when every shape is already numbered", async () => {
    const a = new MockShape({ id: "a", name: "Caja 1", left: 0, top: 0, width: 10, height: 10 });
    const b = new MockShape({ id: "b", name: "Caja 2", left: 20, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([a, b], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("rename-numbered")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, renamePrefix: "Caja" });

    expect(result).toEqual({ ok: true, changed: 0 });
    expect(ctx.syncCount).toBe(1); // solo la lectura de selección; sin sync de escritura
  });

  it("z-front applies setZOrder in ascending zOrder order", async () => {
    const a = new MockShape({ id: "a", zOrderPosition: 5, left: 0, top: 0, width: 10, height: 10 });
    const b = new MockShape({ id: "b", zOrderPosition: 1, left: 0, top: 0, width: 10, height: 10 });
    const c = new MockShape({ id: "c", zOrderPosition: 3, left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([a, b, c], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("z-front")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 3 });
    expect(ctx.calls.map((c) => c.id)).toEqual(["b", "c", "a"]);
  });

  it("z-forward applies setZOrder in descending zOrder order", async () => {
    const a = new MockShape({ id: "a", zOrderPosition: 5, left: 0, top: 0, width: 10, height: 10 });
    const b = new MockShape({ id: "b", zOrderPosition: 1, left: 0, top: 0, width: 10, height: 10 });
    const c = new MockShape({ id: "c", zOrderPosition: 3, left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([a, b, c], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("z-forward")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 3 });
    expect(ctx.calls.map((c) => c.id)).toEqual(["a", "c", "b"]);
  });

  it("z-back applies setZOrder in descending zOrder order", async () => {
    const a = new MockShape({ id: "a", zOrderPosition: 5, left: 0, top: 0, width: 10, height: 10 });
    const b = new MockShape({ id: "b", zOrderPosition: 1, left: 0, top: 0, width: 10, height: 10 });
    const c = new MockShape({ id: "c", zOrderPosition: 3, left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([a, b, c], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("z-back")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 3 });
    expect(ctx.calls.map((c) => c.id)).toEqual(["a", "c", "b"]);
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

  it("group reports the GeneralException message when PowerPoint rejects the write sync", async () => {
    const a = new MockShape({ id: "a", left: 0, top: 0, width: 10, height: 10 });
    const b = new MockShape({ id: "b", left: 20, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([a, b], "");
    const originalSync = ctx.sync.bind(ctx);
    let calls = 0;
    ctx.sync = async () => {
      calls += 1;
      if (calls === 1) return originalSync();
      throw { code: "GeneralException", message: "x" };
    };
    stubPowerPoint(ctx);

    const cmd = findCommand("group")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({
      ok: false,
      message: "PowerPoint rechazó la operación. Puede que alguna forma esté bloqueada o no admita el cambio.",
    });
    expect(ctx.grouped).toBe(true); // se llamó a group(); fue el sync posterior el que falló
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

  it("inspector-apply ignores undefined fields and merges only the given ones", async () => {
    const s = new MockShape({ id: "s1", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([s], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("inspector-apply")!;
    const result = await executeCommand(cmd, {
      ...DEFAULT_OPTIONS,
      inspector: { left: 100, top: undefined, width: undefined, height: undefined },
    });

    expect(result).toEqual({ ok: true, changed: 1 });
    expect(s.writes).toEqual([{ prop: "left", value: 100 }]);
  });

  it("inspector-apply with an invalid width (0 on a non-degenerate axis) reports a message and writes nothing", async () => {
    const s = new MockShape({ id: "s1", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([s], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("inspector-apply")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, inspector: { width: 0 } });

    expect(result).toEqual({ ok: true, changed: 0, message: "Valores del inspector no válidos." });
    expect(s.writes).toEqual([]);
  });

  it("inspector-apply with a NaN left reports a message and writes nothing", async () => {
    const s = new MockShape({ id: "s1", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([s], "");
    stubPowerPoint(ctx);

    const cmd = findCommand("inspector-apply")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, inspector: { left: NaN } });

    expect(result).toEqual({ ok: true, changed: 0, message: "Valores del inspector no válidos." });
    expect(s.writes).toEqual([]);
  });
});
