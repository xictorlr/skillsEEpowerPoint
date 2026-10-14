import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { executeCommand, resetSlideSizeCache } from "../../src/commands/execute";
import { commands, findCommand } from "../../src/commands/registry";
import type { CommandOptions } from "../../src/commands/types";
import { MockContext, MockShape } from "../office/mockPowerPoint";
import { shape, slide16x9 } from "../core/fixtures";

const DEFAULT_OPTIONS: CommandOptions = {
  matrix: { columns: 2, gapX: 8, gapY: 8 },
  colorTarget: "fill",
  pasteParts: { fill: true, line: true, font: true, size: true },
  fontSizeStep: 1,
  nameFilter: "",
  renamePrefix: "Forma",
  shapeKind: "all",
  agenda: { title: "Agenda", sectionsText: "" },
};

function stubPowerPoint(ctx: MockContext): void {
  vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => Promise<number>) => fn(ctx) });
}

describe("executeCommand", () => {
  beforeEach(() => {
    resetSlideSizeCache();
    vi.spyOn(console, "error").mockImplementation(() => undefined);
    vi.spyOn(console, "warn").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("centers a single shape on the slide, deducing the fallback 16:9 size", async () => {
    const s = new MockShape({ id: "s1", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([s], "bad-png");
    stubPowerPoint(ctx);

    const cmd = findCommand("slide-center")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 1 });
    expect(s.writes).toEqual([
      { prop: "left", value: 475 },
      { prop: "top", value: 265 },
    ]);
    expect(ctx.syncCount).toBe(3); // read selección + intento tamaño de diapositiva + escritura
  });

  it("reports the precondition reason and writes nothing when the selection is too small", async () => {
    const s = new MockShape({ id: "s1", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([s], "bad-png");
    stubPowerPoint(ctx);

    const cmd = findCommand("swap")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: false, message: "Selecciona al menos 2 formas." });
    expect(s.writes).toEqual([]);
  });

  it("reports changed:0 and skips the write sync when nothing moves", async () => {
    const a = new MockShape({ id: "a", left: 0, top: 0, width: 10, height: 10 });
    const b = new MockShape({ id: "b", left: 0, top: 5, width: 10, height: 10 });
    const ctx = new MockContext([a, b], "bad-png");
    stubPowerPoint(ctx);

    const cmd = findCommand("align-left")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 0 });
    expect(a.writes).toEqual([]);
    expect(b.writes).toEqual([]);
    expect(ctx.syncCount).toBe(2); // read selección + intento tamaño de diapositiva, sin sync de escritura
  });

  it("bypasses the slide render entirely when options.slideSize is set", async () => {
    const s = new MockShape({ id: "s1", left: 0, top: 0, width: 10, height: 10 });
    const ctx = new MockContext([s], "bad-png");
    stubPowerPoint(ctx);

    const cmd = findCommand("slide-center")!;
    const options: CommandOptions = { ...DEFAULT_OPTIONS, slideSize: { width: 720, height: 405 } };
    const result = await executeCommand(cmd, options);

    expect(result).toEqual({ ok: true, changed: 1 });
    expect(s.writes).toEqual([
      { prop: "left", value: 355 },
      { prop: "top", value: 197.5 },
    ]);
    expect(ctx.imageOptions).toBeUndefined();
    expect(ctx.syncCount).toBe(2); // read selección + escritura, sin render de diapositiva
  });

  it("every geometry command computes patches referencing only the selected shapes, without throwing", () => {
    const selectionShapes = [
      shape({ id: "a", left: 0, top: 0, width: 10, height: 10 }),
      shape({ id: "b", left: 20, top: 5, width: 15, height: 12 }),
      shape({ id: "c", left: 50, top: 30, width: 8, height: 20 }),
    ];
    const ids = new Set(selectionShapes.map((sh) => sh.id));
    const selection = { shapes: selectionShapes, slide: slide16x9 };

    for (const cmd of commands) {
      if (cmd.kind !== "geometry") continue;
      let patches: ReturnType<typeof cmd.compute> = [];
      expect(() => {
        patches = cmd.compute(selection, DEFAULT_OPTIONS);
      }, `command ${cmd.id} threw`).not.toThrow();
      for (const patch of patches) {
        expect(ids.has(patch.id), `command ${cmd.id} produced a patch for unknown id ${patch.id}`).toBe(true);
      }
    }
  });
});
