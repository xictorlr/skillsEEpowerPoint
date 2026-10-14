import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { getClipboard, resetClipboard } from "../../src/commands/clipboard";
import { executeCommand, resetSlideSizeCache } from "../../src/commands/execute";
import { findCommand } from "../../src/commands/registry";
import type { CommandOptions } from "../../src/commands/types";
import { MockContext, MockShape } from "../office/mockPowerPoint";

const OPTS: CommandOptions = {
  matrix: { columns: 2, gapX: 8, gapY: 8 },
  slideSize: { width: 960, height: 540 },
  colorTarget: "fill",
  pasteParts: { fill: true, line: true, font: true, size: true },
};

let ctx: MockContext;

beforeEach(() => {
  resetClipboard();
  resetSlideSizeCache();
  vi.spyOn(console, "error").mockImplementation(() => {});
  vi.stubGlobal("PowerPoint", { run: (fn: (c: unknown) => Promise<unknown>) => fn(ctx) });
});
afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const box = (id: string, color: string) =>
  new MockShape({ id, type: "TextBox", left: 0, top: 0, width: 10, height: 10,
    fill: { type: "Solid", foregroundColor: color, transparency: 0 },
    font: { name: "Arial", size: 12, color: "#000000", bold: false, italic: false, underline: "None" } });

describe("copy-format / paste-format", () => {
  it("copies the first shape's format and pastes it onto the others with one write sync", async () => {
    const a = box("a", "#FF0000");
    ctx = new MockContext([a], "bad-png");
    const copied = await executeCommand(findCommand("copy-format")!, OPTS);
    expect(copied).toEqual({ ok: true, changed: 0, message: "Formato copiado." });
    expect(getClipboard()?.fill.color).toBe("#FF0000");
    expect(ctx.syncCount).toBe(2); // geometría+formato, fuentes

    const b = box("b", "#00FF00");
    ctx = new MockContext([b], "bad-png");
    const pasted = await executeCommand(findCommand("paste-fill")!, OPTS);
    expect(pasted).toEqual({ ok: true, changed: 1 });
    expect(b.formatWrites).toEqual([{ path: "fill.setSolidColor", value: "#FF0000" }, { path: "fill.transparency", value: 0 }]);
    expect(ctx.syncCount).toBe(3);
  });

  it("refuses to paste without a clipboard", async () => {
    ctx = new MockContext([box("a", "#FF0000")], "bad-png");
    const r = await executeCommand(findCommand("paste-format")!, OPTS);
    expect(r).toEqual({ ok: false, message: "Primero copia el formato de una forma." });
  });
});

describe("color commands", () => {
  it("set-color applies the chosen color to the chosen target", async () => {
    const a = box("a", "#FF0000");
    ctx = new MockContext([a], "bad-png");
    const r = await executeCommand(findCommand("set-color")!, { ...OPTS, color: "#123456", colorTarget: "font" });
    expect(r).toEqual({ ok: true, changed: 1 });
    expect(a.formatWrites).toEqual([{ path: "font.color", value: "#123456" }]);
  });

  it("no-fill clears and transparency-up steps by 0.1", async () => {
    const a = box("a", "#FF0000");
    ctx = new MockContext([a], "bad-png");
    await executeCommand(findCommand("transparency-up")!, OPTS);
    expect(a.formatWrites).toEqual([{ path: "fill.transparency", value: 0.1 }]);
    await executeCommand(findCommand("no-fill")!, OPTS);
    expect(a.formatWrites.at(-1)).toEqual({ path: "fill.clear", value: null });
  });

  it("set-color without a color is a no-op", async () => {
    ctx = new MockContext([box("a", "#FF0000")], "bad-png");
    expect(await executeCommand(findCommand("set-color")!, OPTS)).toEqual({ ok: true, changed: 0 });
  });
});
