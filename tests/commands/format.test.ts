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
  fontSizeStep: 1,
  nameFilter: "",
  renamePrefix: "Forma",
  shapeKind: "all",
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
    expect(getClipboard()?.fill?.color).toBe("#FF0000");
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

  it("copy-format reports a 'sin fuente' message for a shape without a text frame (e.g. a picture)", async () => {
    const pic = new MockShape({ id: "pic", type: "Image", left: 0, top: 0, width: 10, height: 10,
      fill: { type: "Solid", foregroundColor: "#FF0000", transparency: 0 }, font: null });
    ctx = new MockContext([pic], "bad-png");
    const r = await executeCommand(findCommand("copy-format")!, OPTS);
    expect(r).toEqual({ ok: true, changed: 0, message: "Formato copiado (sin fuente: la forma no tiene texto)." });
    expect(getClipboard()?.font).toBeNull();
  });

  it("paste-format onto mixed targets: no font for a picture, no width/height when the size already matches", async () => {
    const a = box("a", "#FF0000"); // TextBox 10x10
    ctx = new MockContext([a], "bad-png");
    await executeCommand(findCommand("copy-format")!, OPTS);

    const pic = new MockShape({ id: "pic", type: "Image", left: 0, top: 0, width: 10, height: 10,
      fill: { type: "Solid", foregroundColor: "#0000FF", transparency: 0 }, font: null });
    const big = new MockShape({ id: "big", type: "TextBox", left: 0, top: 0, width: 20, height: 30,
      fill: { type: "Solid", foregroundColor: "#0000FF", transparency: 0 },
      font: { name: "Verdana", size: 20, color: "#FFFFFF", bold: false, italic: false, underline: "None" } });
    ctx = new MockContext([pic, big], "bad-png");
    const r = await executeCommand(findCommand("paste-format")!, OPTS);

    expect(r).toEqual({ ok: true, changed: 2 });
    expect(pic.formatWrites.some((w) => w.path.startsWith("font."))).toBe(false);
    expect(pic.writes).toEqual([]); // misma medida (10x10) que la fuente: sin escritura de tamaño
    expect(big.formatWrites.some((w) => w.path.startsWith("font."))).toBe(true);
    expect(big.writes).toEqual([
      { prop: "width", value: 10 },
      { prop: "height", value: 10 },
    ]);
  });

  it("reports the number of distinct shapes changed across multiple targets", async () => {
    const a = box("a", "#FF0000");
    ctx = new MockContext([a], "bad-png");
    await executeCommand(findCommand("copy-format")!, OPTS);

    const b = box("b", "#00FF00");
    const c = box("c", "#0000FF");
    const d = box("d", "#ABCDEF");
    ctx = new MockContext([b, c, d], "bad-png");
    const r = await executeCommand(findCommand("paste-fill")!, OPTS);
    expect(r).toEqual({ ok: true, changed: 3 });
  });

  it("returns ok:false when a sync rejects while handling format for a shape PowerPoint refuses", async () => {
    const a = box("a", "#FF0000");
    ctx = new MockContext([a], "bad-png");
    await executeCommand(findCommand("copy-format")!, OPTS);

    const bad = new MockShape({ id: "bad", type: "TextBox", left: 0, top: 0, width: 10, height: 10, supportsFormat: false,
      fill: { type: "Solid", foregroundColor: "#00FF00", transparency: 0 },
      font: { name: "Arial", size: 12, color: "#000000", bold: false, italic: false, underline: "None" } });
    ctx = new MockContext([bad], "bad-png");
    const r = await executeCommand(findCommand("paste-fill")!, OPTS);
    expect(r.ok).toBe(false);
  });
});

describe("no-line and transparency on a line-only shape", () => {
  it("clears and steps the line of a shape that doesn't support fill (type Line)", async () => {
    const line = new MockShape({ id: "l", type: "Line", left: 0, top: 0, width: 10, height: 0, font: null,
      line: { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 } });
    ctx = new MockContext([line], "bad-png");

    const up = await executeCommand(findCommand("transparency-up")!, { ...OPTS, colorTarget: "line" });
    expect(up).toEqual({ ok: true, changed: 1 });
    expect(line.formatWrites).toEqual([{ path: "lineFormat.transparency", value: 0.1 }]);

    const cleared = await executeCommand(findCommand("no-line")!, OPTS);
    expect(cleared).toEqual({ ok: true, changed: 1 });
    expect(line.formatWrites.at(-1)).toEqual({ path: "lineFormat.visible", value: false });
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
