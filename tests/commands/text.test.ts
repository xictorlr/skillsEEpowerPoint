import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { executeCommand, resetSlideSizeCache } from "../../src/commands/execute";
import { commands, commandsBySection, findCommand } from "../../src/commands/registry";
import type { CommandOptions } from "../../src/commands/types";
import type { TextShape } from "../../src/core/text";
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
  resetSlideSizeCache();
  vi.spyOn(console, "error").mockImplementation(() => {});
  vi.stubGlobal("PowerPoint", { run: (fn: (c: unknown) => Promise<unknown>) => fn(ctx) });
});
afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const textBox = (id: string) => new MockShape({ id, type: "TextBox", left: 0, top: 0, width: 10, height: 10 });

describe("text commands", () => {
  it("case-upper writes textRange.text on a TextBox with 3 syncs", async () => {
    const a = textBox("a"); // texto por defecto del mock: "hola mundo"
    ctx = new MockContext([a], "bad-png");

    const r = await executeCommand(findCommand("case-upper")!, OPTS);

    expect(r).toEqual({ ok: true, changed: 1 });
    // "hola mundo" -> "HOLA MUNDO": dos ediciones de palabra, en orden descendente por start
    // ("MUNDO" 5..10 antes que "HOLA" 0..4) para no desplazar el offset pendiente.
    expect(a.formatWrites).toEqual([
      { path: "textRange.substring", value: { start: 5, length: 5, text: "MUNDO" } },
      { path: "textRange.substring", value: { start: 0, length: 4, text: "HOLA" } },
    ]);
    expect(ctx.syncCount).toBe(3); // lectura selección, lectura texto, escritura
  });

  it("font-bigger adds fontSizeStep points to font.size", async () => {
    const a = textBox("a"); // tamaño de fuente por defecto del mock: 18
    ctx = new MockContext([a], "bad-png");

    const r = await executeCommand(findCommand("font-bigger")!, { ...OPTS, fontSizeStep: 2 });

    expect(r).toEqual({ ok: true, changed: 1 });
    expect(a.formatWrites).toContainEqual({ path: "font.size", value: 20 });
  });

  it("reports no text for a shape without a text frame (e.g. an image)", async () => {
    const pic = new MockShape({ id: "pic", type: "Image", left: 0, top: 0, width: 10, height: 10, font: null });
    ctx = new MockContext([pic], "bad-png");

    const r = await executeCommand(findCommand("case-upper")!, OPTS);

    expect(r).toEqual({ ok: true, changed: 0, message: "La selección no tiene texto." });
  });

  it("registers the 21 text commands in the section, in order", () => {
    expect(commandsBySection("text").map((c) => c.id)).toEqual([
      "case-upper", "case-lower", "case-sentence", "case-title",
      "font-bigger", "font-smaller",
      "para-left", "para-center", "para-right", "para-justify",
      "valign-top", "valign-middle", "valign-bottom",
      "margins-zero", "margins-standard",
      "autosize-none", "autosize-shrink", "autosize-shape",
      "wrap-on", "wrap-off",
      "clean-text",
    ]);
  });

  it("font-smaller subtracts fontSizeStep points from font.size", async () => {
    const a = textBox("a"); // tamaño de fuente por defecto del mock: 18
    ctx = new MockContext([a], "bad-png");

    const r = await executeCommand(findCommand("font-smaller")!, { ...OPTS, fontSizeStep: 2 });

    expect(r).toEqual({ ok: true, changed: 1 });
    expect(a.formatWrites).toContainEqual({ path: "font.size", value: 16 });
  });

  it("reports changed:0 with no message when there is text but no applicable change (mixed/null font size)", async () => {
    const a = new MockShape({
      id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10,
      font: { name: "Arial", size: null, color: "#000000", bold: false, italic: false, underline: "None" },
    });
    ctx = new MockContext([a], "bad-png");

    const r = await executeCommand(findCommand("font-bigger")!, OPTS);

    expect(r).toEqual({ ok: true, changed: 0 });
  });

  it("every text command computes patches referencing only the selected shapes, without throwing", () => {
    const shapes: TextShape[] = [
      { id: "a", text: "hola mundo", fontSize: 18, horizontalAlignment: "Left", verticalAlignment: "Top", autoSize: "AutoSizeNone", wordWrap: true, margins: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 } },
      { id: "b", text: "otro texto", fontSize: 24, horizontalAlignment: "Center", verticalAlignment: "Middle", autoSize: "AutoSizeTextToFitShape", wordWrap: false, margins: { left: 0, right: 0, top: 0, bottom: 0 } },
    ];
    const ids = new Set(shapes.map((s) => s.id));

    for (const cmd of commands) {
      if (cmd.kind !== "text") continue;
      let patches: ReturnType<typeof cmd.compute> = [];
      expect(() => {
        patches = cmd.compute(shapes, OPTS);
      }, `command ${cmd.id} threw`).not.toThrow();
      for (const patch of patches) {
        expect(ids.has(patch.id), `command ${cmd.id} produced a patch for unknown id ${patch.id}`).toBe(true);
      }
    }
  });
});
