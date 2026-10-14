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
  agenda: { title: "Agenda", sectionsText: "" },
  insert: { count: 3, stickerText: "BORRADOR", stickerColor: "#FFE066" },
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

  it("registers the 30 text commands visible without API 1.10 in the section, in order (indent hidden)", () => {
    expect(commandsBySection("text").map((c) => c.id)).toEqual([
      "case-upper", "case-lower", "case-sentence", "case-title",
      "style-bold", "style-italic", "style-underline", "style-strike",
      "font-bigger", "font-smaller",
      "size-title", "size-subtitle", "size-body", "size-small",
      "para-left", "para-center", "para-right", "para-justify",
      "bullets-toggle",
      "valign-top", "valign-middle", "valign-bottom",
      "margins-zero", "margins-standard",
      "autosize-none", "autosize-shrink", "autosize-shape",
      "wrap-on", "wrap-off",
      "clean-text",
    ]);
    expect(findCommand("indent-more")).toBeDefined();
    expect(findCommand("indent-less")).toBeDefined();
  });

  it("includes indent-more/indent-less right after bullets-toggle when the host supports PowerPointApi 1.10", () => {
    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: () => true } } });
    const ids = commandsBySection("text").map((c) => c.id);
    expect(ids).toHaveLength(32);
    expect(ids.slice(ids.indexOf("bullets-toggle"), ids.indexOf("bullets-toggle") + 3)).toEqual(["bullets-toggle", "indent-more", "indent-less"]);
    expect(findCommand("indent-more")!.requiresApi).toBe("1.10");
    expect(findCommand("indent-less")!.requiresApi).toBe("1.10");
  });

  describe("quick styles (Fase 8)", () => {
    const boldBox = (id: string, bold: boolean) =>
      new MockShape({
        id, type: "TextBox", left: 0, top: 0, width: 10, height: 10,
        font: { name: "Calibri", size: 18, color: "#000000", bold, italic: false, underline: "None" },
      });

    it("style-bold on a mixed selection sets bold on the shapes that were not bold", async () => {
      const a = boldBox("a", true);
      const b = boldBox("b", false);
      ctx = new MockContext([a, b], "bad-png");

      const r = await executeCommand(findCommand("style-bold")!, OPTS);

      expect(r).toEqual({ ok: true, changed: 1 });
      expect(a.formatWrites.map((w) => w.path)).not.toContain("font.bold");
      expect(b.formatWrites).toContainEqual({ path: "font.bold", value: true });
    });

    it("style-bold on an all-bold selection clears bold on every shape", async () => {
      const a = boldBox("a", true);
      const b = boldBox("b", true);
      ctx = new MockContext([a, b], "bad-png");

      const r = await executeCommand(findCommand("style-bold")!, OPTS);

      expect(r).toEqual({ ok: true, changed: 2 });
      expect(a.formatWrites).toContainEqual({ path: "font.bold", value: false });
      expect(b.formatWrites).toContainEqual({ path: "font.bold", value: false });
    });

    it("style-italic / style-underline / bullets-toggle flip the corresponding property (defaults: off, off, on)", async () => {
      const a = textBox("a");
      ctx = new MockContext([a], "bad-png");
      await executeCommand(findCommand("style-italic")!, OPTS);
      expect(a.formatWrites).toContainEqual({ path: "font.italic", value: true });

      const b = textBox("b");
      ctx = new MockContext([b], "bad-png");
      await executeCommand(findCommand("style-underline")!, OPTS);
      expect(b.formatWrites).toContainEqual({ path: "font.underline", value: "Single" });

      const c = textBox("c"); // bulletsVisible por defecto del mock: true -> se apagan
      ctx = new MockContext([c], "bad-png");
      await executeCommand(findCommand("bullets-toggle")!, OPTS);
      expect(c.formatWrites).toContainEqual({ path: "bulletFormat.visible", value: false });
    });

    it("size-title writes font.size 24; size-small writes 11", async () => {
      const a = textBox("a"); // 18 pt por defecto
      ctx = new MockContext([a], "bad-png");
      const r = await executeCommand(findCommand("size-title")!, OPTS);
      expect(r).toEqual({ ok: true, changed: 1 });
      expect(a.formatWrites).toContainEqual({ path: "font.size", value: 24 });

      const b = textBox("b");
      ctx = new MockContext([b], "bad-png");
      await executeCommand(findCommand("size-small")!, OPTS);
      expect(b.formatWrites).toContainEqual({ path: "font.size", value: 11 });
    });

    it("size-subtitle (18) on an 18 pt box reports nothing to change", async () => {
      const a = textBox("a");
      ctx = new MockContext([a], "bad-png");
      const r = await executeCommand(findCommand("size-subtitle")!, OPTS);
      expect(r).toEqual({ ok: true, changed: 0 });
    });

    it("indent-more writes indentLevel + 1 with API 1.10 (indent-less at level 0 changes nothing)", async () => {
      vi.stubGlobal("Office", {
        context: { requirements: { isSetSupported: (name: string, version: string) => name === "PowerPointApi" && version === "1.10" } },
      });
      const a = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10, textProps: { indentLevel: 1 } });
      ctx = new MockContext([a], "bad-png");
      const r = await executeCommand(findCommand("indent-more")!, OPTS);
      expect(r).toEqual({ ok: true, changed: 1 });
      expect(a.formatWrites).toContainEqual({ path: "paragraphFormat.indentLevel", value: 2 });

      const b = textBox("b"); // indentLevel 0
      ctx = new MockContext([b], "bad-png");
      expect(await executeCommand(findCommand("indent-less")!, OPTS)).toEqual({ ok: true, changed: 0 });
    });

    it("indent-more fails with the API reason when the host lacks PowerPointApi 1.10", async () => {
      const a = textBox("a");
      ctx = new MockContext([a], "bad-png");
      const r = await executeCommand(findCommand("indent-more")!, OPTS);
      expect(r.ok).toBe(false);
      if (!r.ok) expect(r.message).toBe("Requiere PowerPoint reciente (API 1.10).");
    });
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
      { id: "a", text: "hola mundo", fontSize: 18, horizontalAlignment: "Left", verticalAlignment: "Top", autoSize: "AutoSizeNone", wordWrap: true, margins: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
        bold: false, italic: false, underline: "None", strikethrough: false, bulletsVisible: true, indentLevel: 0 },
      { id: "b", text: "otro texto", fontSize: 24, horizontalAlignment: "Center", verticalAlignment: "Middle", autoSize: "AutoSizeTextToFitShape", wordWrap: false, margins: { left: 0, right: 0, top: 0, bottom: 0 },
        bold: true, italic: true, underline: "Single", strikethrough: true, bulletsVisible: false, indentLevel: 3 },
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
