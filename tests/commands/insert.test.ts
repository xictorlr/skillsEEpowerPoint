import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { executeCommand, resetSlideSizeCache } from "../../src/commands/execute";
import { commandsBySection, findCommand } from "../../src/commands/registry";
import type { CommandOptions } from "../../src/commands/types";
import { BASIC_SHAPE_RECT, chevronChainLayout, columnsLayout, gridLayout, numberedCirclesLayout, stickerRect } from "../../src/core/insert";
import { MockContext, type MockShape } from "../office/mockPowerPoint";

const SLIDE_SIZE = { width: 960, height: 540 };

const DEFAULT_OPTIONS: CommandOptions = {
  matrix: { columns: 2, gapX: 8, gapY: 8 },
  slideSize: SLIDE_SIZE,
  colorTarget: "fill",
  pasteParts: { fill: true, line: true, font: true, size: true },
  fontSizeStep: 1,
  nameFilter: "",
  renamePrefix: "Forma",
  shapeKind: "all",
  agenda: { title: "Agenda", sectionsText: "" },
  insert: { count: 3, stickerText: "BORRADOR", stickerColor: "#FFE066" },
};

function stubPowerPoint(ctx: MockContext): void {
  vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => Promise<unknown>) => fn(ctx) });
}

function stubOffice1_10(): void {
  vi.stubGlobal("Office", {
    context: {
      requirements: {
        isSetSupported: (name: string, version: string) => name === "PowerPointApi" && version === "1.10",
      },
    },
  });
}

/** Contexto con una única diapositiva "sl1" (activa), sin formas previas. */
function newCtx(opts: Partial<ConstructorParameters<typeof MockContext>[2]> = {}): MockContext {
  const ctx = new MockContext([], "", { slides: [{ id: "sl1" }], ...opts });
  stubPowerPoint(ctx);
  return ctx;
}

const shapesOf = (ctx: MockContext): MockShape[] => ctx.slidesList[0]!.shapesList;

/**
 * Syncs esperados con `options.slideSize` forzado (sin render PNG) y sin PowerPointApi 1.10 (sin
 * lectura del tema): 1 de `readSlideNav` + 2 de escritura (alta + carga de ids, y selección).
 */
const SYNCS_BASE = 3;

describe("insert commands", () => {
  beforeEach(() => {
    resetSlideSizeCache();
    vi.spyOn(console, "error").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("registers the 11 insert commands in the section, in order, none hidden and none needing a selection", () => {
    const cmds = commandsBySection("insert");
    expect(cmds.map((c) => c.id)).toEqual([
      "insert-columns",
      "insert-grid",
      "insert-chevrons",
      "insert-circles",
      "insert-sticker",
      "insert-line",
      "insert-elbow",
      "insert-rect",
      "insert-rounded",
      "insert-ellipse",
      "insert-textbox",
    ]);
    for (const cmd of cmds) {
      expect(cmd.kind).toBe("action");
      expect(cmd.minShapes).toBe(0);
      expect(cmd.kind === "action" && cmd.needsSelection).toBe(false);
    }
  });

  describe("insert-columns", () => {
    it("inserts options.insert.count text boxes with 'Texto' on the active slide, selects them and reports", async () => {
      const ctx = newCtx();
      const result = await executeCommand(findCommand("insert-columns")!, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 3, message: "3 columna(s) de texto insertada(s)." });
      const expected = columnsLayout(SLIDE_SIZE, 3);
      const adds = ctx.calls.filter((c) => c.path === "addTextBox");
      expect(adds).toHaveLength(3);
      adds.forEach((call, i) => expect(call.value).toEqual({ text: "Texto", ...expected[i] }));
      expect(ctx.selectedShapeIds).toEqual(shapesOf(ctx).map((s) => s.id));
      expect(ctx.selectedShapeIds).toHaveLength(3);
      // Cada forma nueva carga su id antes del primer sync de escritura (para poder seleccionarla).
      for (const s of shapesOf(ctx)) expect(s.loads).toContainEqual({ path: "shape", props: "id" });
      expect(ctx.syncCount).toBe(SYNCS_BASE);
    });

    it("honours a different count (5)", async () => {
      const ctx = newCtx();
      const result = await executeCommand(findCommand("insert-columns")!, { ...DEFAULT_OPTIONS, insert: { ...DEFAULT_OPTIONS.insert, count: 5 } });

      expect(result).toEqual({ ok: true, changed: 5, message: "5 columna(s) de texto insertada(s)." });
      expect(ctx.calls.filter((c) => c.path === "addTextBox")).toHaveLength(5);
      expect(ctx.selectedShapeIds).toHaveLength(5);
    });

    it("clamps the count to 12", async () => {
      const ctx = newCtx();
      const result = await executeCommand(findCommand("insert-columns")!, { ...DEFAULT_OPTIONS, insert: { ...DEFAULT_OPTIONS.insert, count: 20 } });

      expect(result).toEqual({ ok: true, changed: 12, message: "12 columna(s) de texto insertada(s)." });
      expect(ctx.calls.filter((c) => c.path === "addTextBox")).toHaveLength(12);
    });
  });

  it("insert-grid inserts a fixed 2×2 grid of text boxes regardless of count", async () => {
    const ctx = newCtx();
    const result = await executeCommand(findCommand("insert-grid")!, { ...DEFAULT_OPTIONS, insert: { ...DEFAULT_OPTIONS.insert, count: 7 } });

    expect(result).toEqual({ ok: true, changed: 4, message: "Rejilla 2×2 insertada (4 cuadros)." });
    const expected = gridLayout(SLIDE_SIZE, 2, 2);
    const adds = ctx.calls.filter((c) => c.path === "addTextBox");
    expect(adds).toHaveLength(4);
    adds.forEach((call, i) => expect(call.value).toEqual({ text: "Texto", ...expected[i] }));
    expect(ctx.selectedShapeIds).toHaveLength(4);
    expect(ctx.syncCount).toBe(SYNCS_BASE);
  });

  describe("insert-chevrons", () => {
    it("inserts count chevrons labelled 'Paso i' with the default accent fill, white bold 12pt centered text and no line", async () => {
      const ctx = newCtx();
      const result = await executeCommand(findCommand("insert-chevrons")!, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 3, message: "3 chevrón(es) insertado(s)." });
      const expected = chevronChainLayout(SLIDE_SIZE, 3);
      const adds = ctx.calls.filter((c) => c.path === "addGeometricShape");
      expect(adds).toHaveLength(3);
      adds.forEach((call, i) => expect(call.value).toEqual({ type: "Chevron", ...expected[i] }));

      const shapes = shapesOf(ctx);
      shapes.forEach((s, i) => {
        expect(s.formatWrites).toContainEqual({ path: "fill.setSolidColor", value: "#1F6FEB" }); // sin tema: acento por defecto
        expect(s.formatWrites).toContainEqual({ path: "lineFormat.visible", value: false });
        expect(s.formatWrites).toContainEqual({ path: "textRange.text", value: `Paso ${i + 1}` });
        expect(s.formatWrites).toContainEqual({ path: "font.color", value: "#FFFFFF" });
        expect(s.formatWrites).toContainEqual({ path: "font.bold", value: true });
        expect(s.formatWrites).toContainEqual({ path: "font.size", value: 12 });
        expect(s.formatWrites).toContainEqual({ path: "paragraphFormat.horizontalAlignment", value: "Center" });
      });
      expect(ctx.selectedShapeIds).toHaveLength(3);
      expect(ctx.syncCount).toBe(SYNCS_BASE);
    });

    it("uses the theme's Accent1 (API 1.10) when available, at the cost of one extra read sync", async () => {
      stubOffice1_10();
      const ctx = newCtx({ themeColors: { Accent1: "#00AA00" } });
      await executeCommand(findCommand("insert-chevrons")!, DEFAULT_OPTIONS);

      for (const s of shapesOf(ctx)) expect(s.formatWrites).toContainEqual({ path: "fill.setSolidColor", value: "#00AA00" });
      // + 1 sync de `getThemeColors` (el tamaño sigue forzado por `options.slideSize`, sin sync).
      expect(ctx.syncCount).toBe(SYNCS_BASE + 1);
    });
  });

  it("insert-circles inserts count numbered ellipses with accent fill, white bold 14pt centered text", async () => {
    const ctx = newCtx();
    const result = await executeCommand(findCommand("insert-circles")!, { ...DEFAULT_OPTIONS, insert: { ...DEFAULT_OPTIONS.insert, count: 4 } });

    expect(result).toEqual({ ok: true, changed: 4, message: "4 círculo(s) numerado(s) insertado(s)." });
    const expected = numberedCirclesLayout(SLIDE_SIZE, 4);
    const adds = ctx.calls.filter((c) => c.path === "addGeometricShape");
    expect(adds).toHaveLength(4);
    adds.forEach((call, i) => expect(call.value).toEqual({ type: "Ellipse", ...expected[i] }));
    shapesOf(ctx).forEach((s, i) => {
      expect(s.formatWrites).toContainEqual({ path: "fill.setSolidColor", value: "#1F6FEB" });
      expect(s.formatWrites).toContainEqual({ path: "lineFormat.visible", value: false });
      expect(s.formatWrites).toContainEqual({ path: "textRange.text", value: String(i + 1) });
      expect(s.formatWrites).toContainEqual({ path: "font.color", value: "#FFFFFF" });
      expect(s.formatWrites).toContainEqual({ path: "font.bold", value: true });
      expect(s.formatWrites).toContainEqual({ path: "font.size", value: 14 });
      expect(s.formatWrites).toContainEqual({ path: "paragraphFormat.horizontalAlignment", value: "Center" });
    });
    expect(ctx.selectedShapeIds).toHaveLength(4);
    expect(ctx.syncCount).toBe(SYNCS_BASE);
  });

  describe("insert-sticker", () => {
    it("inserts a rounded rectangle in the top-right corner with the configured text and fill color", async () => {
      const ctx = newCtx();
      const result = await executeCommand(findCommand("insert-sticker")!, {
        ...DEFAULT_OPTIONS,
        insert: { count: 3, stickerText: "CONFIDENCIAL", stickerColor: "#FF0000" },
      });

      expect(result).toEqual({ ok: true, changed: 1, message: "Sticker insertado." });
      expect(ctx.calls.filter((c) => c.path === "addGeometricShape")).toEqual([
        { id: "new-shape-1", path: "addGeometricShape", value: { type: "RoundRectangle", ...stickerRect(SLIDE_SIZE) } },
      ]);
      const [s] = shapesOf(ctx);
      expect(s!.formatWrites).toContainEqual({ path: "fill.setSolidColor", value: "#FF0000" });
      expect(s!.formatWrites).toContainEqual({ path: "lineFormat.visible", value: false });
      expect(s!.formatWrites).toContainEqual({ path: "textRange.text", value: "CONFIDENCIAL" });
      expect(s!.formatWrites).toContainEqual({ path: "font.bold", value: true });
      expect(s!.formatWrites).toContainEqual({ path: "font.size", value: 12 });
      expect(s!.formatWrites).toContainEqual({ path: "font.color", value: "#000000" });
      expect(s!.formatWrites).toContainEqual({ path: "paragraphFormat.horizontalAlignment", value: "Center" });
      expect(ctx.selectedShapeIds).toEqual(["new-shape-1"]);
      expect(ctx.syncCount).toBe(SYNCS_BASE);
    });

    it("falls back to 'BORRADOR' when stickerText is empty", async () => {
      const ctx = newCtx();
      await executeCommand(findCommand("insert-sticker")!, { ...DEFAULT_OPTIONS, insert: { ...DEFAULT_OPTIONS.insert, stickerText: "" } });

      expect(shapesOf(ctx)[0]!.formatWrites).toContainEqual({ path: "textRange.text", value: "BORRADOR" });
    });
  });

  it.each([
    ["insert-line", "Straight", "Línea insertada."],
    ["insert-elbow", "Elbow", "Conector insertado."],
  ])("%s adds a %s line in the basic centered rect, without styling it", async (id, kind, message) => {
    const ctx = newCtx();
    const result = await executeCommand(findCommand(id)!, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 1, message });
    expect(ctx.calls.filter((c) => c.path === "addLine")).toEqual([
      { id: "new-shape-1", path: "addLine", value: { kind, ...BASIC_SHAPE_RECT(SLIDE_SIZE) } },
    ]);
    expect(shapesOf(ctx)[0]!.formatWrites).toEqual([]);
    expect(ctx.selectedShapeIds).toEqual(["new-shape-1"]);
    expect(ctx.syncCount).toBe(SYNCS_BASE);
  });

  it.each([
    ["insert-rect", "Rectangle"],
    ["insert-rounded", "RoundRectangle"],
    ["insert-ellipse", "Ellipse"],
  ])("%s adds a %s with theme defaults (no fill/line/font writes)", async (id, type) => {
    const ctx = newCtx();
    const result = await executeCommand(findCommand(id)!, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 1, message: "Forma insertada." });
    expect(ctx.calls.filter((c) => c.path === "addGeometricShape")).toEqual([
      { id: "new-shape-1", path: "addGeometricShape", value: { type, ...BASIC_SHAPE_RECT(SLIDE_SIZE) } },
    ]);
    const paths = shapesOf(ctx)[0]!.formatWrites.map((w) => w.path);
    expect(paths).not.toContain("fill.setSolidColor");
    expect(paths).not.toContain("fill.clear");
    expect(paths).not.toContain("lineFormat.visible");
    expect(paths).not.toContain("textRange.text");
    expect(paths.some((p) => p.startsWith("font."))).toBe(false);
    expect(ctx.selectedShapeIds).toEqual(["new-shape-1"]);
    expect(ctx.syncCount).toBe(SYNCS_BASE);
  });

  it("insert-textbox adds a single 'Texto' box in the basic centered rect", async () => {
    const ctx = newCtx();
    const result = await executeCommand(findCommand("insert-textbox")!, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 1, message: "Cuadro de texto insertado." });
    expect(ctx.calls.filter((c) => c.path === "addTextBox")).toEqual([
      { id: "new-shape-1", path: "addTextBox", value: { text: "Texto", ...BASIC_SHAPE_RECT(SLIDE_SIZE) } },
    ]);
    expect(ctx.selectedShapeIds).toEqual(["new-shape-1"]);
    expect(ctx.syncCount).toBe(SYNCS_BASE);
  });

  it("inserts on the active slide, not the first one, when another slide is selected", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "sl1" }, { id: "sl2" }, { id: "sl3" }], activeSlideIndex: 2 });
    stubPowerPoint(ctx);
    await executeCommand(findCommand("insert-textbox")!, DEFAULT_OPTIONS);

    expect(ctx.slidesList[0]!.shapesList).toHaveLength(0);
    expect(ctx.slidesList[2]!.shapesList).toHaveLength(1);
  });
});
