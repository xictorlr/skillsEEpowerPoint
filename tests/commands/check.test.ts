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
  agenda: { title: "Agenda", sectionsText: "" },
};

function stubPowerPoint(ctx: MockContext): void {
  vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => Promise<unknown>) => fn(ctx) });
}

describe("check commands", () => {
  beforeEach(() => {
    resetSlideSizeCache();
    vi.spyOn(console, "error").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  describe("check-empty-text-slide", () => {
    it("reports 'no hay cuadros vacíos' when the active slide has no empty text boxes", async () => {
      const filled = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
      const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [filled] }] });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-empty-text-slide")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 0, message: "No hay cuadros vacíos." });
    });

    it("deletes only the empty text boxes of the active slide, leaving other slides untouched", async () => {
      const emptyOnActive = new MockShape({
        id: "empty-1",
        type: "TextBox",
        left: 0,
        top: 0,
        width: 10,
        height: 10,
        textProps: { hasText: false },
      });
      const filledOnActive = new MockShape({ id: "filled-1", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
      const emptyOnOtherSlide = new MockShape({
        id: "empty-2",
        type: "TextBox",
        left: 0,
        top: 0,
        width: 10,
        height: 10,
        textProps: { hasText: false },
      });
      const ctx = new MockContext([], "", {
        slides: [
          { id: "sl1", shapes: [emptyOnActive, filledOnActive] },
          { id: "sl2", shapes: [emptyOnOtherSlide] },
        ],
        activeSlideIndex: 0,
      });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-empty-text-slide")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 1, message: "1 cuadro(s) vacío(s) eliminado(s)." });
      expect(emptyOnActive.formatWrites).toContainEqual({ path: "delete", value: null });
      expect(filledOnActive.formatWrites).not.toContainEqual({ path: "delete", value: null });
      expect(emptyOnOtherSlide.formatWrites).not.toContainEqual({ path: "delete", value: null });
    });

    it("does not consider non-TextBox shapes, even without text", async () => {
      const img = new MockShape({ id: "img1", type: "Image", left: 0, top: 0, width: 10, height: 10, font: null });
      const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [img] }] });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-empty-text-slide")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 0, message: "No hay cuadros vacíos." });
      expect(ctx.invalidTouched).toBe(false);
    });

    it("does not delete an empty Placeholder (only TextBox shapes are considered)", async () => {
      const emptyPlaceholder = new MockShape({
        id: "ph1",
        type: "Placeholder",
        containedType: null, // contenido libre: tiene marco de texto, pero no es un TextBox
        left: 0,
        top: 0,
        width: 10,
        height: 10,
        textProps: { hasText: false },
      });
      const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [emptyPlaceholder] }] });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-empty-text-slide")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 0, message: "No hay cuadros vacíos." });
      expect(emptyPlaceholder.formatWrites).not.toContainEqual({ path: "delete", value: null });
    });
  });

  describe("check-empty-text-deck", () => {
    it("is flagged for confirmation", () => {
      expect(findCommand("check-empty-text-deck")!.confirm).toBe(true);
    });

    it("reports 'no hay cuadros vacíos' when nothing in the deck is empty", async () => {
      const filled = new MockShape({ id: "a", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
      const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [filled] }] });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-empty-text-deck")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 0, message: "No hay cuadros vacíos." });
    });

    it("deletes empty text boxes across every slide of the deck", async () => {
      const emptyOnSlide1 = new MockShape({
        id: "empty-1",
        type: "TextBox",
        left: 0,
        top: 0,
        width: 10,
        height: 10,
        textProps: { hasText: false },
      });
      const emptyOnSlide2 = new MockShape({
        id: "empty-2",
        type: "TextBox",
        left: 0,
        top: 0,
        width: 10,
        height: 10,
        textProps: { hasText: false },
      });
      const filled = new MockShape({ id: "filled-1", type: "TextBox", left: 0, top: 0, width: 10, height: 10 });
      const ctx = new MockContext([], "", {
        slides: [
          { id: "sl1", shapes: [emptyOnSlide1, filled] },
          { id: "sl2", shapes: [emptyOnSlide2] },
        ],
      });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-empty-text-deck")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 2, message: "2 cuadro(s) vacío(s) eliminado(s)." });
      expect(emptyOnSlide1.formatWrites).toContainEqual({ path: "delete", value: null });
      expect(emptyOnSlide2.formatWrites).toContainEqual({ path: "delete", value: null });
      expect(filled.formatWrites).not.toContainEqual({ path: "delete", value: null });
    });

    it("does not delete an empty Placeholder either", async () => {
      const emptyPlaceholder = new MockShape({
        id: "ph1",
        type: "Placeholder",
        containedType: null,
        left: 0,
        top: 0,
        width: 10,
        height: 10,
        textProps: { hasText: false },
      });
      const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [emptyPlaceholder] }] });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-empty-text-deck")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 0, message: "No hay cuadros vacíos." });
      expect(emptyPlaceholder.formatWrites).not.toContainEqual({ path: "delete", value: null });
    });
  });

  describe("check-offslide-select", () => {
    const SLIDE_SIZE = { width: 960, height: 540 };

    it("reports 'no hay formas fuera de la diapositiva' and leaves the selection untouched", async () => {
      const onSlide = new MockShape({ id: "a", left: 0, top: 0, width: 10, height: 10 });
      const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [onSlide] }] });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-offslide-select")!;
      const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, slideSize: SLIDE_SIZE });

      expect(result).toEqual({ ok: true, changed: 0, message: "No hay formas fuera de la diapositiva." });
      expect(ctx.selectedShapeIds).toBeUndefined();
    });

    it("selects every off-slide shape of the active slide", async () => {
      const inside = new MockShape({ id: "inside", left: 0, top: 0, width: 10, height: 10 });
      const offRight = new MockShape({ id: "off-right", left: 2000, top: 0, width: 10, height: 10 });
      const offBottom = new MockShape({ id: "off-bottom", left: 0, top: 2000, width: 10, height: 10 });
      const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [inside, offRight, offBottom] }] });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-offslide-select")!;
      const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, slideSize: SLIDE_SIZE });

      expect(result).toEqual({ ok: true, changed: 2, message: "2 forma(s) fuera de la diapositiva seleccionada(s)." });
      expect(ctx.selectedShapeIds).toEqual(["off-right", "off-bottom"]);
    });

    // PNG 1x1 (aspecto 1:1): `getSlideSize` lo renderiza y deduce 540x540 (ver tests/office/selection.test.ts).
    const PNG_1X1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

    it("shares the slideSize cache across action-command calls in the non-1.10 path (ctx.slideSizeCache, not a fresh Map per call); con PowerPointApi 1.10, getSlideSize ya no cachea (lee pageSetup en cada llamada, ver slideSize.ts)", async () => {
      const inside = new MockShape({ id: "inside", left: 0, top: 0, width: 10, height: 10 });
      const ctx1 = new MockContext([], PNG_1X1, { slides: [{ id: "sl1", shapes: [inside] }] });
      stubPowerPoint(ctx1);

      const cmd = findCommand("check-offslide-select")!;
      await executeCommand(cmd, DEFAULT_OPTIONS); // sin override ni API 1.10: renderiza y cachea 540x540

      // Un segundo contexto con un PNG inválido (forzaría el fallback 960x540 si se releyera): si
      // la caché se comparte de verdad, este contexto ni siquiera debería intentar renderizar --
      // el tamaño cacheado (540x540) decide si la forma queda fuera.
      const offAt600 = new MockShape({ id: "off-at-600", left: 600, top: 0, width: 10, height: 10 });
      const ctx2 = new MockContext([], "bad-png", { slides: [{ id: "sl1", shapes: [offAt600] }] });
      stubPowerPoint(ctx2);

      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 1, message: "1 forma(s) fuera de la diapositiva seleccionada(s)." });
    });
  });

  describe("check-fonts", () => {
    it("lists the distinct font names used across the deck", async () => {
      const a = new MockShape({
        id: "a",
        type: "TextBox",
        left: 0,
        top: 0,
        width: 10,
        height: 10,
        font: { name: "Georgia", size: 18, color: "#000", bold: false, italic: false, underline: "None" },
      });
      const b = new MockShape({
        id: "b",
        type: "TextBox",
        left: 0,
        top: 0,
        width: 10,
        height: 10,
        font: { name: "Calibri", size: 18, color: "#000", bold: false, italic: false, underline: "None" },
      });
      const c = new MockShape({
        id: "c",
        type: "TextBox",
        left: 0,
        top: 0,
        width: 10,
        height: 10,
        font: { name: "Calibri", size: 18, color: "#000", bold: false, italic: false, underline: "None" },
      });
      const img = new MockShape({ id: "img", type: "Image", left: 0, top: 0, width: 10, height: 10, font: null });
      const ctx = new MockContext([], "", {
        slides: [
          { id: "sl1", shapes: [a, b] },
          { id: "sl2", shapes: [c, img] },
        ],
      });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-fonts")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 0, message: "Fuentes: Calibri, Georgia" });
      expect(ctx.invalidTouched).toBe(false); // nunca tocó textFrame de la imagen
    });

    it("truncates to 12 font names, appending '…'", async () => {
      const shapes = Array.from({ length: 13 }, (_, i) =>
        new MockShape({
          id: `s${i}`,
          type: "TextBox",
          left: 0,
          top: 0,
          width: 10,
          height: 10,
          font: { name: `Font${i}`, size: 18, color: "#000", bold: false, italic: false, underline: "None" },
        }),
      );
      const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes }] });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-fonts")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result.ok).toBe(true);
      const message = (result as { ok: true; message?: string }).message!;
      expect(message.endsWith(", …")).toBe(true);
      expect(message.split(", ")).toHaveLength(13); // 12 nombres + "…"
    });

    it("reports 'no hay texto en la presentación' when no shape has a font name", async () => {
      const img = new MockShape({ id: "img1", type: "Image", left: 0, top: 0, width: 10, height: 10, font: null });
      const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [img] }] });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-fonts")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 0, message: "No hay texto en la presentación." });
    });

    it("does not touch an Image placeholder (no font to read)", async () => {
      const imagePlaceholder = new MockShape({
        id: "ph-img",
        type: "Placeholder",
        containedType: "Image",
        left: 0,
        top: 0,
        width: 10,
        height: 10,
        font: null, // si se tocara su textFrame, el mock marcaría invalidTouched
      });
      const textBox = new MockShape({
        id: "tb1",
        type: "TextBox",
        left: 0,
        top: 0,
        width: 10,
        height: 10,
        font: { name: "Georgia", size: 18, color: "#000", bold: false, italic: false, underline: "None" },
      });
      const ctx = new MockContext([], "", { slides: [{ id: "sl1", shapes: [imagePlaceholder, textBox] }] });
      stubPowerPoint(ctx);

      const cmd = findCommand("check-fonts")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 0, message: "Fuentes: Georgia" });
      expect(imagePlaceholder.loads).toEqual([]);
      expect(ctx.invalidTouched).toBe(false);
    });
  });
});
