import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { executeCommand, resetSlideSizeCache } from "../../src/commands/execute";
import { findCommand } from "../../src/commands/registry";
import type { CommandOptions } from "../../src/commands/types";
import { MockContext } from "../office/mockPowerPoint";

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

function stubPowerPoint(ctx: MockContext, createPresentationMock = vi.fn().mockResolvedValue(undefined)): typeof createPresentationMock {
  vi.stubGlobal("PowerPoint", {
    run: (fn: (context: unknown) => Promise<unknown>) => fn(ctx),
    createPresentation: createPresentationMock,
  });
  return createPresentationMock;
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

describe("presentation commands", () => {
  beforeEach(() => {
    resetSlideSizeCache();
    vi.spyOn(console, "error").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  describe("new-presentation", () => {
    it("creates a blank presentation and reports it", async () => {
      const ctx = new MockContext([], "", { slides: [{ id: "sl1" }] });
      const createPresentationMock = stubPowerPoint(ctx);

      const cmd = findCommand("new-presentation")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 0, message: "Nueva presentación creada." });
      expect(createPresentationMock).toHaveBeenCalledWith(undefined);
    });
  });

  describe("export-selected-slides", () => {
    it("without PowerPointApi 1.10, reports that a recent PowerPoint is required, without syncing at all", async () => {
      const ctx = new MockContext([], "", { slides: [{ id: "sl1" }] });
      const createPresentationMock = stubPowerPoint(ctx);

      const cmd = findCommand("export-selected-slides")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 0, message: "Exportar requiere PowerPoint reciente (API 1.10)." });
      expect(createPresentationMock).not.toHaveBeenCalled();
      // El único sync es el de la lectura común de la selección (execute.ts, previo a cmd.run); el
      // comando en sí no añade ninguno: la comprobación de la API va antes de leer la selección de
      // diapositivas.
      expect(ctx.syncCount).toBe(1);
    });

    it("with PowerPointApi 1.10 but no slide selected, reports a validation message", async () => {
      stubOffice1_10();
      // activeSlideIndex fuera de rango: simula que getSelectedSlides() no devuelve nada.
      const ctx = new MockContext([], "", { slides: [{ id: "sl1" }], activeSlideIndex: 5 });
      const createPresentationMock = stubPowerPoint(ctx);

      const cmd = findCommand("export-selected-slides")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 0, message: "Selecciona al menos una diapositiva." });
      expect(createPresentationMock).not.toHaveBeenCalled();
    });

    it("with PowerPointApi 1.10, exports the selected slide(s) into a new presentation", async () => {
      stubOffice1_10();
      const ctx = new MockContext([], "", { slides: [{ id: "sl1" }, { id: "sl2" }], activeSlideIndex: 1 });
      const createPresentationMock = stubPowerPoint(ctx);

      const cmd = findCommand("export-selected-slides")!;
      const result = await executeCommand(cmd, DEFAULT_OPTIONS);

      expect(result).toEqual({ ok: true, changed: 1, message: "1 diapositiva(s) exportada(s) a una nueva presentación." });
      expect(createPresentationMock).toHaveBeenCalledWith("BASE64:sl2");
    });
  });
});
