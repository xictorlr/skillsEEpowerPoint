import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { executeCommand, resetSlideSizeCache } from "../../src/commands/execute";
import { commandsBySection, findCommand } from "../../src/commands/registry";
import type { CommandOptions } from "../../src/commands/types";
import { MockContext, MockShape, type MockSlide } from "../office/mockPowerPoint";

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

/** Devuelve las formas `TextBox` de `slide`, en orden de inserción (título primero, luego cada ítem). */
function shapesOf(slide: MockSlide): MockShape[] {
  return slide.shapesList;
}

describe("wizards commands", () => {
  beforeEach(() => {
    resetSlideSizeCache();
    vi.spyOn(console, "error").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("is hidden from the section's button list", () => {
    expect(commandsBySection("wizards").map((c) => c.id)).not.toContain("agenda-create");
    expect(findCommand("agenda-create")).toBeDefined();
  });

  it("with an empty sectionsText, reports the validation message without touching the deck", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "sl1" }] });
    stubPowerPoint(ctx);

    const cmd = findCommand("agenda-create")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, agenda: { title: "Agenda", sectionsText: "   \n  " } });

    expect(result).toEqual({ ok: true, changed: 0, message: "Escribe al menos una sección (una por línea)." });
    expect(ctx.slidesList).toHaveLength(1);
    expect(ctx.calls).toEqual([]);
  });

  it("creates one slide per section, in order right after the active slide, each with a title and the full section list (the current one highlighted)", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "sl1" }] });
    stubPowerPoint(ctx);

    const cmd = findCommand("agenda-create")!;
    const result = await executeCommand(cmd, {
      ...DEFAULT_OPTIONS,
      agenda: { title: "Plan", sectionsText: "Intro\nCore\nCierre" },
    });

    expect(result).toEqual({ ok: true, changed: 3, message: "3 diapositiva(s) de agenda creada(s)." });

    // 3 nuevas diapositivas, insertadas en orden justo tras la actual (index 0): moveTo 1, 2, 3.
    expect(ctx.slidesList.map((s) => s.id)).toEqual(["sl1", "new-slide-1", "new-slide-2", "new-slide-3"]);
    const moveToCalls = ctx.calls.filter((c) => c.path === "moveTo");
    expect(moveToCalls).toEqual([
      { id: "new-slide-1", path: "moveTo", value: 1 },
      { id: "new-slide-2", path: "moveTo", value: 2 },
      { id: "new-slide-3", path: "moveTo", value: 3 },
    ]);

    const slide1 = ctx.slidesList.find((s) => s.id === "new-slide-1")!;
    const slide2 = ctx.slidesList.find((s) => s.id === "new-slide-2")!;
    const slide3 = ctx.slidesList.find((s) => s.id === "new-slide-3")!;

    // Cada diapositiva de agenda: 1 título + 3 ítems (una fila por sección) = 4 cuadros de texto.
    expect(shapesOf(slide1)).toHaveLength(4);
    expect(shapesOf(slide2)).toHaveLength(4);
    expect(shapesOf(slide3)).toHaveLength(4);

    // El texto de cada cuadro se verifica vía ctx.calls (el mock no persiste el texto en la forma).
    const addTextBoxCalls = ctx.calls.filter((c) => c.path === "addTextBox");
    expect(addTextBoxCalls).toHaveLength(12); // 3 diapositivas x 4 cuadros
    const titleTexts = [addTextBoxCalls[0]!, addTextBoxCalls[4]!, addTextBoxCalls[8]!].map((c) => (c.value as { text: string }).text);
    expect(titleTexts).toEqual(["Plan", "Plan", "Plan"]);
    const slide1ItemTexts = addTextBoxCalls.slice(1, 4).map((c) => (c.value as { text: string }).text);
    expect(slide1ItemTexts).toEqual(["Intro", "Core", "Cierre"]);

    // En la diapositiva de la sección 0 ("Intro"), el ítem resaltado es el primero (bold + acento);
    // los demás llevan el color gris por defecto y bold:false.
    const [title1, item1a, item1b, item1c] = shapesOf(slide1);
    expect(title1!.formatWrites).toContainEqual({ path: "font.size", value: 28 });
    expect(title1!.formatWrites).toContainEqual({ path: "font.bold", value: true });
    expect(title1!.formatWrites.map((w) => w.path)).not.toContain("font.color");

    expect(item1a!.formatWrites).toContainEqual({ path: "font.bold", value: true });
    expect(item1a!.formatWrites).toContainEqual({ path: "font.color", value: "#1F6FEB" }); // sin tema: acento por defecto
    expect(item1b!.formatWrites).toContainEqual({ path: "font.bold", value: false });
    expect(item1b!.formatWrites).toContainEqual({ path: "font.color", value: "#7F7F7F" });
    expect(item1c!.formatWrites).toContainEqual({ path: "font.bold", value: false });
    expect(item1c!.formatWrites).toContainEqual({ path: "font.color", value: "#7F7F7F" });

    // En la diapositiva de la sección 1 ("Core"), el segundo ítem es el resaltado.
    const [, item2a, item2b, item2c] = shapesOf(slide2);
    expect(item2a!.formatWrites).toContainEqual({ path: "font.bold", value: false });
    expect(item2b!.formatWrites).toContainEqual({ path: "font.bold", value: true });
    expect(item2b!.formatWrites).toContainEqual({ path: "font.color", value: "#1F6FEB" });
    expect(item2c!.formatWrites).toContainEqual({ path: "font.bold", value: false });
  });

  it("uses the theme's Accent1 color (API 1.10) instead of the default accent when available", async () => {
    stubOffice1_10();
    const ctx = new MockContext([], "", { slides: [{ id: "sl1" }], themeColors: { Accent1: "#00AA00" } });
    stubPowerPoint(ctx);

    const cmd = findCommand("agenda-create")!;
    await executeCommand(cmd, { ...DEFAULT_OPTIONS, agenda: { title: "Plan", sectionsText: "Solo" } });

    const slide1 = ctx.slidesList.find((s) => s.id === "new-slide-1")!;
    const [, item1] = shapesOf(slide1);
    expect(item1!.formatWrites).toContainEqual({ path: "font.color", value: "#00AA00" });
  });

  it("passes the blank layout's slide master id and layout id (found via findBlankLayout) to each slides.add call", async () => {
    // Mock por defecto: masterId "master-1", layouts = [{ id: "lay-title", ... }, { id: "lay-blank", name: "Blank" }].
    const ctx = new MockContext([], "", { slides: [{ id: "sl1" }] });
    stubPowerPoint(ctx);

    const cmd = findCommand("agenda-create")!;
    await executeCommand(cmd, { ...DEFAULT_OPTIONS, agenda: { title: "Plan", sectionsText: "Uno\nDos" } });

    const addCalls = ctx.calls.filter((c) => c.path === "slides.add");
    expect(addCalls).toHaveLength(2);
    expect(addCalls.every((c) => c.value && typeof c.value === "object")).toBe(true);
    for (const call of addCalls) {
      expect(call.value).toEqual({ slideMasterId: "master-1", layoutId: "lay-blank" });
    }
  });

  it("without a blank layout (layouts: []), deletes the placeholders of every newly created slide", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "sl1" }], layouts: [], newSlidePlaceholders: 2 });
    stubPowerPoint(ctx);

    const cmd = findCommand("agenda-create")!;
    await executeCommand(cmd, { ...DEFAULT_OPTIONS, agenda: { title: "Plan", sectionsText: "Uno" } });

    // slides.add no debe llevar layoutId (ninguno encontrado).
    const addCalls = ctx.calls.filter((c) => c.path === "slides.add");
    expect(addCalls).toEqual([{ id: "presentation", path: "slides.add", value: undefined }]);

    const slide1 = ctx.slidesList.find((s) => s.id === "new-slide-1")!;
    const placeholders = shapesOf(slide1).filter((s) => s.type === "Placeholder");
    expect(placeholders).toHaveLength(2);
    for (const ph of placeholders) {
      expect(ph.formatWrites).toContainEqual({ path: "delete", value: null });
    }
    // La diapositiva conserva, además, los cuadros propios de la agenda (título + 1 ítem).
    expect(shapesOf(slide1).filter((s) => s.type === "TextBox")).toHaveLength(2);
  });

  it("with the active slide at index 1 of 3, inserts sections sequentially right after it (moveTo 2, 3, 4)", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "sl1" }, { id: "sl2" }, { id: "sl3" }], activeSlideIndex: 1 });
    stubPowerPoint(ctx);

    const cmd = findCommand("agenda-create")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, agenda: { title: "Plan", sectionsText: "Uno\nDos\nTres" } });

    expect(result).toEqual({ ok: true, changed: 3, message: "3 diapositiva(s) de agenda creada(s)." });
    const moveToCalls = ctx.calls.filter((c) => c.path === "moveTo");
    expect(moveToCalls.map((c) => c.value)).toEqual([2, 3, 4]);
  });
});
