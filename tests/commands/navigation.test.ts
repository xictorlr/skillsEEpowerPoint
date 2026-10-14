import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { executeCommand, resetSlideSizeCache } from "../../src/commands/execute";
import { findCommand } from "../../src/commands/registry";
import type { CommandOptions } from "../../src/commands/types";
import { MockContext, type MockSlideInit } from "../office/mockPowerPoint";

const DEFAULT_OPTIONS: CommandOptions = {
  matrix: { columns: 2, gapX: 8, gapY: 8 },
  colorTarget: "fill",
  pasteParts: { fill: true, line: true, font: true, size: true },
  fontSizeStep: 1,
  nameFilter: "",
  renamePrefix: "Forma",
  shapeKind: "all",
  agenda: { title: "Agenda", sectionsText: "" },
  insert: { count: 3, stickerText: "BORRADOR", stickerColor: "#FFE066" },
};

const THREE_SLIDES: MockSlideInit[] = [{ id: "sl1" }, { id: "sl2" }, { id: "sl3" }];

function stubPowerPoint(ctx: MockContext): void {
  vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => Promise<unknown>) => fn(ctx) });
}

describe("navigation commands", () => {
  beforeEach(() => {
    resetSlideSizeCache();
    vi.spyOn(console, "error").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("slide-next from index 1 of 3 selects the id at index 2 and reports 'Diapositiva 3 de 3'", async () => {
    const ctx = new MockContext([], "", { slides: THREE_SLIDES, activeSlideIndex: 1 });
    stubPowerPoint(ctx);

    const cmd = findCommand("slide-next")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 0, message: "Diapositiva 3 de 3" });
    expect(ctx.selectedSlideIds).toEqual(["sl3"]);
    // needsSelection: false (ver types.ts/execute.ts): sin el sync de readSelection, solo
    // readSlideNav + la escritura de selectSlideByIndex (antes eran 3, con el de readSelection).
    expect(ctx.syncCount).toBe(2);
  });

  it("slide-next on the last slide reports the edge message without changing the selection", async () => {
    const ctx = new MockContext([], "", { slides: THREE_SLIDES, activeSlideIndex: 2 });
    stubPowerPoint(ctx);

    const cmd = findCommand("slide-next")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 0, message: "Ya estás en la última diapositiva." });
    expect(ctx.selectedSlideIds).toBeUndefined();
  });

  it("slide-prev on the first slide reports the edge message", async () => {
    const ctx = new MockContext([], "", { slides: THREE_SLIDES, activeSlideIndex: 0 });
    stubPowerPoint(ctx);

    const cmd = findCommand("slide-prev")!;
    const result = await executeCommand(cmd, DEFAULT_OPTIONS);

    expect(result).toEqual({ ok: true, changed: 0, message: "Ya estás en la primera diapositiva." });
    expect(ctx.selectedSlideIds).toBeUndefined();
  });

  it("slide-first and slide-last jump straight to the extremes", async () => {
    const ctxFirst = new MockContext([], "", { slides: THREE_SLIDES, activeSlideIndex: 1 });
    stubPowerPoint(ctxFirst);
    const first = await executeCommand(findCommand("slide-first")!, DEFAULT_OPTIONS);
    expect(first).toEqual({ ok: true, changed: 0, message: "Diapositiva 1 de 3" });
    expect(ctxFirst.selectedSlideIds).toEqual(["sl1"]);

    const ctxLast = new MockContext([], "", { slides: THREE_SLIDES, activeSlideIndex: 1 });
    stubPowerPoint(ctxLast);
    const last = await executeCommand(findCommand("slide-last")!, DEFAULT_OPTIONS);
    expect(last).toEqual({ ok: true, changed: 0, message: "Diapositiva 3 de 3" });
    expect(ctxLast.selectedSlideIds).toEqual(["sl3"]);
  });

  it("slide-goto out of range reports the slide count without changing the selection", async () => {
    const ctx = new MockContext([], "", { slides: THREE_SLIDES, activeSlideIndex: 0 });
    stubPowerPoint(ctx);

    const cmd = findCommand("slide-goto")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, gotoSlide: 5 });

    expect(result).toEqual({ ok: true, changed: 0, message: "La presentación tiene 3 diapositivas." });
    expect(ctx.selectedSlideIds).toBeUndefined();
  });

  it("slide-goto with a non-integer gotoSlide (NaN) reports the slide count without selecting", async () => {
    const ctx = new MockContext([], "", { slides: THREE_SLIDES, activeSlideIndex: 0 });
    stubPowerPoint(ctx);

    const cmd = findCommand("slide-goto")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, gotoSlide: NaN });

    expect(result).toEqual({ ok: true, changed: 0, message: "La presentación tiene 3 diapositivas." });
    expect(ctx.selectedSlideIds).toBeUndefined();
  });

  it("slide-goto with a fractional gotoSlide (2.5) reports the slide count without selecting", async () => {
    const ctx = new MockContext([], "", { slides: THREE_SLIDES, activeSlideIndex: 0 });
    stubPowerPoint(ctx);

    const cmd = findCommand("slide-goto")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, gotoSlide: 2.5 });

    expect(result).toEqual({ ok: true, changed: 0, message: "La presentación tiene 3 diapositivas." });
    expect(ctx.selectedSlideIds).toBeUndefined();
  });

  it("slide-goto in range selects the requested slide", async () => {
    const ctx = new MockContext([], "", { slides: THREE_SLIDES, activeSlideIndex: 0 });
    stubPowerPoint(ctx);

    const cmd = findCommand("slide-goto")!;
    const result = await executeCommand(cmd, { ...DEFAULT_OPTIONS, gotoSlide: 3 });

    expect(result).toEqual({ ok: true, changed: 0, message: "Diapositiva 3 de 3" });
    expect(ctx.selectedSlideIds).toEqual(["sl3"]);
  });
});
