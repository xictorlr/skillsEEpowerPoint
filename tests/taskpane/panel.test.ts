// @vitest-environment happy-dom
/// <reference lib="dom" />
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { resetClipboard, setClipboard } from "../../src/commands/clipboard";
import { executeCommand } from "../../src/commands/execute";
import { commandsBySection, findCommand, hiddenCommandIds, SECTION_ORDER } from "../../src/commands/registry";
import { supportsApi } from "../../src/office/capabilities";
import type { CommandResult } from "../../src/office/runtime";
import type { SelectionSummary } from "../../src/office/summary";
import { getThemeColors, type ThemeColor } from "../../src/office/theme";
import { createPanel, deferred, flushMicrotasks, getCommandButton } from "./domHelpers";

vi.mock("../../src/commands/execute", () => ({ executeCommand: vi.fn(), resetSlideSizeCache: vi.fn() }));
vi.mock("../../src/office/theme", () => ({ getThemeColors: vi.fn() }));
vi.mock("../../src/office/capabilities", () => ({ supportsApi: vi.fn() }));

const executeCommandMock = vi.mocked(executeCommand);
const getThemeColorsMock = vi.mocked(getThemeColors);
const supportsApiMock = vi.mocked(supportsApi);

const OK_RESULT: CommandResult = { ok: true, changed: 1 };

/** Resumen de selección mínimo con los campos que no importan a cada test. */
function summary(partial: Partial<SelectionSummary>): SelectionSummary {
  return { count: 0, slideId: null, slideIndex: 0, slideCount: 0, bbox: null, ...partial };
}

beforeEach(() => {
  localStorage.clear();
  resetClipboard();
  executeCommandMock.mockReset();
  executeCommandMock.mockResolvedValue(OK_RESULT);
  getThemeColorsMock.mockReset();
  getThemeColorsMock.mockResolvedValue(null);
  supportsApiMock.mockReset();
  supportsApiMock.mockReturnValue(false);
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

describe("render()", () => {
  it("creates one section per SECTION_ORDER, with a tagged button per visible command", () => {
    const { root } = createPanel();

    // SECTION_ORDER.length secciones de comandos + la sección "Settings" que añade el panel aparte
    // (ver `render()` en panel.ts, que también usa la clase "st-section").
    expect(root.querySelectorAll(".st-section").length).toBe(SECTION_ORDER.length + 1);

    for (const section of SECTION_ORDER) {
      for (const cmd of commandsBySection(section)) {
        const btn = getCommandButton(root, cmd.id);
        expect(btn.tagName).toBe("BUTTON");
        expect(btn.getAttribute("aria-label")).toBe(cmd.label);
      }
    }

    // Los comandos "ocultos" (set-color, select-kind, ...) no tienen un botón genérico: sus
    // secciones los disparan desde controles propios (ver colorSection.ts/expertSection.ts/...).
    for (const hiddenId of hiddenCommandIds) {
      expect(root.querySelector(`[data-command-id="${hiddenId}"]`)).toBeNull();
    }
  });
});

describe("setSelectionCount", () => {
  it("disables commands whose minShapes isn't met and enables those with minShapes 0; null enables everything; the title carries the reason", () => {
    const { panel, root } = createPanel();
    const alignLeftCmd = findCommand("align-left")!; // minShapes: 2
    const checkFontsCmd = findCommand("check-fonts")!; // minShapes: 0
    const alignLeft = getCommandButton(root, alignLeftCmd.id);
    const checkFonts = getCommandButton(root, checkFontsCmd.id);

    panel.setSelectionCount(0);
    expect(alignLeft.disabled).toBe(true);
    expect(alignLeft.title).toBe(`${alignLeftCmd.label} — Selecciona al menos 2 formas.`);
    expect(checkFonts.disabled).toBe(false);
    expect(checkFonts.title).toBe(checkFontsCmd.label);

    panel.setSelectionCount(null);
    expect(alignLeft.disabled).toBe(false);
    expect(alignLeft.title).toBe(alignLeftCmd.label);
    expect(checkFonts.disabled).toBe(false);
  });
});

describe("clipboard gating", () => {
  it("disables paste-* without a copied format, with the matching reason; enables after setClipboard + setSelectionCount(1)", () => {
    const { panel, root } = createPanel();
    const pasteFormatCmd = findCommand("paste-format")!;
    const pasteFormat = getCommandButton(root, pasteFormatCmd.id);

    panel.setSelectionCount(1);
    expect(pasteFormat.disabled).toBe(true);
    expect(pasteFormat.title).toBe(`${pasteFormatCmd.label} — Primero copia el formato de una forma.`);

    setClipboard({ id: "clip", fill: null, line: null, font: null, size: { width: 10, height: 10 } });
    panel.setSelectionCount(1);
    expect(pasteFormat.disabled).toBe(false);
    expect(pasteFormat.title).toBe(pasteFormatCmd.label);
  });
});

describe("color target 'font' blocks transparency", () => {
  it("disables transparency-up/down with the matching reason when colorTarget is 'font'", () => {
    localStorage.setItem("slide-tools.options", JSON.stringify({ colorTarget: "font" }));
    const { panel, root } = createPanel();
    const upCmd = findCommand("transparency-up")!;
    const downCmd = findCommand("transparency-down")!;

    panel.setSelectionCount(1);

    const up = getCommandButton(root, upCmd.id);
    const down = getCommandButton(root, downCmd.id);
    expect(up.disabled).toBe(true);
    expect(up.title).toBe(`${upCmd.label} — La transparencia no aplica al texto.`);
    expect(down.disabled).toBe(true);
    expect(down.title).toBe(`${downCmd.label} — La transparencia no aplica al texto.`);
  });
});

describe("busy state", () => {
  it("a click executes; a second click during execution doesn't call executeCommand; aria-busy toggles; requestRefresh runs when done", async () => {
    const requestRefresh = vi.fn();
    const { root } = createPanel(requestRefresh);
    const pending = deferred<CommandResult>();
    executeCommandMock.mockReturnValueOnce(pending.promise);

    const btn = getCommandButton(root, "check-fonts");
    btn.click();

    expect(executeCommandMock).toHaveBeenCalledTimes(1);
    expect(root.getAttribute("aria-busy")).toBe("true");
    expect(btn.disabled).toBe(true);

    btn.click(); // segundo clic mientras está ocupado: no debe volver a llamar a executeCommand
    expect(executeCommandMock).toHaveBeenCalledTimes(1);

    pending.resolve({ ok: true, changed: 0, message: "Nada que cambiar." });
    await flushMicrotasks();

    expect(root.hasAttribute("aria-busy")).toBe(false);
    expect(requestRefresh).toHaveBeenCalledTimes(1);
  });
});

describe("confirm arming (two-step)", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  it("arms on the first click without executing, executes on the second, another command disarms it, and the 5s timer disarms it too", async () => {
    const { root } = createPanel();
    const deckCmd = findCommand("check-empty-text-deck")!;
    const deck = getCommandButton(root, deckCmd.id);
    const fonts = getCommandButton(root, "check-fonts");

    deck.click();
    expect(executeCommandMock).not.toHaveBeenCalled();
    expect(deck.classList.contains("is-armed")).toBe(true);
    expect(deck.title).toBe("Pulsa de nuevo para confirmar");

    deck.click(); // segundo clic sobre el mismo comando: ejecuta
    expect(executeCommandMock).toHaveBeenCalledTimes(1);
    expect(executeCommandMock.mock.calls[0]?.[0]?.id).toBe(deckCmd.id);
    expect(deck.classList.contains("is-armed")).toBe(false);
    await flushMicrotasks(); // deja que termine el comando (busy=false) antes de rearmar

    deck.click(); // arma de nuevo
    expect(deck.classList.contains("is-armed")).toBe(true);

    fonts.click(); // comando distinto sin confirmación: desarma el anterior y ejecuta el suyo
    expect(deck.classList.contains("is-armed")).toBe(false);
    expect(executeCommandMock).toHaveBeenCalledTimes(2);
    expect(executeCommandMock.mock.calls[1]?.[0]?.id).toBe("check-fonts");
    await flushMicrotasks();

    deck.click(); // arma por tercera vez, para probar la expiración por temporizador
    expect(deck.classList.contains("is-armed")).toBe(true);

    vi.advanceTimersByTime(5000);
    expect(deck.classList.contains("is-armed")).toBe(false);
    expect(executeCommandMock).toHaveBeenCalledTimes(2); // la expiración no ejecuta nada
  });
});

describe("Expert: inspector", () => {
  it("fills the fields from setSummary({ bbox }); an emptied field restores the value without executing; a valid value executes inspector-apply with { inspector: { left: n } }", async () => {
    const { panel, root } = createPanel();
    panel.setSummary(summary({ count: 1, bbox: { left: 10, top: 20, width: 30, height: 40 } }));

    const inputs = root.querySelectorAll<HTMLInputElement>('input[step="0.1"]');
    expect(inputs.length).toBe(4); // X, Y, Ancho, Alto (ver INSPECTOR_FIELDS en expertSection.ts)
    const [left, top, width, height] = Array.from(inputs);
    expect(left!.disabled).toBe(false);
    expect(left!.value).toBe("10.0");
    expect(top!.value).toBe("20.0");
    expect(width!.value).toBe("30.0");
    expect(height!.value).toBe("40.0");

    // Campo vacío: restaura el último valor conocido y no ejecuta nada.
    left!.value = "";
    left!.dispatchEvent(new Event("change", { bubbles: true }));
    expect(left!.value).toBe("10.0");
    expect(executeCommandMock).not.toHaveBeenCalled();

    // Valor válido: ejecuta inspector-apply con { inspector: { left: 15 } }.
    left!.value = "15";
    left!.dispatchEvent(new Event("change", { bubbles: true }));
    expect(executeCommandMock).toHaveBeenCalledTimes(1);
    const [calledCmd, calledOptions] = executeCommandMock.mock.calls[0]!;
    expect(calledCmd.id).toBe("inspector-apply");
    expect(calledOptions.inspector).toEqual({ left: 15 });
    await flushMicrotasks();
  });
});

describe("theme colors", () => {
  function stubPowerPoint(): void {
    vi.stubGlobal("PowerPoint", { run: (fn: (context: unknown) => unknown) => fn({}) });
  }

  it("a new slideId in setSummary calls getThemeColors once; two rapid changes apply only the last (token)", async () => {
    supportsApiMock.mockReturnValue(true);
    stubPowerPoint();
    const { panel, root } = createPanel();
    const themeRow = (): HTMLElement => root.querySelector('.st-row[aria-label="Colores del tema"]') as HTMLElement;

    const first = deferred<ThemeColor[] | null>();
    const second = deferred<ThemeColor[] | null>();
    getThemeColorsMock.mockReturnValueOnce(first.promise).mockReturnValueOnce(second.promise);

    panel.setSummary(summary({ count: 1, slideId: "slide-a" }));
    expect(getThemeColorsMock).toHaveBeenCalledTimes(1);

    // Mismo slideId: no vuelve a leer los colores del tema.
    panel.setSummary(summary({ count: 1, slideId: "slide-a" }));
    expect(getThemeColorsMock).toHaveBeenCalledTimes(1);

    // Cambia de diapositiva de nuevo antes de que la primera lectura resuelva.
    panel.setSummary(summary({ count: 1, slideId: "slide-b" }));
    expect(getThemeColorsMock).toHaveBeenCalledTimes(2);

    // Resuelve primero la lectura más reciente...
    second.resolve([{ name: "Accent1", hex: "#222222" }]);
    await flushMicrotasks();
    expect(themeRow().children.length).toBe(1);
    expect((themeRow().children[0] as HTMLElement).title).toContain("#222222");

    // ...y luego la respuesta tardía de la primera: se descarta (el token ya avanzó).
    first.resolve([{ name: "Dark1", hex: "#111111" }]);
    await flushMicrotasks();
    expect(themeRow().children.length).toBe(1);
    expect((themeRow().children[0] as HTMLElement).title).toContain("#222222");
  });

  it("does not call getThemeColors when supportsApi returns false", async () => {
    supportsApiMock.mockReturnValue(false);
    const { panel } = createPanel();

    panel.setSummary(summary({ count: 1, slideId: "slide-a" }));
    await flushMicrotasks();

    expect(getThemeColorsMock).not.toHaveBeenCalled();
  });
});

describe("recent colors", () => {
  it("a successful applyColor with changed > 0 adds to recents and keeps focus within the row", async () => {
    localStorage.setItem("slide-tools.options", JSON.stringify({ recentColors: ["#AABBCC"] }));
    executeCommandMock.mockResolvedValue({ ok: true, changed: 1 });

    const { root } = createPanel();
    const recentSubtitle = Array.from(root.querySelectorAll<HTMLElement>(".st-subtitle")).find(
      (el) => el.textContent === "Recientes",
    )!;
    const recentRow = recentSubtitle.nextElementSibling as HTMLElement;
    expect(recentRow.children.length).toBe(1);

    const firstSwatch = recentRow.firstElementChild as HTMLElement;
    expect(firstSwatch.getAttribute("aria-label")).toBe("Color #AABBCC");
    firstSwatch.focus();
    expect(recentRow.contains(document.activeElement)).toBe(true);

    firstSwatch.click();
    await flushMicrotasks();

    expect(executeCommandMock).toHaveBeenCalledWith(
      expect.objectContaining({ id: "set-color" }),
      expect.objectContaining({ color: "#AABBCC" }),
    );
    expect(recentRow.children.length).toBe(1); // ya era el único/primero reciente: no se duplica
    // El nodo del DOM se reemplaza (replaceChildren), pero el foco se restaura dentro de la fila.
    expect(document.activeElement).toBe(recentRow.firstElementChild);
  });
});
