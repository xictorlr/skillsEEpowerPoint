import { getClipboard } from "../commands/clipboard";
import { executeCommand } from "../commands/execute";
import { SECTION_ORDER, commandsBySection, disabledReason, findCommand } from "../commands/registry";
import { SECTION_LABELS, type Command, type CommandOptions } from "../commands/types";
import { DEFAULT_PALETTE, normalizeHex, pushRecent } from "../core/color";
import type { ColorTarget, FormatParts } from "../core/format";
import type { CommandResult } from "../office/runtime";
import { iconSvg } from "./icons";
import {
  MAX_RECENT_COLORS,
  SLIDE_SIZE_PRESETS,
  sanitizeMatrix,
  sanitizeSlideSizePreset,
  toCommandOptions,
  type SlideSizePreset,
} from "./sanitize";
import { loadPersistedOptions, saveOptions, type PersistedOptions } from "./settings";

const SLIDE_SIZE_LABELS: Record<SlideSizePreset, string> = {
  auto: "Auto",
  "16x9": "16:9 (960×540)",
  "16x9-legacy": "16:9 clásico (720×405)",
  "4x3": "4:3 (720×540)",
  a4: "A4 (780×540)",
  letter: "Carta (720×540)",
};

const COLOR_TARGETS: ReadonlyArray<{ id: ColorTarget; icon: string; label: string }> = [
  { id: "fill", icon: "targetFill", label: "Destino: relleno" },
  { id: "line", icon: "targetLine", label: "Destino: línea" },
  { id: "font", icon: "targetFont", label: "Destino: texto" },
];

const PASTE_PART_FIELDS: ReadonlyArray<{ key: keyof FormatParts; label: string }> = [
  { key: "fill", label: "Relleno" },
  { key: "line", label: "Línea" },
  { key: "font", label: "Fuente" },
  { key: "size", label: "Tamaño" },
];

/** La transparencia solo tiene sentido para relleno/línea; con destino "texto" no aplica. */
const TRANSPARENCY_COMMAND_IDS: ReadonlySet<string> = new Set(["transparency-up", "transparency-down"]);
const TRANSPARENCY_FONT_REASON = "La transparencia no aplica al texto.";

export class Panel {
  private persisted: PersistedOptions = loadPersistedOptions();
  private options: CommandOptions = toCommandOptions(this.persisted);
  private buttons = new Map<string, { el: HTMLButtonElement; cmd: Command }>();
  private targetButtons = new Map<ColorTarget, HTMLButtonElement>();
  private recentRowEl: HTMLElement | undefined;
  private noticeTimer: number | undefined;
  private lastCount: number | null = null;
  private busy = false;

  constructor(private root: HTMLElement, private notice: HTMLElement) {
    notice.setAttribute("role", "status");
    notice.setAttribute("aria-live", "polite");
  }

  render(): void {
    this.root.replaceChildren();
    this.targetButtons.clear();
    SECTION_ORDER.forEach((section, i) => {
      if (i > 0) this.root.appendChild(Object.assign(document.createElement("hr"), { className: "st-divider" }));
      const sec = document.createElement("section");
      sec.className = "st-section";
      const title = document.createElement("div");
      title.className = "st-section-title";
      title.textContent = SECTION_LABELS[section];
      sec.appendChild(title);
      const row = document.createElement("div");
      row.className = "st-row";
      for (const cmd of commandsBySection(section)) row.appendChild(this.button(cmd));
      sec.appendChild(row);
      if (section === "format") sec.appendChild(this.pastePartsRow());
      if (section === "color") {
        sec.appendChild(this.colorTargetRow());
        sec.appendChild(this.paletteRow());
        sec.appendChild(this.recentSection());
        sec.appendChild(this.customColorField());
      }
      this.root.appendChild(sec);
    });
    this.root.appendChild(Object.assign(document.createElement("hr"), { className: "st-divider" }));
    this.root.appendChild(this.settingsSection());
    this.setSelectionCount(0);
  }

  /**
   * Habilita/deshabilita botones según el número de formas seleccionadas.
   * `null` significa que no se pudo determinar la selección: se habilita todo y
   * `executeCommand` revalida la selección real al ejecutar.
   * No toca los botones mientras hay un comando en curso (`busy`); el estado se
   * reaplica al terminar ese comando.
   */
  setSelectionCount(count: number | null): void {
    this.lastCount = count;
    if (this.busy) return;
    const hasClipboard = getClipboard() !== null;
    const transparencyBlocked = this.persisted.colorTarget === "font";
    for (const { el, cmd } of this.buttons.values()) {
      const overrideReason = transparencyBlocked && TRANSPARENCY_COMMAND_IDS.has(cmd.id) ? TRANSPARENCY_FONT_REASON : undefined;
      if (count === null) {
        el.disabled = overrideReason !== undefined;
        el.title = overrideReason ? `${cmd.label} — ${overrideReason}` : cmd.label;
        continue;
      }
      const reason = disabledReason(cmd, count, hasClipboard) ?? overrideReason;
      el.disabled = reason !== undefined;
      el.title = reason ? `${cmd.label} — ${reason}` : cmd.label;
    }
  }

  showNotice(text: string, isError = false): void {
    this.notice.textContent = text;
    this.notice.classList.toggle("is-error", isError);
    this.notice.hidden = false;
    window.clearTimeout(this.noticeTimer);
    this.noticeTimer = window.setTimeout(() => (this.notice.hidden = true), isError ? 6000 : 3000);
  }

  private button(cmd: Command): HTMLButtonElement {
    const el = document.createElement("button");
    el.className = "st-btn";
    el.type = "button";
    el.innerHTML = iconSvg(cmd.icon);
    el.setAttribute("aria-label", cmd.label);
    el.addEventListener("click", () => void this.run(cmd));
    this.buttons.set(cmd.id, { el, cmd });
    return el;
  }

  /** Crea una muestra de color. El color ya pasó por `normalizeHex` antes de llegar aquí (paleta fija,
   *  recientes saneados, o el valor normalizado que devuelve `applyColor`), así que nunca se vuelca
   *  HTML arbitrario: solo se usa como valor de una custom property CSS. */
  private swatch(color: string): HTMLButtonElement {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "st-swatch";
    btn.setAttribute("aria-label", `Color ${color}`);
    btn.title = color;
    btn.style.setProperty("--swatch", color);
    btn.addEventListener("click", () => void this.applyColor(color));
    return btn;
  }

  private paletteRow(): HTMLElement {
    const wrap = document.createElement("div");
    const subtitle = document.createElement("div");
    subtitle.className = "st-subtitle";
    subtitle.textContent = "Paleta";
    wrap.appendChild(subtitle);
    const row = document.createElement("div");
    row.className = "st-row";
    for (const color of DEFAULT_PALETTE) row.appendChild(this.swatch(color));
    wrap.appendChild(row);
    return wrap;
  }

  private recentSection(): HTMLElement {
    const wrap = document.createElement("div");
    const subtitle = document.createElement("div");
    subtitle.className = "st-subtitle";
    subtitle.textContent = "Recientes";
    wrap.appendChild(subtitle);
    this.recentRowEl = document.createElement("div");
    this.recentRowEl.className = "st-row";
    this.renderRecentRow();
    wrap.appendChild(this.recentRowEl);
    return wrap;
  }

  private renderRecentRow(): void {
    this.recentRowEl?.replaceChildren(...this.persisted.recentColors.map((c) => this.swatch(c)));
  }

  private colorTargetRow(): HTMLElement {
    const row = document.createElement("div");
    row.className = "st-row";
    for (const t of COLOR_TARGETS) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "st-btn st-toggle";
      btn.innerHTML = iconSvg(t.icon);
      btn.setAttribute("aria-label", t.label);
      btn.title = t.label;
      const active = this.persisted.colorTarget === t.id;
      btn.setAttribute("aria-pressed", String(active));
      btn.classList.toggle("is-active", active);
      btn.addEventListener("click", () => this.setColorTarget(t.id));
      this.targetButtons.set(t.id, btn);
      row.appendChild(btn);
    }
    return row;
  }

  private setColorTarget(target: ColorTarget): void {
    this.persisted = { ...this.persisted, colorTarget: target };
    this.options = toCommandOptions(this.persisted);
    saveOptions(this.persisted);
    for (const [id, btn] of this.targetButtons) {
      const active = id === target;
      btn.setAttribute("aria-pressed", String(active));
      btn.classList.toggle("is-active", active);
    }
    this.setSelectionCount(this.lastCount);
  }

  private customColorField(): HTMLElement {
    const field = document.createElement("label");
    field.className = "st-field";
    field.append("Otro");
    const input = document.createElement("input");
    input.type = "color";
    input.addEventListener("change", () => void this.applyColor(input.value));
    field.appendChild(input);
    return field;
  }

  private async applyColor(color: string): Promise<void> {
    const hex = normalizeHex(color);
    if (!hex) return;
    const cmd = findCommand("set-color");
    if (!cmd) return;
    const result = await this.run(cmd, { color: hex });
    if (result?.ok) {
      this.persisted = { ...this.persisted, recentColors: pushRecent(this.persisted.recentColors, hex, MAX_RECENT_COLORS) };
      saveOptions(this.persisted);
      this.renderRecentRow();
    }
  }

  private pastePartsRow(): HTMLElement {
    const row = document.createElement("div");
    row.className = "st-row";
    for (const part of PASTE_PART_FIELDS) {
      const label = document.createElement("label");
      label.className = "st-check";
      const input = document.createElement("input");
      input.type = "checkbox";
      input.checked = this.persisted.pasteParts[part.key];
      input.addEventListener("change", () => {
        const pasteParts = { ...this.persisted.pasteParts, [part.key]: input.checked };
        this.persisted = { ...this.persisted, pasteParts };
        this.options = toCommandOptions(this.persisted);
        saveOptions(this.persisted);
      });
      label.append(input, ` ${part.label}`);
      row.appendChild(label);
    }
    return row;
  }

  private async run(cmd: Command, extra?: Partial<CommandOptions>): Promise<CommandResult | undefined> {
    if (this.busy) return undefined;
    this.busy = true;
    for (const { el } of this.buttons.values()) el.disabled = true;
    this.root.setAttribute("aria-busy", "true");
    try {
      const options = extra ? { ...this.options, ...extra } : this.options;
      const result = await executeCommand(cmd, options);
      if (!result.ok) this.showNotice(result.message, true);
      else if (result.message) this.showNotice(result.message);
      else if (result.changed === 0) this.showNotice("Nada que cambiar.");
      return result;
    } finally {
      this.busy = false;
      this.root.removeAttribute("aria-busy");
      this.setSelectionCount(this.lastCount);
    }
  }

  /** Sección "Ajustes": tamaño de diapositiva forzado y parámetros de la matriz. */
  private settingsSection(): HTMLElement {
    const sec = document.createElement("section");
    sec.className = "st-section";
    const title = document.createElement("div");
    title.className = "st-section-title";
    title.textContent = "Ajustes";
    sec.appendChild(title);
    const row = document.createElement("div");
    row.className = "st-row";
    row.appendChild(this.slideSizeField());
    sec.appendChild(row);
    sec.appendChild(this.matrixFields());
    return sec;
  }

  private slideSizeField(): HTMLElement {
    const field = document.createElement("label");
    field.className = "st-field";
    field.append("Diapositiva");
    const select = document.createElement("select");
    select.className = "st-select";
    for (const key of Object.keys(SLIDE_SIZE_PRESETS) as SlideSizePreset[]) {
      const option = document.createElement("option");
      option.value = key;
      option.textContent = SLIDE_SIZE_LABELS[key];
      select.appendChild(option);
    }
    select.value = this.persisted.slideSizePreset;
    select.addEventListener("change", () => {
      const slideSizePreset = sanitizeSlideSizePreset(select.value);
      select.value = slideSizePreset;
      this.persisted = { ...this.persisted, slideSizePreset };
      this.options = toCommandOptions(this.persisted);
      saveOptions(this.persisted);
    });
    field.appendChild(select);
    return field;
  }

  private matrixFields(): HTMLElement {
    const wrap = document.createElement("div");
    wrap.className = "st-row";
    const make = (label: string, key: "columns" | "gapX" | "gapY", min: number, step?: string) => {
      const field = document.createElement("label");
      field.className = "st-field";
      field.append(label);
      const input = document.createElement("input");
      input.type = "number";
      input.min = String(min);
      if (step) input.step = step;
      input.value = String(this.persisted.matrix[key]);
      input.addEventListener("change", () => {
        const raw = Number(input.value);
        const candidate = {
          ...this.persisted.matrix,
          [key]: Number.isFinite(raw) ? Math.max(min, raw) : this.persisted.matrix[key],
        };
        const matrix = sanitizeMatrix(candidate);
        input.value = String(matrix[key]);
        this.persisted = { ...this.persisted, matrix };
        this.options = toCommandOptions(this.persisted);
        saveOptions(this.persisted);
      });
      field.appendChild(input);
      return field;
    };
    wrap.append(make("Cols", "columns", 1, "1"), make("Gap X", "gapX", 0), make("Gap Y", "gapY", 0));
    return wrap;
  }
}
