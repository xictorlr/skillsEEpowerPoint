import { executeCommand } from "../commands/execute";
import { SECTION_ORDER, commandsBySection, disabledReason } from "../commands/registry";
import { SECTION_LABELS, type Command, type CommandOptions } from "../commands/types";
import { iconSvg } from "./icons";
import {
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

export class Panel {
  private persisted: PersistedOptions = loadPersistedOptions();
  private options: CommandOptions = toCommandOptions(this.persisted);
  private buttons = new Map<string, { el: HTMLButtonElement; cmd: Command }>();
  private noticeTimer: number | undefined;
  private lastCount: number | null = null;
  private busy = false;

  constructor(private root: HTMLElement, private notice: HTMLElement) {
    notice.setAttribute("role", "status");
    notice.setAttribute("aria-live", "polite");
  }

  render(): void {
    this.root.replaceChildren();
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
    for (const { el, cmd } of this.buttons.values()) {
      if (count === null) {
        el.disabled = false;
        el.title = cmd.label;
        continue;
      }
      const reason = disabledReason(cmd, count);
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

  private async run(cmd: Command): Promise<void> {
    if (this.busy) return;
    this.busy = true;
    for (const { el } of this.buttons.values()) el.disabled = true;
    this.root.setAttribute("aria-busy", "true");
    try {
      const result = await executeCommand(cmd, this.options);
      if (!result.ok) this.showNotice(result.message, true);
      else if (result.changed === 0) this.showNotice("Nada que cambiar.");
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
        const candidate = { ...this.persisted.matrix, [key]: Number(input.value) };
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
