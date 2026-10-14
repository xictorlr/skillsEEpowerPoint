import { executeCommand } from "../commands/execute";
import { SECTION_ORDER, commandsBySection, disabledReason } from "../commands/registry";
import { SECTION_LABELS, type Command, type CommandOptions } from "../commands/types";
import { iconSvg } from "./icons";
import { loadOptions, saveOptions } from "./settings";

export class Panel {
  private options: CommandOptions = loadOptions();
  private buttons = new Map<string, { el: HTMLButtonElement; cmd: Command }>();
  private noticeTimer: number | undefined;

  constructor(private root: HTMLElement, private notice: HTMLElement) {}

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
      if (section === "alignment") sec.appendChild(this.matrixFields());
      this.root.appendChild(sec);
    });
    this.setSelectionCount(0);
  }

  /** Habilita/deshabilita botones según el número de formas seleccionadas. */
  setSelectionCount(count: number): void {
    for (const { el, cmd } of this.buttons.values()) {
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
    this.noticeTimer = window.setTimeout(() => (this.notice.hidden = true), 3000);
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
    const result = await executeCommand(cmd, this.options);
    if (!result.ok) this.showNotice(result.message, true);
    else if (result.changed === 0) this.showNotice("Nada que cambiar.");
  }

  private matrixFields(): HTMLElement {
    const wrap = document.createElement("div");
    wrap.className = "st-row";
    const make = (label: string, key: "columns" | "gapX" | "gapY", min: number) => {
      const field = document.createElement("label");
      field.className = "st-field";
      field.append(label);
      const input = document.createElement("input");
      input.type = "number";
      input.min = String(min);
      input.value = String(this.options.matrix[key]);
      input.addEventListener("change", () => {
        const value = Math.max(min, Number(input.value) || min);
        input.value = String(value);
        this.options = { ...this.options, matrix: { ...this.options.matrix, [key]: value } };
        saveOptions(this.options);
      });
      field.appendChild(input);
      return field;
    };
    wrap.append(make("Cols", "columns", 1), make("Gap X", "gapX", 0), make("Gap Y", "gapY", 0));
    return wrap;
  }
}
