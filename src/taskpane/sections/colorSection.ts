import { findCommand } from "../../commands/registry";
import { DEFAULT_PALETTE, normalizeHex, pushRecent } from "../../core/color";
import type { ColorTarget } from "../../core/format";
import { iconSvg } from "../icons";
import { MAX_RECENT_COLORS } from "../sanitize";
import type { SectionContext } from "./types";

const COLOR_TARGETS: ReadonlyArray<{ id: ColorTarget; icon: string; label: string }> = [
  { id: "fill", icon: "targetFill", label: "Destino: relleno" },
  { id: "line", icon: "targetLine", label: "Destino: línea" },
  { id: "font", icon: "targetFont", label: "Destino: texto" },
];

/** Sección "Color": destino, paleta, recientes y selector de color personalizado. */
export class ColorSection {
  private targetButtons = new Map<ColorTarget, HTMLButtonElement>();
  private recentRowEl: HTMLElement | undefined;

  constructor(private ctx: SectionContext) {}

  render(): HTMLElement[] {
    return [this.colorTargetRow(), this.paletteRow(), this.recentSection(), this.customColorField()];
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
    this.recentRowEl?.replaceChildren(...this.ctx.persisted.recentColors.map((c) => this.swatch(c)));
  }

  private colorTargetRow(): HTMLElement {
    const wrap = document.createElement("div");
    const subtitle = document.createElement("div");
    subtitle.className = "st-subtitle";
    subtitle.textContent = "Destino";
    wrap.appendChild(subtitle);
    const row = document.createElement("div");
    row.className = "st-row";
    row.setAttribute("role", "group");
    row.setAttribute("aria-label", "Destino del color");
    for (const t of COLOR_TARGETS) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "st-btn st-toggle";
      btn.innerHTML = iconSvg(t.icon);
      btn.setAttribute("aria-label", t.label);
      btn.title = t.label;
      const active = this.ctx.persisted.colorTarget === t.id;
      btn.setAttribute("aria-pressed", String(active));
      btn.classList.toggle("is-active", active);
      btn.addEventListener("click", () => this.setColorTarget(t.id));
      this.targetButtons.set(t.id, btn);
      row.appendChild(btn);
    }
    wrap.appendChild(row);
    return wrap;
  }

  private setColorTarget(target: ColorTarget): void {
    this.ctx.updatePersisted({ colorTarget: target });
    for (const [id, btn] of this.targetButtons) {
      const active = id === target;
      btn.setAttribute("aria-pressed", String(active));
      btn.classList.toggle("is-active", active);
    }
    this.ctx.refreshButtons();
  }

  private customColorField(): HTMLElement {
    const field = document.createElement("label");
    field.className = "st-field";
    field.append("Otro");
    const input = document.createElement("input");
    input.type = "color";
    input.value = this.ctx.persisted.recentColors[0] ?? "#1F6FEB";
    let debounce: number | undefined;
    input.addEventListener("change", () => {
      window.clearTimeout(debounce);
      const value = input.value;
      debounce = window.setTimeout(() => {
        void this.applyColor(value).then((hex) => {
          if (hex) input.value = hex;
        });
      }, 300);
    });
    field.appendChild(input);
    return field;
  }

  /** Aplica el color y, si tuvo efecto, lo añade a recientes. Devuelve el hex normalizado si se aplicó. */
  private async applyColor(color: string): Promise<string | undefined> {
    const hex = normalizeHex(color);
    if (!hex) return undefined;
    const cmd = findCommand("set-color");
    if (!cmd) return undefined;
    const result = await this.ctx.run(cmd, { color: hex });
    if (result?.ok && result.changed > 0) {
      this.ctx.updatePersisted({ recentColors: pushRecent(this.ctx.persisted.recentColors, hex, MAX_RECENT_COLORS) });
      const hadFocus = this.recentRowEl?.contains(document.activeElement);
      this.renderRecentRow();
      if (hadFocus) (this.recentRowEl?.firstElementChild as HTMLElement | null)?.focus();
      return hex;
    }
    return undefined;
  }
}
