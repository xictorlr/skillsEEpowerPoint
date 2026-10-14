import { findCommand } from "../../commands/registry";
import type { SelectionSummary } from "../../office/summary";
import type { SectionContext } from "./types";

/**
 * Sección "Navigation": los botones de ir a primera/anterior/siguiente/última diapositiva ya los
 * renderiza el panel (son los comandos de la sección, ver `commandsBySection`). Esta clase añade
 * el indicador "N / M" y el campo "Ir a diapositiva", que ejecuta `slide-goto` con un `gotoSlide`
 * explícito (el botón estándar de `slide-goto`, sin este campo, repite la diapositiva actual).
 */
export class NavigationSection {
  private labelEl: HTMLElement | undefined;
  private gotoInput: HTMLInputElement | undefined;

  constructor(private ctx: SectionContext) {}

  render(): HTMLElement[] {
    return [this.positionRow(), this.gotoRow()];
  }

  /** Actualiza la etiqueta "N / M" y el máximo del campo "Ir a" según el resumen de selección. */
  update(summary: SelectionSummary | null): void {
    const count = summary?.slideCount ?? 0;
    const index = summary?.slideIndex ?? 0;
    if (this.labelEl) this.labelEl.textContent = count > 0 ? `${index + 1} / ${count}` : "— / —";
    if (this.gotoInput) {
      this.gotoInput.max = count > 0 ? String(count) : "";
      if (document.activeElement !== this.gotoInput) {
        this.gotoInput.placeholder = count > 0 ? String(index + 1) : "";
      }
    }
  }

  private positionRow(): HTMLElement {
    const field = document.createElement("div");
    field.className = "st-field";
    field.setAttribute("role", "status");
    field.setAttribute("aria-label", "Diapositiva actual");
    field.append("Diapositiva ");
    this.labelEl = document.createElement("strong");
    this.labelEl.textContent = "— / —";
    field.appendChild(this.labelEl);
    return field;
  }

  private gotoRow(): HTMLElement {
    const row = document.createElement("div");
    row.className = "st-row";
    row.setAttribute("role", "group");
    row.setAttribute("aria-label", "Ir a diapositiva");

    const field = document.createElement("label");
    field.className = "st-field";
    field.append("Ir a");
    const input = document.createElement("input");
    input.type = "number";
    input.min = "1";
    input.style.maxWidth = "160px";
    this.gotoInput = input;
    field.appendChild(input);
    row.appendChild(field);

    const button = document.createElement("button");
    button.type = "button";
    button.textContent = "Ir";
    button.addEventListener("click", () => void this.goto());
    row.appendChild(button);

    return row;
  }

  private async goto(): Promise<void> {
    const cmd = findCommand("slide-goto");
    const input = this.gotoInput;
    if (!cmd || !input) return;
    const raw = Number(input.value);
    if (!Number.isFinite(raw) || raw < 1) return;
    // Fuera de rango: lo valida `slide-goto` (ver office/navigation.ts selectSlideByIndex), que
    // devuelve "La presentación tiene N diapositivas." sin tocar la selección.
    await this.ctx.run(cmd, { gotoSlide: Math.round(raw) });
  }
}
