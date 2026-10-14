import type { FormatParts } from "../../core/format";
import type { SectionContext } from "./types";

const PASTE_PART_FIELDS: ReadonlyArray<{ key: keyof FormatParts; label: string }> = [
  { key: "fill", label: "Relleno" },
  { key: "line", label: "Línea" },
  { key: "font", label: "Fuente" },
  { key: "size", label: "Tamaño" },
];

/** Sección "Format": checkboxes de qué partes pega "Pegar formato". */
export class FormatSection {
  constructor(private ctx: SectionContext) {}

  render(): HTMLElement[] {
    return [this.pastePartsRow()];
  }

  private pastePartsRow(): HTMLElement {
    const row = document.createElement("div");
    row.className = "st-row";
    row.setAttribute("role", "group");
    row.setAttribute("aria-label", "Partes a pegar");
    for (const part of PASTE_PART_FIELDS) {
      const label = document.createElement("label");
      label.className = "st-check";
      const input = document.createElement("input");
      input.type = "checkbox";
      input.checked = this.ctx.persisted.pasteParts[part.key];
      input.addEventListener("change", () => {
        const pasteParts = { ...this.ctx.persisted.pasteParts, [part.key]: input.checked };
        this.ctx.updatePersisted({ pasteParts });
        this.ctx.refreshButtons();
      });
      label.append(input, ` ${part.label}`);
      row.appendChild(label);
    }
    return row;
  }
}
