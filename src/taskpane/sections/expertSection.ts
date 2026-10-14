import { findCommand } from "../../commands/registry";
import type { ShapeKind } from "../../core/expert";
import type { Rect } from "../../core/types";
import type { SelectionSummary } from "../../office/summary";
import { sanitizeRenamePrefix } from "../sanitize";
import type { SectionContext } from "./types";

const KIND_OPTIONS: ReadonlyArray<{ value: ShapeKind; label: string }> = [
  { value: "all", label: "Todas" },
  { value: "text", label: "Texto" },
  { value: "shape", label: "Formas" },
  { value: "image", label: "Imágenes" },
  { value: "line", label: "Líneas" },
  { value: "table", label: "Tablas" },
  { value: "group", label: "Grupos" },
  { value: "other", label: "Otras" },
];

const INSPECTOR_FIELDS: ReadonlyArray<{ key: keyof Rect; label: string }> = [
  { key: "left", label: "X" },
  { key: "top", label: "Y" },
  { key: "width", label: "Ancho" },
  { key: "height", label: "Alto" },
];

/**
 * Sección "Expert": seleccionar por tipo/nombre, prefijo de "Renombrar numerado" (persistido) e
 * inspector de posición/tamaño en vivo. Los botones de z-order, agrupar/desagrupar y "Renombrar
 * numerado" ya los renderiza el panel (son comandos de la sección, no ocultos); esta clase solo
 * añade los controles que necesitan un valor en vivo (select, campos de texto, inspector).
 */
export class ExpertSection {
  private kindSelect: HTMLSelectElement | undefined;
  private nameInput: HTMLInputElement | undefined;
  private inspectorInputs = new Map<keyof Rect, HTMLInputElement>();
  /** Última caja envolvente recibida en `update()`; restaura un campo del inspector si el usuario
   *  lo deja vacío o con un valor no numérico (ver `applyInspectorField`). */
  private lastBbox: Rect | null = null;

  constructor(private ctx: SectionContext) {}

  render(): HTMLElement[] {
    return [this.kindRow(), this.nameRow(), this.prefixField(), this.inspectorSection()];
  }

  /** Rellena el inspector con la caja envolvente de la selección; sin selección, lo deshabilita. */
  update(summary: SelectionSummary | null): void {
    const bbox = summary && summary.count > 0 ? summary.bbox : null;
    this.lastBbox = bbox;
    for (const [key, input] of this.inspectorInputs) {
      input.disabled = bbox === null;
      if (document.activeElement === input) continue;
      input.value = bbox ? bbox[key].toFixed(1) : "";
    }
  }

  private kindRow(): HTMLElement {
    const row = document.createElement("div");
    row.className = "st-row";
    row.setAttribute("role", "group");
    row.setAttribute("aria-label", "Seleccionar por tipo");

    const select = document.createElement("select");
    select.className = "st-select st-select-inline";
    select.setAttribute("aria-label", "Tipo de forma");
    for (const opt of KIND_OPTIONS) {
      const option = document.createElement("option");
      option.value = opt.value;
      option.textContent = opt.label;
      select.appendChild(option);
    }
    this.kindSelect = select;
    row.appendChild(select);

    const button = document.createElement("button");
    button.type = "button";
    button.className = "st-action";
    button.textContent = "Seleccionar";
    button.setAttribute("aria-label", "Seleccionar por tipo");
    button.addEventListener("click", () => void this.selectKind());
    row.appendChild(button);

    return row;
  }

  private nameRow(): HTMLElement {
    const row = document.createElement("div");
    row.className = "st-row";
    row.setAttribute("role", "group");
    row.setAttribute("aria-label", "Seleccionar por nombre");

    const field = document.createElement("label");
    field.className = "st-field";
    field.append("Nombre contiene");
    const input = document.createElement("input");
    input.type = "text";
    input.className = "st-input-text";
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") void this.selectByName();
    });
    this.nameInput = input;
    field.appendChild(input);
    row.appendChild(field);

    const button = document.createElement("button");
    button.type = "button";
    button.className = "st-action";
    button.textContent = "Seleccionar";
    button.setAttribute("aria-label", "Seleccionar por nombre");
    button.addEventListener("click", () => void this.selectByName());
    row.appendChild(button);

    return row;
  }

  private prefixField(): HTMLElement {
    const field = document.createElement("label");
    field.className = "st-field";
    field.append("Prefijo");
    const input = document.createElement("input");
    input.type = "text";
    input.className = "st-input-text";
    input.maxLength = 40;
    input.value = this.ctx.persisted.renamePrefix;
    input.addEventListener("change", () => {
      const renamePrefix = sanitizeRenamePrefix(input.value);
      input.value = renamePrefix;
      this.ctx.updatePersisted({ renamePrefix });
    });
    field.appendChild(input);
    return field;
  }

  private inspectorSection(): HTMLElement {
    const wrap = document.createElement("div");
    const subtitle = document.createElement("div");
    subtitle.className = "st-subtitle";
    subtitle.textContent = "Inspector";
    wrap.appendChild(subtitle);

    const row = document.createElement("div");
    row.className = "st-row";
    row.setAttribute("role", "group");
    row.setAttribute("aria-label", "Inspector de posición y tamaño");
    for (const f of INSPECTOR_FIELDS) {
      const field = document.createElement("label");
      field.className = "st-field";
      field.append(f.label);
      const input = document.createElement("input");
      input.type = "number";
      input.className = "st-input-num";
      input.step = "0.1";
      input.disabled = true;
      input.addEventListener("change", () => void this.applyInspectorField(f.key, input));
      this.inspectorInputs.set(f.key, input);
      field.appendChild(input);
      row.appendChild(field);
    }
    wrap.appendChild(row);
    return wrap;
  }

  private async selectKind(): Promise<void> {
    const cmd = findCommand("select-kind");
    const select = this.kindSelect;
    if (!cmd || !select) return;
    await this.ctx.run(cmd, { shapeKind: select.value as ShapeKind });
  }

  private async selectByName(): Promise<void> {
    const cmd = findCommand("select-by-name");
    const input = this.nameInput;
    if (!cmd || !input) return;
    await this.ctx.run(cmd, { nameFilter: input.value });
  }

  private async applyInspectorField(key: keyof Rect, input: HTMLInputElement): Promise<void> {
    const cmd = findCommand("inspector-apply");
    if (!cmd) return;
    // Vacío o no numérico: no hay un valor válido que aplicar. Se restaura el último valor
    // conocido de la caja envolvente en vez de mandar un patch con NaN/undefined.
    if (input.value.trim() === "" || !Number.isFinite(input.valueAsNumber)) {
      input.value = this.lastBbox ? this.lastBbox[key].toFixed(1) : "";
      return;
    }
    const value = Math.round(input.valueAsNumber * 10) / 10;
    input.value = value.toFixed(1);
    await this.ctx.run(cmd, { inspector: { [key]: value } });
  }
}
