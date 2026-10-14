import { MAX_INSERT_COUNT, MAX_STICKER_TEXT_LENGTH, MIN_INSERT_COUNT, sanitizeInsert, type InsertSettings } from "../sanitize";
import type { SectionContext } from "./types";

/**
 * Sección "Insert": los botones (columnas, rejilla, chevrones, círculos, sticker, líneas, formas,
 * cuadro de texto) los renderiza el panel como comandos genéricos de la sección; esta clase añade
 * debajo los tres parámetros persistidos que usan: "N" (cantidad, 1..12), el texto de la sticker y
 * su color. Cada cambio se sanea (`sanitizeInsert`), se refleja de vuelta en los controles y se
 * persiste.
 */
export class InsertSection {
  private countInput: HTMLInputElement | undefined;
  private textInput: HTMLInputElement | undefined;
  private colorInput: HTMLInputElement | undefined;

  constructor(private ctx: SectionContext) {}

  render(): HTMLElement[] {
    const row = document.createElement("div");
    row.className = "st-row";
    row.setAttribute("role", "group");
    row.setAttribute("aria-label", "Parámetros de inserción");
    row.append(this.countField(), this.textField(), this.colorField());
    return [row];
  }

  private countField(): HTMLElement {
    const field = document.createElement("label");
    field.className = "st-field";
    field.append("N");
    const input = document.createElement("input");
    input.type = "number";
    input.className = "st-input-num";
    input.min = String(MIN_INSERT_COUNT);
    input.max = String(MAX_INSERT_COUNT);
    input.step = "1";
    input.value = String(this.ctx.persisted.insert.count);
    input.title = "Cantidad de columnas, chevrones o círculos (1..12)";
    input.addEventListener("change", () => this.persist());
    this.countInput = input;
    field.appendChild(input);
    return field;
  }

  private textField(): HTMLElement {
    const field = document.createElement("label");
    field.className = "st-field";
    field.append("Sticker");
    const input = document.createElement("input");
    input.type = "text";
    input.className = "st-input-text";
    input.maxLength = MAX_STICKER_TEXT_LENGTH;
    input.value = this.ctx.persisted.insert.stickerText;
    input.placeholder = "BORRADOR";
    input.addEventListener("change", () => this.persist());
    this.textInput = input;
    field.appendChild(input);
    return field;
  }

  private colorField(): HTMLElement {
    const field = document.createElement("label");
    field.className = "st-field";
    field.append("Color");
    const input = document.createElement("input");
    input.type = "color";
    input.className = "st-input-color";
    input.value = this.ctx.persisted.insert.stickerColor;
    input.setAttribute("aria-label", "Color de la sticker");
    input.addEventListener("change", () => this.persist());
    this.colorInput = input;
    field.appendChild(input);
    return field;
  }

  /** Sanea los tres campos juntos (para que lo persistido sea siempre una `InsertSettings` válida)
   *  y refleja los valores saneados de vuelta en los controles. `valueAsNumber` da `NaN` con un
   *  campo vacío: `sanitizeInsert` lo convierte en el valor por defecto (3). */
  private sanitize(): InsertSettings {
    const insert = sanitizeInsert({
      count: this.countInput?.valueAsNumber,
      stickerText: this.textInput?.value,
      stickerColor: this.colorInput?.value,
    });
    if (this.countInput) this.countInput.value = String(insert.count);
    if (this.textInput) this.textInput.value = insert.stickerText;
    if (this.colorInput) this.colorInput.value = insert.stickerColor;
    return insert;
  }

  private persist(): void {
    this.ctx.updatePersisted({ insert: this.sanitize() });
  }
}
