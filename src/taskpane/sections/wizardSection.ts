import { findCommand } from "../../commands/registry";
import { MAX_AGENDA_TITLE_LENGTH, sanitizeAgenda } from "../sanitize";
import type { SectionContext } from "./types";

/**
 * Sección "Wizards": el asistente "Crear agenda" (título + lista de secciones, una por línea).
 * `agenda-create` está oculto (ver `commands/registry.ts`): el panel no le renderiza un botón
 * genérico, así que esta sección es el único contenido visible de Wizards (no hay fila de botones
 * porque `commandsBySection("wizards")` queda vacía).
 */
export class WizardSection {
  private titleInput: HTMLInputElement | undefined;
  private sectionsTextarea: HTMLTextAreaElement | undefined;

  constructor(private ctx: SectionContext) {}

  render(): HTMLElement[] {
    return [this.titleField(), this.sectionsField(), this.createButton()];
  }

  private titleField(): HTMLElement {
    const field = document.createElement("label");
    field.className = "st-field";
    field.append("Título");
    const input = document.createElement("input");
    input.type = "text";
    input.className = "st-input-text";
    input.maxLength = MAX_AGENDA_TITLE_LENGTH;
    input.value = this.ctx.persisted.agenda.title;
    input.addEventListener("change", () => this.persist());
    this.titleInput = input;
    field.appendChild(input);
    return field;
  }

  private sectionsField(): HTMLElement {
    const field = document.createElement("label");
    field.className = "st-field-block";
    field.append("Secciones (una por línea)");
    const textarea = document.createElement("textarea");
    textarea.rows = 5;
    textarea.className = "st-textarea";
    textarea.value = this.ctx.persisted.agenda.sectionsText;
    textarea.addEventListener("change", () => this.persist());
    this.sectionsTextarea = textarea;
    field.appendChild(textarea);
    return field;
  }

  private createButton(): HTMLElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "st-action";
    button.textContent = "Crear agenda";
    button.setAttribute("aria-label", "Crear agenda");
    button.addEventListener("click", () => void this.create());
    return button;
  }

  /** Sanea título y texto de secciones juntos (para que uno no pueda quedar desincronizado con el
   *  otro en lo persistido) y los guarda; refleja el valor saneado de vuelta en los controles. */
  private persist(): void {
    const agenda = sanitizeAgenda({
      title: this.titleInput?.value ?? "",
      sectionsText: this.sectionsTextarea?.value ?? "",
    });
    if (this.titleInput) this.titleInput.value = agenda.title;
    if (this.sectionsTextarea) this.sectionsTextarea.value = agenda.sectionsText;
    this.ctx.updatePersisted({ agenda });
  }

  private async create(): Promise<void> {
    const cmd = findCommand("agenda-create");
    if (!cmd || !this.titleInput || !this.sectionsTextarea) return;
    await this.ctx.run(cmd, {
      agenda: { title: this.titleInput.value, sectionsText: this.sectionsTextarea.value },
    });
  }
}
