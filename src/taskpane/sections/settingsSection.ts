import {
  SLIDE_SIZE_PRESETS,
  sanitizeFontSizeStep,
  sanitizeMatrix,
  sanitizeSlideSizePreset,
  type SlideSizePreset,
} from "../sanitize";
import type { SectionContext } from "./types";

const SLIDE_SIZE_LABELS: Record<SlideSizePreset, string> = {
  auto: "Auto (exacto en PowerPoint reciente)",
  "16x9": "16:9 (960×540)",
  "16x9-legacy": "16:9 clásico (720×405)",
  "4x3": "4:3 (720×540)",
  a4: "A4 (780×540)",
  letter: "Carta (720×540)",
};

/** Sección "Ajustes": tamaño de diapositiva forzado, parámetros de la matriz y paso de tamaño de fuente. */
export class SettingsSection {
  constructor(private ctx: SectionContext) {}

  render(): HTMLElement {
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
    const fontStepRow = document.createElement("div");
    fontStepRow.className = "st-row";
    fontStepRow.appendChild(this.fontStepField());
    sec.appendChild(fontStepRow);
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
    select.value = this.ctx.persisted.slideSizePreset;
    select.addEventListener("change", () => {
      const slideSizePreset = sanitizeSlideSizePreset(select.value);
      select.value = slideSizePreset;
      this.ctx.updatePersisted({ slideSizePreset });
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
      input.value = String(this.ctx.persisted.matrix[key]);
      input.addEventListener("change", () => {
        const raw = Number(input.value);
        const candidate = {
          ...this.ctx.persisted.matrix,
          [key]: Number.isFinite(raw) ? Math.max(min, raw) : this.ctx.persisted.matrix[key],
        };
        const matrix = sanitizeMatrix(candidate);
        input.value = String(matrix[key]);
        this.ctx.updatePersisted({ matrix });
      });
      field.appendChild(input);
      return field;
    };
    wrap.append(make("Cols", "columns", 1, "1"), make("Gap X", "gapX", 0), make("Gap Y", "gapY", 0));
    return wrap;
  }

  private fontStepField(): HTMLElement {
    const field = document.createElement("label");
    field.className = "st-field";
    field.append("Paso de fuente");
    const input = document.createElement("input");
    input.type = "number";
    input.step = "0.5";
    input.min = "0.5";
    input.max = "20";
    input.value = String(this.ctx.persisted.fontSizeStep);
    input.addEventListener("change", () => {
      const fontSizeStep = sanitizeFontSizeStep(Number(input.value));
      input.value = String(fontSizeStep);
      this.ctx.updatePersisted({ fontSizeStep });
    });
    field.appendChild(input);
    return field;
  }
}
