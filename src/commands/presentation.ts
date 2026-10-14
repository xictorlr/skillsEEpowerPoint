import { supportsApi } from "../office/capabilities";
import { createPresentation, exportSlides } from "../office/deck";
import type { ActionCommand } from "./types";

/** `new-presentation`: abre una presentación nueva en vacío en otra ventana. */
const newPresentation: ActionCommand = {
  kind: "action",
  id: "new-presentation",
  section: "presentation",
  label: "Nueva presentación",
  icon: "newPresentation",
  minShapes: 0,
  run: async () => {
    await createPresentation();
    return { changed: 0, message: "Nueva presentación creada." };
  },
};

/** `export-selected-slides`: exporta las diapositivas seleccionadas a una presentación nueva
 *  (requiere PowerPointApi 1.10; ver `exportSlides`). Comprueba la API antes de cualquier `sync()`:
 *  sin soporte, no tiene sentido pagar la lectura de la selección. */
const exportSelectedSlides: ActionCommand = {
  kind: "action",
  id: "export-selected-slides",
  section: "presentation",
  label: "Exportar diapositivas seleccionadas",
  icon: "exportSlides",
  minShapes: 0,
  run: async ({ context }) => {
    if (!supportsApi("1.10")) return { changed: 0, message: "Exportar requiere PowerPoint reciente (API 1.10)." };

    const selected = context.presentation.getSelectedSlides();
    selected.load("items/id");
    await context.sync();
    const ids = selected.items.map((s) => s.id);
    if (ids.length === 0) return { changed: 0, message: "Selecciona al menos una diapositiva." };

    const base64 = await exportSlides(context, ids);
    if (base64 === null) return { changed: 0, message: "Exportar requiere PowerPoint reciente (API 1.10)." };

    await createPresentation(base64);
    return { changed: ids.length, message: `${ids.length} diapositiva(s) exportada(s) a una nueva presentación.` };
  },
};

export const presentationCommands: readonly ActionCommand[] = [newPresentation, exportSelectedSlides];
