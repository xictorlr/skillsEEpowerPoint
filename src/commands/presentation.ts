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
  needsSelection: false,
  run: async () => {
    await createPresentation();
    return { changed: 0, message: "Nueva presentación creada." };
  },
};

/** `export-selected-slides`: exporta las diapositivas seleccionadas a una presentación nueva
 *  (requiere PowerPointApi 1.10; ver `exportSlides`). `requiresApi: "1.10"` ya hace que
 *  `disabledReason` (registry.ts) rechace el comando antes de llegar aquí sin esa API -- el
 *  comando tampoco se lista en el panel (`commandsBySection`) -- así que `run` no necesita volver
 *  a comprobarlo. */
const exportSelectedSlides: ActionCommand = {
  kind: "action",
  id: "export-selected-slides",
  section: "presentation",
  label: "Exportar diapositivas seleccionadas",
  icon: "exportSlides",
  minShapes: 0,
  requiresApi: "1.10",
  needsSelection: false,
  run: async ({ context }) => {
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
