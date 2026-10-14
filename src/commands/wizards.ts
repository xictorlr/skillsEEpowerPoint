import { agendaLayout, parseAgendaSections } from "../core/check";
import { addSlideAfter, addTextBox } from "../office/deck";
import { readSlideNav } from "../office/navigation";
import { getSlideSize } from "../office/slideSize";
import { getThemeColors } from "../office/theme";
import type { ActionCommand } from "./types";

/** Color de acento por defecto cuando `getThemeColors` no está disponible (sin PowerPointApi 1.10)
 *  o no encuentra "Accent1". */
const DEFAULT_ACCENT = "#1F6FEB";
/** Color de las secciones no resaltadas de la agenda. */
const UNHIGHLIGHTED_COLOR = "#7F7F7F";

/** `agenda-create` (oculto; lo dispara el asistente "Crear agenda" de la sección Wizards): crea una
 *  diapositiva por sección, cada una con su título y la lista completa de secciones, resaltando la
 *  propia. Las diapositivas se insertan en orden justo después de la diapositiva activa. */
const agendaCreate: ActionCommand = {
  kind: "action",
  id: "agenda-create",
  section: "wizards",
  label: "Crear agenda",
  icon: "agenda",
  minShapes: 0,
  run: async ({ context, options }) => {
    const sections = parseAgendaSections(options.agenda.sectionsText);
    if (sections.length < 1) return { changed: 0, message: "Escribe al menos una sección (una por línea)." };

    const nav = await readSlideNav(context);
    const size = await getSlideSize(context, new Map(), options.slideSize);
    const themeColors = await getThemeColors(context);
    const accent = themeColors?.find((c) => c.name === "Accent1")?.hex ?? DEFAULT_ACCENT;

    for (let i = 0; i < sections.length; i++) {
      const slide = await addSlideAfter(context, nav.index + i);
      const spec = agendaLayout(size, { title: options.agenda.title, sections }, i);
      addTextBox(slide, spec.title.text, spec.title.rect, { size: 28, bold: true, color: null });
      for (const item of spec.items) {
        addTextBox(slide, item.text, item.rect, {
          size: 18,
          bold: item.highlighted,
          color: item.highlighted ? accent : UNHIGHLIGHTED_COLOR,
        });
      }
    }

    await context.sync();
    return { changed: sections.length, message: `${sections.length} diapositiva(s) de agenda creada(s).` };
  },
};

export const wizardsCommands: readonly ActionCommand[] = [agendaCreate];
