import type { CommandOutcome } from "../office/runtime";
import { readSlideNav, selectSlideByIndex, type SlideNav } from "../office/navigation";
import type { ActionCommand, ActionContext } from "./types";

const FIRST_EDGE = "Ya estás en la primera diapositiva.";
const LAST_EDGE = "Ya estás en la última diapositiva.";

/**
 * Lee la navegación, resuelve el índice destino y, si ya coincide con el actual, devuelve el
 * mensaje de extremo sin tocar la selección ni hacer sync de escritura. Si no, selecciona esa
 * diapositiva, hace sync y devuelve "Diapositiva N de M".
 */
async function go(ctx: ActionContext, resolveIndex: (nav: SlideNav) => number, edgeMessage: string): Promise<CommandOutcome> {
  const { context } = ctx;
  const nav = await readSlideNav(context);
  const index = resolveIndex(nav);
  if (index === nav.index) return { changed: 0, message: edgeMessage };
  selectSlideByIndex(context, nav, index);
  await context.sync();
  return { changed: 0, message: `Diapositiva ${index + 1} de ${nav.count}` };
}

function nav(
  id: string,
  label: string,
  icon: string,
  resolveIndex: (nav: SlideNav) => number,
  edgeMessage: string,
): ActionCommand {
  return {
    kind: "action",
    id,
    section: "navigation",
    label,
    icon,
    minShapes: 0,
    needsSelection: false,
    run: (ctx) => go(ctx, resolveIndex, edgeMessage),
  };
}

export const navigationCommands: readonly ActionCommand[] = [
  nav("slide-first", "Primera diapositiva", "slideFirst", () => 0, FIRST_EDGE),
  nav("slide-prev", "Diapositiva anterior", "slidePrev", (n) => Math.max(0, n.index - 1), FIRST_EDGE),
  nav("slide-next", "Diapositiva siguiente", "slideNext", (n) => Math.min(n.count - 1, n.index + 1), LAST_EDGE),
  nav("slide-last", "Última diapositiva", "slideLast", (n) => n.count - 1, LAST_EDGE),
  {
    kind: "action",
    id: "slide-goto",
    section: "navigation",
    label: "Ir a diapositiva",
    icon: "slideGoto",
    minShapes: 0,
    needsSelection: false,
    run: async ({ context, options }) => {
      const n = await readSlideNav(context);
      const target = (options.gotoSlide ?? n.index + 1) - 1;
      if (!selectSlideByIndex(context, n, target)) {
        return { changed: 0, message: `La presentación tiene ${n.count} diapositivas.` };
      }
      await context.sync();
      return { changed: 0, message: `Diapositiva ${target + 1} de ${n.count}` };
    },
  },
];
