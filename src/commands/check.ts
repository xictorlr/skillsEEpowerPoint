import { distinctFonts, isOffSlide } from "../core/check";
import { deleteShapes, loadFontNames, loadHasText, readDeckShapes, readFontName } from "../office/deck";
import { readSlideNav, readSlideShapes, selectShapes } from "../office/navigation";
import { getSlideSize } from "../office/slideSize";
import type { ActionCommand } from "./types";

/** Máximo de nombres de fuente que muestra `check-fonts` antes de recortar con "…". */
const MAX_FONTS_SHOWN = 12;

/** `check-empty-text-slide`: en la diapositiva activa, borra los cuadros de texto (`TextBox`) sin
 *  texto (`textFrame.hasText === false`). */
const checkEmptyTextSlide: ActionCommand = {
  kind: "action",
  id: "check-empty-text-slide",
  section: "check",
  label: "Eliminar cuadros de texto vacíos (diapositiva)",
  icon: "emptyTextSlide",
  minShapes: 0,
  run: async ({ context }) => {
    const nav = await readSlideNav(context);
    const { shapes, proxies, textCapable } = await readDeckShapes(context);
    const ids = shapes.filter((d) => d.slideId === nav.slideId && d.shape.type === "TextBox").map((d) => d.shape.id);
    loadHasText(proxies, ids, textCapable);
    await context.sync();

    const emptyIds = ids.filter((id) => proxies.get(id)?.textFrame.hasText === false);
    if (emptyIds.length === 0) return { changed: 0, message: "No hay cuadros vacíos." };

    deleteShapes(proxies, emptyIds);
    await context.sync();
    return { changed: emptyIds.length, message: `${emptyIds.length} cuadro(s) vacío(s) eliminado(s).` };
  },
};

/** `check-empty-text-deck`: igual que `check-empty-text-slide` pero en todo el mazo. Destructivo
 *  sobre varias diapositivas a la vez: el panel pide confirmación (`confirm: true`). */
const checkEmptyTextDeck: ActionCommand = {
  kind: "action",
  id: "check-empty-text-deck",
  section: "check",
  label: "Eliminar cuadros de texto vacíos (mazo)",
  icon: "emptyTextDeck",
  minShapes: 0,
  confirm: true,
  run: async ({ context }) => {
    const { shapes, proxies, textCapable } = await readDeckShapes(context);
    const ids = shapes.filter((d) => d.shape.type === "TextBox").map((d) => d.shape.id);
    loadHasText(proxies, ids, textCapable);
    await context.sync();

    const emptyIds = ids.filter((id) => proxies.get(id)?.textFrame.hasText === false);
    if (emptyIds.length === 0) return { changed: 0, message: "No hay cuadros vacíos." };

    deleteShapes(proxies, emptyIds);
    await context.sync();
    return { changed: emptyIds.length, message: `${emptyIds.length} cuadro(s) vacío(s) eliminado(s).` };
  },
};

/** `check-offslide-select`: selecciona las formas de la diapositiva activa que quedan completamente
 *  fuera del lienzo (ver `isOffSlide`). Sin coincidencias, no toca la selección actual. */
const checkOffslideSelect: ActionCommand = {
  kind: "action",
  id: "check-offslide-select",
  section: "check",
  label: "Seleccionar formas fuera de la diapositiva",
  icon: "offSlide",
  minShapes: 0,
  run: async ({ context, options, slideSizeCache }) => {
    const nav = await readSlideNav(context);
    const size = await getSlideSize(context, slideSizeCache, options.slideSize);
    const shapes = await readSlideShapes(context, nav.slideId);
    const offSlide = shapes.filter((s) => isOffSlide(s, size));
    if (offSlide.length === 0) return { changed: 0, message: "No hay formas fuera de la diapositiva." };

    selectShapes(
      context,
      nav.slideId,
      offSlide.map((s) => s.id),
    );
    await context.sync();
    return { changed: offSlide.length, message: `${offSlide.length} forma(s) fuera de la diapositiva seleccionada(s).` };
  },
};

/** `check-fonts`: lista los nombres de fuente distintos usados en todo el mazo (máx. 12, luego "…").
 *  Una forma con fuente mixta dentro de su propio texto (p. ej. dos fragmentos con fuentes
 *  distintas en el mismo cuadro) reporta `font.name === null` (comportamiento de Office.js para
 *  rangos con formato no uniforme); `readFontName`/`distinctFonts` la tratan igual que una forma
 *  sin fuente y la omiten del listado, sin que eso cuente como error. */
const checkFonts: ActionCommand = {
  kind: "action",
  id: "check-fonts",
  section: "check",
  label: "Fuentes del mazo",
  icon: "fonts",
  minShapes: 0,
  run: async ({ context }) => {
    const { shapes, proxies, textCapable } = await readDeckShapes(context);
    const ids = shapes.filter((d) => textCapable.has(d.shape.id)).map((d) => d.shape.id);
    loadFontNames(proxies, ids, textCapable);
    await context.sync();

    const names = distinctFonts(ids.map((id) => readFontName(proxies.get(id)!)));
    if (names.length === 0) return { changed: 0, message: "No hay texto en la presentación." };

    const shown = names.length > MAX_FONTS_SHOWN ? [...names.slice(0, MAX_FONTS_SHOWN), "…"] : names;
    return { changed: 0, message: `Fuentes: ${shown.join(", ")}` };
  },
};

export const checkCommands: readonly ActionCommand[] = [checkEmptyTextSlide, checkEmptyTextDeck, checkOffslideSelect, checkFonts];
