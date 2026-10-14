/**
 * Comprobaciones de mazo (deck checks) y layout de la diapositiva de agenda.
 * Todas las funciones usan el marco sin rotar (left/top/width/height) tal como lo expone Office.js.
 */
import type { Rect, SlideSize } from "./types";

/** true si la forma está completamente fuera del lienzo (marco sin rotar). */
export function isOffSlide(shape: Rect, slide: SlideSize): boolean {
  const right = shape.left + shape.width;
  const bottom = shape.top + shape.height;
  return right <= 0 || shape.left >= slide.width || bottom <= 0 || shape.top >= slide.height;
}

/** Nombres de fuente únicos, ordenados alfabéticamente, sin nulos ni vacíos. */
export function distinctFonts(names: Array<string | null | undefined>): string[] {
  const set = new Set<string>();
  for (const raw of names) {
    if (raw == null) continue;
    const trimmed = raw.trim();
    if (trimmed.length > 0) set.add(trimmed);
  }
  return [...set].sort((a, b) => a.localeCompare(b, "es"));
}

const MAX_AGENDA_SECTIONS = 20;

/** Separa el texto del textarea en secciones no vacías (una por línea, recortadas), máximo 20. */
export function parseAgendaSections(text: string): string[] {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
    .slice(0, MAX_AGENDA_SECTIONS);
}

export interface AgendaItem {
  text: string;
  rect: Rect;
  highlighted: boolean;
}

export interface AgendaSlideSpec {
  title: { text: string; rect: Rect };
  items: AgendaItem[];
}

export interface AgendaOptions {
  title: string;
  sections: string[];
}

const MARGIN_X_RATIO = 0.08;
const FULL_WIDTH_RATIO = 0.84;
const TITLE_TOP_RATIO = 0.1;
const TITLE_HEIGHT_RATIO = 0.12;
const ITEMS_TOP_RATIO = 0.28;
const ROW_HEIGHT_RATIO = 0.08;
const MAX_ROWS_PER_COLUMN = 10;
const COLUMN_WIDTH_RATIO = 0.42;
const COLUMN_GUTTER_RATIO = 0.02;

/**
 * Layout de una diapositiva de agenda para `currentIndex` (0-based): título arriba, lista de secciones
 * en columna, la actual resaltada. Márgenes del 8 % lateral y 10 % superior; filas del 8 % de alto
 * hasta un máximo de 10 filas; si hay más secciones, dos columnas.
 */
export function agendaLayout(slide: SlideSize, opts: AgendaOptions, currentIndex: number): AgendaSlideSpec {
  const { width: w, height: h } = slide;
  const marginX = MARGIN_X_RATIO * w;
  const fullWidth = FULL_WIDTH_RATIO * w;
  const rowsTop = ITEMS_TOP_RATIO * h;
  const rowHeight = ROW_HEIGHT_RATIO * h;

  const title = {
    text: opts.title,
    rect: {
      left: marginX,
      top: TITLE_TOP_RATIO * h,
      width: fullWidth,
      height: TITLE_HEIGHT_RATIO * h,
    },
  };

  const sections = opts.sections;
  const twoColumns = sections.length > MAX_ROWS_PER_COLUMN;

  if (!twoColumns) {
    const items: AgendaItem[] = sections.map((text, i) => ({
      text,
      rect: { left: marginX, top: rowsTop + i * rowHeight, width: fullWidth, height: rowHeight },
      highlighted: i === currentIndex,
    }));
    return { title, items };
  }

  const leftCount = Math.ceil(sections.length / 2);
  const columnWidth = COLUMN_WIDTH_RATIO * w;
  const leftColumnLeft = marginX;
  const rightColumnLeft = marginX + columnWidth + COLUMN_GUTTER_RATIO * w;

  const items: AgendaItem[] = sections.map((text, i) => {
    const inLeftColumn = i < leftCount;
    const columnLeft = inLeftColumn ? leftColumnLeft : rightColumnLeft;
    const rowIndex = inLeftColumn ? i : i - leftCount;
    return {
      text,
      rect: { left: columnLeft, top: rowsTop + rowIndex * rowHeight, width: columnWidth, height: rowHeight },
      highlighted: i === currentIndex,
    };
  });

  return { title, items };
}
