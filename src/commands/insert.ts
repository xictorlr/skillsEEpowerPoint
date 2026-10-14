import { BASIC_SHAPE_RECT, chevronChainLayout, columnsLayout, gridLayout, numberedCirclesLayout, stickerRect } from "../core/insert";
import type { SlideSize } from "../core/types";
import { addGeometric, addLine, addTextBoxes, selectCreated, type GeometricShapeType, type LineType } from "../office/insert";
import { readSlideNav } from "../office/navigation";
import type { CommandOutcome } from "../office/runtime";
import { getSlideSize } from "../office/slideSize";
import { DEFAULT_ACCENT, resolveAccent } from "./accent";
import type { ActionCommand, ActionContext } from "./types";

/** Texto de relleno de los cuadros de texto creados (columnas, rejilla, cuadro suelto). */
const PLACEHOLDER_TEXT = "Texto";
/** Texto de la sticker cuando `options.insert.stickerText` está vacío. */
const DEFAULT_STICKER_TEXT = "BORRADOR";
/** Color del texto sobre el relleno de acento (chevrones, círculos numerados). */
const ON_ACCENT_TEXT = "#FFFFFF";
/** Estilo común de las etiquetas sobre acento (chevrones, círculos numerados): sin contorno, texto
 *  blanco en negrita y centrado. Cada comando añade `fill`, `text` y `fontSize`. */
const ACCENT_LABEL = { line: null, fontColor: ON_ACCENT_TEXT, bold: true, centered: true } as const;

/** Lo que recibe `build`: la diapositiva activa, su tamaño y el color de acento (solo resuelto
 *  contra el tema si el comando lo pidió con `{ accent: true }`; si no, `DEFAULT_ACCENT`). */
interface InsertTarget {
  slide: PowerPoint.Slide;
  size: SlideSize;
  accent: string;
}

/** Opciones de `insertShapes`: `accent` pide resolver el color de acento del tema. */
interface InsertOptions {
  accent?: boolean;
}

/**
 * Esqueleto común de todos los comandos de Insert: lee la navegación (diapositiva activa) y el
 * tamaño de diapositiva, crea las formas con `build`, carga sus ids en un sync de escritura, las
 * selecciona (`selectCreated`) en un segundo sync y devuelve `{ changed, message }`.
 *
 * Syncs: 1 (`readSlideNav`) + 0..1 (`getSlideSize`, 0 con `options.slideSize` forzado) + 0..1
 * (`resolveAccent`, solo con `{ accent: true }` y PowerPointApi 1.10) + 2 de escritura. Los dos syncs
 * de escritura son inevitables: `setSelectedShapes` necesita los ids reales, que Office.js solo
 * asigna al sincronizar la creación.
 */
async function insertShapes(
  { context, options, slideSizeCache }: ActionContext,
  { accent: withAccent = false }: InsertOptions,
  build: (target: InsertTarget) => PowerPoint.Shape[],
  message: (n: number) => string,
): Promise<CommandOutcome> {
  const nav = await readSlideNav(context);
  const size = await getSlideSize(context, slideSizeCache, options.slideSize);
  const accent = withAccent ? await resolveAccent(context) : DEFAULT_ACCENT;
  const slide = context.presentation.slides.getItem(nav.slideId);

  const shapes = build({ slide, size, accent });
  for (const s of shapes) s.load("id");
  await context.sync();

  selectCreated(slide, shapes);
  await context.sync();
  return { changed: shapes.length, message: message(shapes.length) };
}

const base = { kind: "action", section: "insert", minShapes: 0, needsSelection: false } as const;

const insertColumns: ActionCommand = {
  ...base,
  id: "insert-columns",
  label: "Insertar N columnas de texto",
  icon: "insertColumns",
  run: (ctx) =>
    insertShapes(
      ctx,
      {},
      ({ slide, size }) => {
        const rects = columnsLayout(size, ctx.options.insert.count);
        return addTextBoxes(slide, rects, rects.map(() => PLACEHOLDER_TEXT));
      },
      (n) => `${n} columna(s) de texto insertada(s).`,
    ),
};

const insertGrid: ActionCommand = {
  ...base,
  id: "insert-grid",
  label: "Insertar rejilla 2×2",
  icon: "insertGrid",
  run: (ctx) =>
    insertShapes(
      ctx,
      {},
      ({ slide, size }) => {
        const rects = gridLayout(size, 2, 2);
        return addTextBoxes(slide, rects, rects.map(() => PLACEHOLDER_TEXT));
      },
      () => "Rejilla 2×2 insertada (4 cuadros).",
    ),
};

const insertChevrons: ActionCommand = {
  ...base,
  id: "insert-chevrons",
  label: "Insertar cadena de chevrones",
  icon: "insertChevrons",
  run: (ctx) =>
    insertShapes(
      ctx,
      { accent: true },
      ({ slide, size, accent }) =>
        chevronChainLayout(size, ctx.options.insert.count).map((rect, i) =>
          addGeometric(slide, "Chevron", rect, { ...ACCENT_LABEL, fill: accent, text: `Paso ${i + 1}`, fontSize: 12 }),
        ),
      (n) => `${n} chevrón(es) insertado(s).`,
    ),
};

const insertCircles: ActionCommand = {
  ...base,
  id: "insert-circles",
  label: "Insertar círculos numerados",
  icon: "insertCircles",
  run: (ctx) =>
    insertShapes(
      ctx,
      { accent: true },
      ({ slide, size, accent }) =>
        numberedCirclesLayout(size, ctx.options.insert.count).map((rect, i) =>
          addGeometric(slide, "Ellipse", rect, { ...ACCENT_LABEL, fill: accent, text: String(i + 1), fontSize: 14 }),
        ),
      (n) => `${n} círculo(s) numerado(s) insertado(s).`,
    ),
};

const insertSticker: ActionCommand = {
  ...base,
  id: "insert-sticker",
  label: "Insertar sticker",
  icon: "insertSticker",
  run: (ctx) =>
    insertShapes(
      ctx,
      {},
      ({ slide, size }) => [
        addGeometric(slide, "RoundRectangle", stickerRect(size), {
          fill: ctx.options.insert.stickerColor,
          line: null,
          text: ctx.options.insert.stickerText || DEFAULT_STICKER_TEXT,
          bold: true,
          fontSize: 12,
          fontColor: "#000000",
          centered: true,
        }),
      ],
      () => "Sticker insertado.",
    ),
};

/** Comando que inserta una única línea (`kind`) en el rectángulo básico centrado. */
function lineCommand(id: string, label: string, icon: string, kind: LineType, message: string): ActionCommand {
  return {
    ...base,
    id,
    label,
    icon,
    run: (ctx) => insertShapes(ctx, {}, ({ slide, size }) => [addLine(slide, kind, BASIC_SHAPE_RECT(size))], () => message),
  };
}

/** Comando que inserta una única forma geométrica (`type`) con el estilo por defecto del tema. */
function shapeCommand(id: string, label: string, icon: string, type: GeometricShapeType): ActionCommand {
  return {
    ...base,
    id,
    label,
    icon,
    run: (ctx) => insertShapes(ctx, {}, ({ slide, size }) => [addGeometric(slide, type, BASIC_SHAPE_RECT(size))], () => "Forma insertada."),
  };
}

const insertTextbox: ActionCommand = {
  ...base,
  id: "insert-textbox",
  label: "Insertar cuadro de texto",
  icon: "insertTextbox",
  run: (ctx) =>
    insertShapes(
      ctx,
      {},
      ({ slide, size }) => addTextBoxes(slide, [BASIC_SHAPE_RECT(size)], [PLACEHOLDER_TEXT]),
      () => "Cuadro de texto insertado.",
    ),
};

export const insertCommands: readonly ActionCommand[] = [
  insertColumns,
  insertGrid,
  insertChevrons,
  insertCircles,
  insertSticker,
  lineCommand("insert-line", "Insertar línea recta", "insertLine", "Straight", "Línea insertada."),
  lineCommand("insert-elbow", "Insertar conector en ángulo", "insertElbow", "Elbow", "Conector insertado."),
  shapeCommand("insert-rect", "Insertar rectángulo", "insertRect", "Rectangle"),
  shapeCommand("insert-rounded", "Insertar rectángulo redondeado", "insertRounded", "RoundRectangle"),
  shapeCommand("insert-ellipse", "Insertar elipse", "insertEllipse", "Ellipse"),
  insertTextbox,
];
