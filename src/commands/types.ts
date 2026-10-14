import type { MatrixOptions } from "../core/geometry";
import type { ColorTarget, FormatParts, FormatPatch, ShapeFormat } from "../core/format";
import type { ShapeKind } from "../core/expert";
import type { Patch, Rect, SelectionContext, SlideSize } from "../core/types";
import type { TextPatch, TextShape } from "../core/text";
import type { CommandOutcome } from "../office/runtime";

export type SectionId =
  | "alignment"
  | "size"
  | "format"
  | "color"
  | "text"
  | "navigation"
  | "expert"
  | "check"
  | "presentation"
  | "wizards";

export const SECTION_LABELS: Readonly<Record<SectionId, string>> = {
  alignment: "Alignment",
  size: "Size",
  format: "Format",
  color: "Color",
  text: "Text",
  navigation: "Navigation",
  expert: "Expert",
  check: "Check",
  presentation: "Presentation",
  wizards: "Wizards",
};

export interface CommandOptions {
  matrix: MatrixOptions;
  /** Tamaño de diapositiva forzado por el usuario; si falta se deduce automáticamente. */
  slideSize?: SlideSize;
  /** Color elegido en la paleta ("#RRGGBB"); lo usan los comandos de color. */
  color?: string;
  /** A qué se aplica el color: relleno, línea o texto. */
  colorTarget: ColorTarget;
  /** Qué partes pega "Pegar formato". */
  pasteParts: FormatParts;
  /** Puntos que suman/restan "Aumentar"/"Reducir tamaño de fuente". */
  fontSizeStep: number;
  /** Diapositiva destino (1-based) para "Ir a diapositiva". */
  gotoSlide?: number;
  /** Subcadena del nombre para "Seleccionar por nombre". */
  nameFilter: string;
  /** Prefijo para "Renombrar numerado". */
  renamePrefix: string;
  /** Caja destino (parcial) para el inspector; los campos ausentes no se tocan. */
  inspector?: Partial<Rect>;
  /** Categoría para "Seleccionar por tipo". */
  shapeKind: ShapeKind;
  /** Título y texto del textarea (una sección por línea) del asistente "Crear agenda". */
  agenda: { title: string; sectionsText: string };
}

export interface ActionContext {
  context: PowerPoint.RequestContext;
  selection: SelectionContext;
  proxies: Map<string, PowerPoint.Shape>;
  options: CommandOptions;
}

interface BaseCommand {
  id: string;
  section: SectionId;
  label: string;
  /** id en src/taskpane/icons.ts */
  icon: string;
  minShapes: number;
  maxShapes?: number;
  /** Si es true, el panel pide confirmación antes de ejecutar (p. ej. un borrado sobre todo el mazo). */
  confirm?: boolean;
}

export interface GeometryCommand extends BaseCommand {
  kind: "geometry";
  compute: (selection: SelectionContext, options: CommandOptions) => Patch[];
}

export interface FormatCommand extends BaseCommand {
  kind: "format";
  /** Si es true el botón se deshabilita mientras no haya formato copiado. */
  needsClipboard?: boolean;
  compute: (formats: ShapeFormat[], options: CommandOptions, clipboard: ShapeFormat | null) => FormatPatch[];
}

/** Guarda el formato de la primera forma seleccionada en el portapapeles interno. */
export interface CopyFormatCommand extends BaseCommand {
  kind: "copy-format";
}

export interface TextCommand extends BaseCommand {
  kind: "text";
  compute: (texts: TextShape[], options: CommandOptions) => TextPatch[];
}

/** Firma de `TextCommand.compute`; evita el tipo condicional al tipar el helper en `text.ts`. */
export type TextCompute = TextCommand["compute"];

export interface ActionCommand extends BaseCommand {
  kind: "action";
  run: (ctx: ActionContext) => Promise<CommandOutcome>;
}

export type Command = GeometryCommand | FormatCommand | CopyFormatCommand | TextCommand | ActionCommand;
