import type { MatrixOptions } from "../core/geometry";
import type { ColorTarget, FormatParts, FormatPatch, ShapeFormat } from "../core/format";
import type { Patch, SelectionContext, SlideSize } from "../core/types";
import type { TextPatch, TextShape } from "../core/text";

export type SectionId = "alignment" | "size" | "format" | "color" | "text";

export const SECTION_LABELS: Readonly<Record<SectionId, string>> = {
  alignment: "Alignment",
  size: "Size",
  format: "Format",
  color: "Color",
  text: "Text",
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
}

interface BaseCommand {
  id: string;
  section: SectionId;
  label: string;
  /** id en src/taskpane/icons.ts */
  icon: string;
  minShapes: number;
  maxShapes?: number;
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

export type Command = GeometryCommand | FormatCommand | CopyFormatCommand | TextCommand;
