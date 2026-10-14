import type { Patch, SelectionContext, SlideSize } from "../core/types";
import type { MatrixOptions } from "../core/geometry";

export type SectionId = "alignment" | "size";

export const SECTION_LABELS: Readonly<Record<SectionId, string>> = {
  alignment: "Alignment",
  size: "Size",
};

/** Opciones de usuario que algunos comandos necesitan (vienen del panel). */
export interface CommandOptions {
  matrix: MatrixOptions;
  /** Tamaño de diapositiva forzado por el usuario; si falta se deduce automáticamente. */
  slideSize?: SlideSize;
}

export interface GeometryCommand {
  kind: "geometry";
  id: string;
  section: SectionId;
  label: string;
  /** id en src/taskpane/icons.ts */
  icon: string;
  minShapes: number;
  maxShapes?: number;
  compute: (selection: SelectionContext, options: CommandOptions) => Patch[];
}

export type Command = GeometryCommand;
