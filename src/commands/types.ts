import type { Patch, SelectionContext } from "../core/types";
import type { MatrixOptions } from "../core/geometry";

export type SectionId = "alignment" | "size";

export const SECTION_LABELS: Record<SectionId, string> = {
  alignment: "Alignment",
  size: "Size",
};

/** Opciones de usuario que algunos comandos necesitan (vienen del panel). */
export interface CommandOptions {
  matrix: MatrixOptions;
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
