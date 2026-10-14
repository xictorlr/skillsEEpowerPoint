import { alignmentCommands } from "./alignment";
import { sizeCommands } from "./size";
import type { Command, SectionId } from "./types";

export const SECTION_ORDER: SectionId[] = ["alignment", "size"];

export const commands: Command[] = [...alignmentCommands, ...sizeCommands];

export function commandsBySection(section: SectionId): Command[] {
  return commands.filter((c) => c.section === section);
}

export function findCommand(id: string): Command | undefined {
  return commands.find((c) => c.id === id);
}

/** Devuelve el motivo por el que un comando no aplica a `count` formas, o undefined si aplica. */
export function disabledReason(cmd: Command, count: number): string | undefined {
  if (count < cmd.minShapes) return `Selecciona al menos ${cmd.minShapes} forma${cmd.minShapes > 1 ? "s" : ""}.`;
  if (cmd.maxShapes !== undefined && count > cmd.maxShapes) return `Selecciona como máximo ${cmd.maxShapes} formas.`;
  return undefined;
}
