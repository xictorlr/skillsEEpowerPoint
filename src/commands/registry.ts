import { alignmentCommands } from "./alignment";
import { sizeCommands } from "./size";
import type { Command, SectionId } from "./types";

export const SECTION_ORDER: readonly SectionId[] = ["alignment", "size"];

export const commands: readonly Command[] = [...alignmentCommands, ...sizeCommands];

export function commandsBySection(section: SectionId): Command[] {
  return commands.filter((c) => c.section === section);
}

export function findCommand(id: string): Command | undefined {
  return commands.find((c) => c.id === id);
}

/** Pluraliza "forma(s)" para los mensajes de disabledReason. */
function formas(n: number): string {
  return `${n} forma${n > 1 ? "s" : ""}`;
}

/** Devuelve el motivo por el que un comando no aplica a `count` formas, o undefined si aplica. */
export function disabledReason(cmd: Command, count: number): string | undefined {
  if (count < cmd.minShapes) return `Selecciona al menos ${formas(cmd.minShapes)}.`;
  if (cmd.maxShapes !== undefined && count > cmd.maxShapes) return `Selecciona como máximo ${formas(cmd.maxShapes)}.`;
  return undefined;
}
