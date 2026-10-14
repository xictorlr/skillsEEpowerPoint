import { alignmentCommands } from "./alignment";
import { colorCommands } from "./color";
import { formatCommands } from "./format";
import { sizeCommands } from "./size";
import { textCommands } from "./text";
import type { Command, SectionId } from "./types";

export const SECTION_ORDER: readonly SectionId[] = ["alignment", "size", "format", "color", "text"];

export const commands: readonly Command[] = [...alignmentCommands, ...sizeCommands, ...formatCommands, ...colorCommands, ...textCommands];

/** Comandos que el panel no renderiza como botón (se disparan desde otros controles). */
export const hiddenCommandIds: ReadonlySet<string> = new Set(["set-color"]);

export function commandsBySection(section: SectionId): Command[] {
  return commands.filter((c) => c.section === section && !hiddenCommandIds.has(c.id));
}

export function findCommand(id: string): Command | undefined {
  return commands.find((c) => c.id === id);
}

function formas(n: number): string {
  return `${n} forma${n === 1 ? "" : "s"}`;
}

/** Motivo por el que un comando no aplica, o undefined si aplica. `hasClipboard` solo importa para pegar formato. */
export function disabledReason(cmd: Command, count: number, hasClipboard = true): string | undefined {
  if (count < cmd.minShapes) return `Selecciona al menos ${formas(cmd.minShapes)}.`;
  if (cmd.maxShapes !== undefined && count > cmd.maxShapes) return `Selecciona como máximo ${formas(cmd.maxShapes)}.`;
  if (cmd.kind === "format" && cmd.needsClipboard && !hasClipboard) return "Primero copia el formato de una forma.";
  return undefined;
}
