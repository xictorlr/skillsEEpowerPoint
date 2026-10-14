import { buildPastePatches } from "../core/format";
import type { Command } from "./types";

const paste = (id: string, label: string, icon: string, parts: { fill: boolean; line: boolean; font: boolean; size: boolean } | "options"): Command => ({
  kind: "format",
  id,
  section: "format",
  label,
  icon,
  minShapes: 1,
  needsClipboard: true,
  compute: (formats, options, clipboard) =>
    clipboard ? buildPastePatches(clipboard, formats, parts === "options" ? options.pasteParts : parts) : [],
});

export const formatCommands: readonly Command[] = [
  { kind: "copy-format", id: "copy-format", section: "format", label: "Copiar formato de la primera forma", icon: "copyFormat", minShapes: 1 },
  paste("paste-format", "Pegar formato (según partes elegidas)", "pasteFormat", "options"),
  paste("paste-fill", "Pegar solo relleno", "pasteFill", { fill: true, line: false, font: false, size: false }),
  paste("paste-line", "Pegar solo línea", "pasteLine", { fill: false, line: true, font: false, size: false }),
  paste("paste-font", "Pegar solo fuente", "pasteFont", { fill: false, line: false, font: true, size: false }),
  paste("paste-size", "Pegar solo tamaño", "pasteSize", { fill: false, line: false, font: false, size: true }),
];
