import { applyPatches } from "../office/apply";
import { applyFormatPatch } from "../office/format";
import { runCommand, type CommandOutcome, type CommandResult } from "../office/runtime";
import { readSelection } from "../office/selection";
import type { SlideSize } from "../core/types";
import { getClipboard, setClipboard } from "./clipboard";
import { disabledReason } from "./registry";
import type { Command, CommandOptions } from "./types";

// La caché no se invalida si el usuario cambia Diseño > Tamaño de diapositiva con el panel abierto
// (limitación conocida; la opción slideSize es la alternativa).
const slideSizeCache = new Map<string, SlideSize>();

/** Solo para tests. */
export function resetSlideSizeCache(): void {
  slideSizeCache.clear();
}

/** Lee la selección, calcula los parches del comando y los aplica en un único sync de escritura. */
export async function executeCommand(cmd: Command, options: CommandOptions): Promise<CommandResult> {
  return runCommand(cmd.id, async (context): Promise<CommandOutcome> => {
    const withFormat = cmd.kind !== "geometry";
    const { selection, proxies, formats, caps } = await readSelection(context, {
      slideSizeCache,
      slideSizeOverride: options.slideSize,
      withFormat,
      // Los comandos de formato no leen selection.slide: evita pagar su sync.
      withSlideSize: !withFormat,
    });
    const reason = disabledReason(cmd, selection.shapes.length, getClipboard() !== null);
    if (reason) throw new Error(reason);

    switch (cmd.kind) {
      case "geometry": {
        const patches = cmd.compute(selection, options);
        if (patches.length === 0) return { changed: 0 };
        applyPatches(proxies, patches);
        await context.sync();
        return { changed: new Set(patches.map((p) => p.id)).size };
      }
      case "copy-format": {
        const first = formats[0];
        if (!first) return { changed: 0 };
        setClipboard(first);
        const message = first.font === null ? "Formato copiado (sin fuente: la forma no tiene texto)." : "Formato copiado.";
        return { changed: 0, message };
      }
      case "format": {
        const patches = cmd.compute(formats, options, getClipboard());
        if (patches.length === 0) return { changed: 0 };
        for (const p of patches) {
          const shape = proxies.get(p.id);
          const shapeCaps = caps.get(p.id);
          if (shape && shapeCaps) applyFormatPatch(shape, p, shapeCaps);
        }
        await context.sync();
        return { changed: new Set(patches.map((p) => p.id)).size };
      }
    }
  });
}
