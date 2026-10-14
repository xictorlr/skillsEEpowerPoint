import { applyPatches } from "../office/apply";
import { runCommand, type CommandResult } from "../office/runtime";
import { readSelection } from "../office/selection";
import type { SlideSize } from "../core/types";
import { disabledReason } from "./registry";
import type { Command, CommandOptions } from "./types";

// Cacheado por toda la sesión del panel: el tamaño de diapositiva no cambia una vez deducido.
// Limitación conocida: esta caché nunca se invalida si el usuario cambia Diseño > Tamaño de
// diapositiva en medio de la sesión; `options.slideSize` es el workaround (bypassa la caché).
const slideSizeCache = new Map<string, SlideSize>();

/** Vacía la caché de tamaño de diapositiva. Seam de test: los tests deben resetearla en cada caso. */
export function resetSlideSizeCache(): void {
  slideSizeCache.clear();
}

/** Lee la selección, calcula los parches y los aplica en un único sync final. */
export async function executeCommand(cmd: Command, options: CommandOptions): Promise<CommandResult> {
  return runCommand(cmd.id, async (context) => {
    const { selection, proxies } = await readSelection(context, {
      slideSizeCache,
      slideSizeOverride: options.slideSize,
    });
    const reason = disabledReason(cmd, selection.shapes.length);
    if (reason) throw new Error(reason);
    const patches = cmd.compute(selection, options);
    if (patches.length === 0) return { changed: 0 };
    applyPatches(proxies, patches);
    await context.sync();
    return { changed: new Set(patches.map((p) => p.id)).size };
  });
}
