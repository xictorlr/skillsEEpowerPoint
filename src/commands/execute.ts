import { applyPatches } from "../office/apply";
import { runCommand, type CommandResult } from "../office/runtime";
import { readSelection } from "../office/selection";
import type { SlideSize } from "../core/types";
import { disabledReason } from "./registry";
import type { Command, CommandOptions } from "./types";

const slideSizeCache = new Map<string, SlideSize>();

/** Lee la selección, calcula los parches y los aplica en un único sync final. */
export async function executeCommand(cmd: Command, options: CommandOptions): Promise<CommandResult> {
  return runCommand(cmd.id, async (context) => {
    const { selection, proxies } = await readSelection(context, { slideSizeCache });
    const reason = disabledReason(cmd, selection.shapes.length);
    if (reason) throw new Error(reason);
    const patches = cmd.compute(selection, options);
    if (patches.length === 0) return 0;
    applyPatches(proxies, patches);
    await context.sync();
    return new Set(patches.map((p) => p.id)).size;
  });
}
