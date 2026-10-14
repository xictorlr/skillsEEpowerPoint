export type CommandResult = { ok: true; changed: number } | { ok: false; message: string };

/** Mensajes de error legibles por código de OfficeExtension. */
function describeError(err: unknown): string {
  if (err instanceof OfficeExtension.Error) {
    if (err.code === "ApiNotFound") return "Esta función requiere una versión más reciente de PowerPoint.";
    if (err.code === "InvalidOperation") return "Operación no válida para la selección actual.";
    return `${err.code}: ${err.message}`;
  }
  return err instanceof Error ? err.message : String(err);
}

/** Ejecuta `fn` dentro de PowerPoint.run y normaliza el resultado. `fn` devuelve cuántas formas cambió. */
export async function runCommand(
  commandId: string,
  fn: (context: PowerPoint.RequestContext) => Promise<number>,
): Promise<CommandResult> {
  try {
    const changed = await PowerPoint.run(fn);
    return { ok: true, changed };
  } catch (err) {
    const message = describeError(err);
    console.error(`[slide-tools] ${commandId} falló:`, err);
    return { ok: false, message };
  }
}
