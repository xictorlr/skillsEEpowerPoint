export type CommandResult = { ok: true; changed: number } | { ok: false; message: string };

/**
 * Mensajes de error legibles por código de error de Office.js.
 * Usa duck typing (en vez de `instanceof OfficeExtension.Error`) porque `OfficeExtension` no está
 * definido en todos los entornos donde puede correr este código (p. ej. en tests que no cargan
 * los tipos/globals de Office.js).
 */
function describeError(err: unknown): string {
  if (typeof err === "object" && err !== null && "code" in err) {
    const code = String((err as Record<string, unknown>).code);
    const rawMessage = (err as Record<string, unknown>).message;
    if (code === "ApiNotFound") return "Esta función requiere una versión más reciente de PowerPoint.";
    if (code === "InvalidOperation") return "Operación no válida para la selección actual.";
    if (code === "GeneralException")
      return "PowerPoint rechazó la operación. Puede que alguna forma esté bloqueada o no admita el cambio.";
    if (code === "ItemNotFound") return "No se encontró la forma o diapositiva seleccionada.";
    return `${code}: ${typeof rawMessage === "string" ? rawMessage : String(rawMessage)}`;
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
