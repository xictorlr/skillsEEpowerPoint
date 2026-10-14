/** Consulta soporte de un requirement set de PowerPointApi. Devuelve false si Office no está disponible. */
export function supportsApi(version: string): boolean {
  try {
    return typeof Office !== "undefined" && Office.context?.requirements?.isSetSupported("PowerPointApi", version) === true;
  } catch {
    return false;
  }
}
