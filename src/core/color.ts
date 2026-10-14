/** Normaliza a "#RRGGBB" en mayúsculas. Acepta "abc", "#abc", "aabbcc", "#AABBCC". */
export function normalizeHex(input: string): string | null {
  const s = input.trim().replace(/^#/, "");
  if (/^[0-9a-fA-F]{6}$/.test(s)) return `#${s.toUpperCase()}`;
  if (/^[0-9a-fA-F]{3}$/.test(s)) {
    const [r, g, b] = s.toUpperCase();
    return `#${r}${r}${g}${g}${b}${b}`;
  }
  return null;
}

/** Paleta neutra: grises, azules, verdes, ámbar/naranja, rojo, morado. Sin colores de marca. */
export const DEFAULT_PALETTE: readonly string[] = [
  "#000000", "#404040", "#808080", "#BFBFBF", "#D9D9D9", "#FFFFFF",
  "#0B3D91", "#1F6FEB", "#7FB3FF",
  "#0F6E56", "#2EA043", "#9BE9A8",
  "#B35900", "#F2A900",
  "#B42318", "#7A1FA2",
];

/** Inserta `color` al principio de `recent` sin duplicados, limitado a `max` entradas.
 *  Las entradas ya guardadas también se normalizan (y las inválidas se descartan) antes de deduplicar. */
export function pushRecent(recent: readonly string[], color: string, max: number): string[] {
  const normalizedRecent = recent.map((c) => normalizeHex(c)).filter((c): c is string => c !== null);
  const hex = normalizeHex(color);
  if (!hex) return normalizedRecent;
  return [hex, ...normalizedRecent.filter((c) => c !== hex)].slice(0, max);
}
