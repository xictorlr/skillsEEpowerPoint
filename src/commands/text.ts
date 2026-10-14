import {
  autoSizePatches,
  casePatches,
  cleanPatches,
  fontSizePatches,
  fontSizePresetPatches,
  indentPatches,
  marginPatches,
  paragraphAlignPatches,
  toggleBoldPatches,
  toggleBulletsPatches,
  toggleItalicPatches,
  toggleStrikethroughPatches,
  toggleUnderlinePatches,
  verticalAlignPatches,
  wordWrapPatches,
} from "../core/text";
import type { Command, TextCompute } from "./types";

/** `extra` permite añadir campos opcionales de `BaseCommand` (p. ej. `requiresApi` en la sangría). */
const t = (id: string, label: string, icon: string, compute: TextCompute, extra: { requiresApi?: string } = {}): Command => ({
  kind: "text",
  id,
  section: "text",
  label,
  icon,
  minShapes: 1,
  compute,
  ...extra,
});

/** La sangría (`paragraphFormat.indentLevel`) solo se puede leer/escribir con PowerPointApi 1.10. */
const INDENT_API = { requiresApi: "1.10" };

export const textCommands: readonly Command[] = [
  t("case-upper", "MAYÚSCULAS", "caseUpper", (s) => casePatches(s, "upper")),
  t("case-lower", "minúsculas", "caseLower", (s) => casePatches(s, "lower")),
  t("case-sentence", "Tipo oración", "caseSentence", (s) => casePatches(s, "sentence")),
  t("case-title", "Tipo Título", "caseTitle", (s) => casePatches(s, "title")),
  t("style-bold", "Negrita (conmutar)", "styleBold", (s) => toggleBoldPatches(s)),
  t("style-italic", "Cursiva (conmutar)", "styleItalic", (s) => toggleItalicPatches(s)),
  t("style-underline", "Subrayado (conmutar)", "styleUnderline", (s) => toggleUnderlinePatches(s)),
  t("style-strike", "Tachado (conmutar)", "styleStrike", (s) => toggleStrikethroughPatches(s)),
  t("font-bigger", "Aumentar tamaño de fuente", "fontBigger", (s, o) => fontSizePatches(s, o.fontSizeStep)),
  t("font-smaller", "Reducir tamaño de fuente", "fontSmaller", (s, o) => fontSizePatches(s, -o.fontSizeStep)),
  t("size-title", "Tamaño título (24 pt)", "sizeTitle", (s) => fontSizePresetPatches(s, 24)),
  t("size-subtitle", "Tamaño subtítulo (18 pt)", "sizeSubtitle", (s) => fontSizePresetPatches(s, 18)),
  t("size-body", "Tamaño cuerpo (14 pt)", "sizeBody", (s) => fontSizePresetPatches(s, 14)),
  t("size-small", "Tamaño pequeño (11 pt)", "sizeSmall", (s) => fontSizePresetPatches(s, 11)),
  t("para-left", "Alinear párrafo a la izquierda", "paraLeft", (s) => paragraphAlignPatches(s, "Left")),
  t("para-center", "Centrar párrafo", "paraCenter", (s) => paragraphAlignPatches(s, "Center")),
  t("para-right", "Alinear párrafo a la derecha", "paraRight", (s) => paragraphAlignPatches(s, "Right")),
  t("para-justify", "Justificar párrafo", "paraJustify", (s) => paragraphAlignPatches(s, "Justify")),
  t("bullets-toggle", "Viñetas (conmutar)", "bullets", (s) => toggleBulletsPatches(s)),
  t("indent-more", "Aumentar sangría", "indentMore", (s) => indentPatches(s, 1), INDENT_API),
  t("indent-less", "Reducir sangría", "indentLess", (s) => indentPatches(s, -1), INDENT_API),
  t("valign-top", "Texto arriba", "valignTop", (s) => verticalAlignPatches(s, "Top")),
  t("valign-middle", "Texto en el medio", "valignMiddle", (s) => verticalAlignPatches(s, "Middle")),
  t("valign-bottom", "Texto abajo", "valignBottom", (s) => verticalAlignPatches(s, "Bottom")),
  t("margins-zero", "Márgenes internos a cero", "marginsZero", (s) => marginPatches(s, "zero")),
  t("margins-standard", "Márgenes internos estándar", "marginsStandard", (s) => marginPatches(s, "standard")),
  t("autosize-none", "Sin autoajuste", "autosizeNone", (s) => autoSizePatches(s, "AutoSizeNone")),
  t("autosize-shrink", "Reducir texto al desbordar", "autosizeShrink", (s) => autoSizePatches(s, "AutoSizeTextToFitShape")),
  t("autosize-shape", "Ajustar forma al texto", "autosizeShape", (s) => autoSizePatches(s, "AutoSizeShapeToFitText")),
  t("wrap-on", "Ajustar texto a la forma", "wrapOn", (s) => wordWrapPatches(s, true)),
  t("wrap-off", "No ajustar texto", "wrapOff", (s) => wordWrapPatches(s, false)),
  t("clean-text", "Limpiar texto (espacios y líneas vacías)", "cleanText", (s) => cleanPatches(s)),
];
