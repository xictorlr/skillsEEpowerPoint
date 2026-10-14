import { autoSizePatches, casePatches, cleanPatches, fontSizePatches, marginPatches, paragraphAlignPatches, verticalAlignPatches, wordWrapPatches } from "../core/text";
import type { Command, TextCompute } from "./types";

const t = (id: string, label: string, icon: string, compute: TextCompute): Command => ({
  kind: "text",
  id,
  section: "text",
  label,
  icon,
  minShapes: 1,
  compute,
});

export const textCommands: readonly Command[] = [
  t("case-upper", "MAYÚSCULAS", "caseUpper", (s) => casePatches(s, "upper")),
  t("case-lower", "minúsculas", "caseLower", (s) => casePatches(s, "lower")),
  t("case-sentence", "Tipo oración", "caseSentence", (s) => casePatches(s, "sentence")),
  t("case-title", "Tipo Título", "caseTitle", (s) => casePatches(s, "title")),
  t("font-bigger", "Aumentar tamaño de fuente", "fontBigger", (s, o) => fontSizePatches(s, o.fontSizeStep)),
  t("font-smaller", "Reducir tamaño de fuente", "fontSmaller", (s, o) => fontSizePatches(s, -o.fontSizeStep)),
  t("para-left", "Alinear párrafo a la izquierda", "paraLeft", (s) => paragraphAlignPatches(s, "Left")),
  t("para-center", "Centrar párrafo", "paraCenter", (s) => paragraphAlignPatches(s, "Center")),
  t("para-right", "Alinear párrafo a la derecha", "paraRight", (s) => paragraphAlignPatches(s, "Right")),
  t("para-justify", "Justificar párrafo", "paraJustify", (s) => paragraphAlignPatches(s, "Justify")),
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
