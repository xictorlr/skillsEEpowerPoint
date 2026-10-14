import { clearPatches, colorPatches, transparencyPatches } from "../core/format";
import type { Command } from "./types";

export const TRANSPARENCY_STEP = 0.1;

export const colorCommands: readonly Command[] = [
  {
    kind: "format", id: "set-color", section: "color", label: "Aplicar color", icon: "setColor", minShapes: 1,
    compute: (formats, options) => (options.color ? colorPatches(formats, options.colorTarget, options.color) : []),
  },
  { kind: "format", id: "no-fill", section: "color", label: "Sin relleno", icon: "noFill", minShapes: 1, compute: (f) => clearPatches(f, "fill") },
  { kind: "format", id: "no-line", section: "color", label: "Sin línea", icon: "noLine", minShapes: 1, compute: (f) => clearPatches(f, "line") },
  {
    kind: "format", id: "transparency-up", section: "color", label: "Más transparente (+10 %)", icon: "transparencyUp", minShapes: 1,
    compute: (f, o) => (o.colorTarget === "font" ? [] : transparencyPatches(f, o.colorTarget, TRANSPARENCY_STEP)),
  },
  {
    kind: "format", id: "transparency-down", section: "color", label: "Menos transparente (−10 %)", icon: "transparencyDown", minShapes: 1,
    compute: (f, o) => (o.colorTarget === "font" ? [] : transparencyPatches(f, o.colorTarget, -TRANSPARENCY_STEP)),
  },
];
