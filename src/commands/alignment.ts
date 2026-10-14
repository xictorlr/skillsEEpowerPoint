import { alignShapes, alignToSlide, arrangeMatrix, distribute, swapPositions } from "../core/geometry";
import type { Command } from "./types";

export const alignmentCommands: readonly Command[] = [
  { kind: "geometry", id: "align-left", section: "alignment", label: "Alinear a la izquierda", icon: "alignLeft", minShapes: 2, compute: (s) => alignShapes(s.shapes, "left") },
  { kind: "geometry", id: "align-center-h", section: "alignment", label: "Centrar horizontalmente", icon: "alignCenterH", minShapes: 2, compute: (s) => alignShapes(s.shapes, "centerH") },
  { kind: "geometry", id: "align-right", section: "alignment", label: "Alinear a la derecha", icon: "alignRight", minShapes: 2, compute: (s) => alignShapes(s.shapes, "right") },
  { kind: "geometry", id: "align-top", section: "alignment", label: "Alinear arriba", icon: "alignTop", minShapes: 2, compute: (s) => alignShapes(s.shapes, "top") },
  { kind: "geometry", id: "align-center-v", section: "alignment", label: "Centrar verticalmente", icon: "alignCenterV", minShapes: 2, compute: (s) => alignShapes(s.shapes, "centerV") },
  { kind: "geometry", id: "align-bottom", section: "alignment", label: "Alinear abajo", icon: "alignBottom", minShapes: 2, compute: (s) => alignShapes(s.shapes, "bottom") },
  { kind: "geometry", id: "distribute-h", section: "alignment", label: "Distribuir horizontalmente", icon: "distributeH", minShapes: 3, compute: (s) => distribute(s.shapes, "h") },
  { kind: "geometry", id: "distribute-v", section: "alignment", label: "Distribuir verticalmente", icon: "distributeV", minShapes: 3, compute: (s) => distribute(s.shapes, "v") },
  { kind: "geometry", id: "swap", section: "alignment", label: "Intercambiar posiciones", icon: "swap", minShapes: 2, maxShapes: 2, compute: (s) => swapPositions(s.shapes) },
  { kind: "geometry", id: "matrix", section: "alignment", label: "Alinear en matriz", icon: "matrix", minShapes: 2, compute: (s, o) => arrangeMatrix(s.shapes, o.matrix) },
  { kind: "geometry", id: "slide-center-h", section: "alignment", label: "Centrar en la diapositiva (horizontal)", icon: "slideCenterH", minShapes: 1, compute: (s) => alignToSlide(s.shapes, s.slide, "centerH") },
  { kind: "geometry", id: "slide-center-v", section: "alignment", label: "Centrar en la diapositiva (vertical)", icon: "slideCenterV", minShapes: 1, compute: (s) => alignToSlide(s.shapes, s.slide, "centerV") },
  { kind: "geometry", id: "slide-center", section: "alignment", label: "Centrar en la diapositiva", icon: "slideCenter", minShapes: 1, compute: (s) => [...alignToSlide(s.shapes, s.slide, "centerH"), ...alignToSlide(s.shapes, s.slide, "centerV")] },
];
