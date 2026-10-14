import { matchSize, stretchToEdge } from "../core/geometry";
import type { Command } from "./types";

export const sizeCommands: readonly Command[] = [
  { kind: "geometry", id: "match-width", section: "size", label: "Igualar ancho al primero", icon: "matchWidth", minShapes: 2, compute: (s) => matchSize(s.shapes, "width", "first") },
  { kind: "geometry", id: "match-height", section: "size", label: "Igualar alto al primero", icon: "matchHeight", minShapes: 2, compute: (s) => matchSize(s.shapes, "height", "first") },
  { kind: "geometry", id: "match-both", section: "size", label: "Igualar tamaño al primero", icon: "matchBoth", minShapes: 2, compute: (s) => matchSize(s.shapes, "both", "first") },
  { kind: "geometry", id: "match-widest", section: "size", label: "Igualar al más ancho", icon: "matchWidest", minShapes: 2, compute: (s) => matchSize(s.shapes, "width", "max") },
  { kind: "geometry", id: "match-tallest", section: "size", label: "Igualar al más alto", icon: "matchTallest", minShapes: 2, compute: (s) => matchSize(s.shapes, "height", "max") },
  { kind: "geometry", id: "match-narrowest", section: "size", label: "Igualar al más estrecho", icon: "matchNarrowest", minShapes: 2, compute: (s) => matchSize(s.shapes, "width", "min") },
  { kind: "geometry", id: "match-shortest", section: "size", label: "Igualar al más bajo", icon: "matchShortest", minShapes: 2, compute: (s) => matchSize(s.shapes, "height", "min") },
  { kind: "geometry", id: "stretch-left", section: "size", label: "Estirar hasta el borde izquierdo", icon: "stretchLeft", minShapes: 2, compute: (s) => stretchToEdge(s.shapes, "left") },
  { kind: "geometry", id: "stretch-right", section: "size", label: "Estirar hasta el borde derecho", icon: "stretchRight", minShapes: 2, compute: (s) => stretchToEdge(s.shapes, "right") },
  { kind: "geometry", id: "stretch-top", section: "size", label: "Estirar hasta el borde superior", icon: "stretchTop", minShapes: 2, compute: (s) => stretchToEdge(s.shapes, "top") },
  { kind: "geometry", id: "stretch-bottom", section: "size", label: "Estirar hasta el borde inferior", icon: "stretchBottom", minShapes: 2, compute: (s) => stretchToEdge(s.shapes, "bottom") },
];
