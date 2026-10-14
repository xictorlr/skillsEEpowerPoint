import { applyPatches } from "../office/apply";
import {
  applyZOrder,
  groupSelected,
  readSlideNav,
  readSlideShapes,
  renameShapes,
  selectShapes,
  ungroupShapes,
} from "../office/navigation";
import { applyBoundingBox, boundingBox, filterByKind, filterByName, numberedNames, zOrderApplyOrder, type ZOrderOp } from "../core/expert";
import type { ActionCommand } from "./types";

/** `select-kind`: selecciona las formas de la diapositiva activa que coinciden con `options.shapeKind`. */
const selectKind: ActionCommand = {
  kind: "action",
  id: "select-kind",
  section: "expert",
  label: "Seleccionar por tipo",
  icon: "selectKind",
  minShapes: 0,
  run: async ({ context, options }) => {
    const nav = await readSlideNav(context);
    const shapes = await readSlideShapes(context, nav.slideId);
    const filtered = filterByKind(shapes, options.shapeKind);
    selectShapes(
      context,
      nav.slideId,
      filtered.map((s) => s.id),
    );
    await context.sync();
    return { changed: filtered.length, message: filtered.length === 0 ? "No hay formas de ese tipo." : undefined };
  },
};

/** `select-by-name`: selecciona las formas cuyo nombre contiene `options.nameFilter`. */
const selectByName: ActionCommand = {
  kind: "action",
  id: "select-by-name",
  section: "expert",
  label: "Seleccionar por nombre",
  icon: "selectByName",
  minShapes: 0,
  run: async ({ context, options }) => {
    if (options.nameFilter.trim() === "") return { changed: 0, message: "Escribe parte del nombre." };
    const nav = await readSlideNav(context);
    const shapes = await readSlideShapes(context, nav.slideId);
    const filtered = filterByName(shapes, options.nameFilter);
    selectShapes(
      context,
      nav.slideId,
      filtered.map((s) => s.id),
    );
    await context.sync();
    return { changed: filtered.length, message: filtered.length === 0 ? "No hay formas de ese tipo." : undefined };
  },
};

/** `rename-numbered`: renombra la selección en orden de lectura con `options.renamePrefix`. */
const renameNumbered: ActionCommand = {
  kind: "action",
  id: "rename-numbered",
  section: "expert",
  label: "Renombrar numerado",
  icon: "renameNumbered",
  minShapes: 1,
  run: async ({ context, selection, proxies, options }) => {
    const names = numberedNames(selection.shapes, options.renamePrefix || "Forma");
    renameShapes(proxies, names);
    await context.sync();
    return { changed: names.length };
  },
};

function zOrderCommand(id: string, label: string, icon: string, op: ZOrderOp): ActionCommand {
  return {
    kind: "action",
    id,
    section: "expert",
    label,
    icon,
    minShapes: 1,
    run: async ({ context, selection, proxies }) => {
      const ordered = zOrderApplyOrder(selection.shapes, op);
      applyZOrder(
        proxies,
        ordered.map((s) => s.id),
        op,
      );
      await context.sync();
      return { changed: ordered.length };
    },
  };
}

/** `group`: agrupa la selección actual (requiere 2+ formas). */
const group: ActionCommand = {
  kind: "action",
  id: "group",
  section: "expert",
  label: "Agrupar",
  icon: "group",
  minShapes: 2,
  run: async ({ context }) => {
    groupSelected(context);
    await context.sync();
    return { changed: 1, message: "Agrupado." };
  },
};

/** `ungroup`: desagrupa los grupos de la selección actual, si los hay. */
const ungroup: ActionCommand = {
  kind: "action",
  id: "ungroup",
  section: "expert",
  label: "Desagrupar",
  icon: "ungroup",
  minShapes: 1,
  run: async ({ context, selection, proxies }) => {
    const groupIds = selection.shapes.filter((s) => s.type === "Group").map((s) => s.id);
    if (groupIds.length === 0) return { changed: 0, message: "No hay grupos en la selección." };
    ungroupShapes(proxies, groupIds);
    await context.sync();
    return { changed: groupIds.length };
  },
};

/** `inspector-apply`: mueve/escala la selección para que su caja envolvente pase a `options.inspector`. */
const inspectorApply: ActionCommand = {
  kind: "action",
  id: "inspector-apply",
  section: "expert",
  label: "Aplicar inspector",
  icon: "inspector",
  minShapes: 1,
  run: async ({ context, selection, proxies, options }) => {
    const bbox = boundingBox(selection.shapes);
    if (!bbox) return { changed: 0 };
    const target = { ...bbox, ...options.inspector };
    const patches = applyBoundingBox(selection.shapes, target);
    if (patches.length === 0) return { changed: 0 };
    applyPatches(proxies, patches);
    await context.sync();
    return { changed: patches.length };
  },
};

export const expertCommands: readonly ActionCommand[] = [
  selectKind,
  selectByName,
  renameNumbered,
  zOrderCommand("z-front", "Traer al frente", "zFront", "BringToFront"),
  zOrderCommand("z-forward", "Traer adelante", "zForward", "BringForward"),
  zOrderCommand("z-backward", "Enviar atrás", "zBackward", "SendBackward"),
  zOrderCommand("z-back", "Enviar al fondo", "zBack", "SendToBack"),
  group,
  ungroup,
  inspectorApply,
];
