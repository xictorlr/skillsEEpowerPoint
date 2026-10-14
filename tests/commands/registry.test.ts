import { describe, expect, it } from "vitest";
import { SECTION_ORDER, commands, commandsBySection, disabledReason, findCommand } from "../../src/commands/registry";
import type { CommandOptions } from "../../src/commands/types";
import { shape, slide16x9 } from "../core/fixtures";

const OPTS: CommandOptions = {
  matrix: { columns: 2, gapX: 8, gapY: 8 },
  colorTarget: "fill",
  pasteParts: { fill: true, line: true, font: true, size: true },
  fontSizeStep: 1,
  nameFilter: "",
  renamePrefix: "Forma",
  shapeKind: "all",
};

describe("registry", () => {
  it("has unique ids and icons defined", () => {
    const ids = commands.map((c) => c.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(commands.every((c) => c.icon.length > 0 && c.label.length > 0)).toBe(true);
  });

  it("reports why a command is disabled", () => {
    const swap = findCommand("swap");
    expect(swap).toBeDefined();
    expect(disabledReason(swap!, 1)).toMatch(/al menos 2/);
    expect(disabledReason(swap!, 3)).toMatch(/como máximo 2/);
    expect(disabledReason(swap!, 2)).toBeUndefined();
  });

  it("align-left computes patches via core", () => {
    const cmd = findCommand("align-left")!;
    if (cmd.kind !== "geometry") throw new Error("expected a geometry command");
    const ref = shape({ id: "r", left: 10, top: 0, width: 10, height: 10 });
    const b = shape({ id: "b", left: 50, top: 0, width: 10, height: 10 });
    const patches = cmd.compute({ shapes: [ref, b], slide: slide16x9 }, OPTS);
    expect(patches).toEqual([{ id: "b", left: 10 }]);
  });

  it("includes format and color sections and hides set-color from buttons", () => {
    expect(SECTION_ORDER).toEqual(["alignment", "size", "format", "color", "text", "navigation", "expert"]);
    expect(commandsBySection("format").map((c) => c.id)).toEqual(["copy-format", "paste-format", "paste-fill", "paste-line", "paste-font", "paste-size"]);
    expect(commandsBySection("color").map((c) => c.id)).toEqual(["no-fill", "no-line", "transparency-up", "transparency-down"]);
    expect(findCommand("set-color")).toBeDefined();
  });

  it("hides select-kind, select-by-name and inspector-apply from buttons but keeps the rest of expert visible", () => {
    expect(commandsBySection("expert").map((c) => c.id)).toEqual([
      "rename-numbered",
      "z-front",
      "z-forward",
      "z-backward",
      "z-back",
      "group",
      "ungroup",
    ]);
    expect(findCommand("select-kind")).toBeDefined();
    expect(findCommand("select-by-name")).toBeDefined();
    expect(findCommand("inspector-apply")).toBeDefined();
  });

  it("includes navigation commands, hiding slide-goto (replaced by the 'Ir a' field)", () => {
    expect(commandsBySection("navigation").map((c) => c.id)).toEqual(["slide-first", "slide-prev", "slide-next", "slide-last"]);
    expect(findCommand("slide-goto")).toBeDefined();
  });

  it("disables paste commands without a clipboard", () => {
    const paste = findCommand("paste-format")!;
    expect(disabledReason(paste, 1, false)).toMatch(/copia el formato/);
    expect(disabledReason(paste, 1, true)).toBeUndefined();
    expect(disabledReason(findCommand("copy-format")!, 1, false)).toBeUndefined();
  });
});
