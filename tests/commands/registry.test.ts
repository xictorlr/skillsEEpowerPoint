import { describe, expect, it } from "vitest";
import { commands, disabledReason, findCommand } from "../../src/commands/registry";
import { shape, slide16x9 } from "../core/fixtures";

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
    const ref = shape({ id: "r", left: 10, top: 0, width: 10, height: 10 });
    const b = shape({ id: "b", left: 50, top: 0, width: 10, height: 10 });
    const patches = cmd.compute({ shapes: [ref, b], slide: slide16x9 }, { matrix: { columns: 2, gapX: 8, gapY: 8 } });
    expect(patches).toEqual([{ id: "b", left: 10 }]);
  });
});
