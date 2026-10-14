import { describe, expect, it } from "vitest";
import { commands } from "../../src/commands/registry";
import { icons } from "../../src/taskpane/icons";

describe("icons", () => {
  it("every command has an icon", () => {
    const missing = commands.filter((c) => !(c.icon in icons)).map((c) => c.id);
    expect(missing).toEqual([]);
  });
});
