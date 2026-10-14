import { afterEach, describe, expect, it, vi } from "vitest";
import { getThemeColors, THEME_COLOR_NAMES } from "../../src/office/theme";
import { MockContext } from "./mockPowerPoint";

const THEME_COLORS: Record<string, string> = {
  Dark1: "000000",
  Light1: "#FFFFFF",
  Dark2: "#1F1F1F",
  Light2: "#E7E6E6",
  Accent1: "#4472C4",
  Accent2: "#ED7D31",
  Accent3: "#A5A5A5",
  Accent4: "#FFC000",
  Accent5: "#5B9BD5",
  Accent6: "#70AD47",
  Hyperlink: "#0563C1",
  FollowedHyperlink: "#954F72",
};

function stubOffice1_10(): void {
  vi.stubGlobal("Office", {
    context: {
      requirements: {
        isSetSupported: (name: string, version: string) => name === "PowerPointApi" && version === "1.10",
      },
    },
  });
}

describe("getThemeColors", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("returns null when PowerPointApi 1.10 isn't supported", async () => {
    const ctx = new MockContext([], "", { themeColors: THEME_COLORS });
    const result = await getThemeColors(ctx as never);
    expect(result).toBeNull();
  });

  it("returns the 12 theme colors, normalized to #RRGGBB, with a single sync", async () => {
    stubOffice1_10();
    const ctx = new MockContext([], "", { themeColors: THEME_COLORS });
    const result = await getThemeColors(ctx as never);
    expect(result).toHaveLength(THEME_COLOR_NAMES.length);
    expect(result).toEqual([
      { name: "Dark1", hex: "#000000" },
      { name: "Light1", hex: "#FFFFFF" },
      { name: "Dark2", hex: "#1F1F1F" },
      { name: "Light2", hex: "#E7E6E6" },
      { name: "Accent1", hex: "#4472C4" },
      { name: "Accent2", hex: "#ED7D31" },
      { name: "Accent3", hex: "#A5A5A5" },
      { name: "Accent4", hex: "#FFC000" },
      { name: "Accent5", hex: "#5B9BD5" },
      { name: "Accent6", hex: "#70AD47" },
      { name: "Hyperlink", hex: "#0563C1" },
      { name: "FollowedHyperlink", hex: "#954F72" },
    ]);
    expect(ctx.syncCount).toBe(1);
  });

  it("drops colors whose value can't be normalized to a hex color", async () => {
    stubOffice1_10();
    const ctx = new MockContext([], "", { themeColors: { ...THEME_COLORS, Accent1: "not-a-color" } });
    const result = await getThemeColors(ctx as never);
    expect(result?.find((c) => c.name === "Accent1")).toBeUndefined();
    expect(result).toHaveLength(THEME_COLOR_NAMES.length - 1);
  });

  it("drops only the color whose value is undefined, keeping the other 11", async () => {
    stubOffice1_10();
    const { Accent3: _omitted, ...partial } = THEME_COLORS;
    const ctx = new MockContext([], "", { themeColors: partial });
    const result = await getThemeColors(ctx as never);
    expect(result?.find((c) => c.name === "Accent3")).toBeUndefined();
    expect(result).toHaveLength(THEME_COLOR_NAMES.length - 1);
  });

  it("returns null if reading the theme colors fails (e.g. no slide selected)", async () => {
    stubOffice1_10();
    const ctx = new MockContext([], "", { themeColors: THEME_COLORS });
    ctx.invalidTouched = true; // fuerza que el próximo sync() lance
    const result = await getThemeColors(ctx as never);
    expect(result).toBeNull();
  });
});
