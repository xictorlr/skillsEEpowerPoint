import { describe, expect, it } from "vitest";
import {
  DEFAULT_MATRIX,
  DEFAULT_PERSISTED,
  MAX_RECENT_COLORS,
  SLIDE_SIZE_PRESETS,
  sanitizeColorTarget,
  sanitizeFontSizeStep,
  sanitizeMatrix,
  sanitizePasteParts,
  sanitizeRecentColors,
  sanitizeSlideSizePreset,
  toCommandOptions,
} from "../../src/taskpane/sanitize";

describe("sanitizeMatrix", () => {
  it("falls back to defaults for undefined", () => {
    expect(sanitizeMatrix(undefined)).toEqual({ columns: 3, gapX: 8, gapY: 8 });
  });

  it("falls back to defaults for null", () => {
    expect(sanitizeMatrix(null)).toEqual({ columns: 3, gapX: 8, gapY: 8 });
  });

  it("falls back to defaults for a non-object value", () => {
    expect(sanitizeMatrix("not-an-object")).toEqual({ columns: 3, gapX: 8, gapY: 8 });
  });

  it("falls back to defaults for string fields", () => {
    expect(sanitizeMatrix({ columns: "4", gapX: "10", gapY: "12" })).toEqual({ columns: 3, gapX: 8, gapY: 8 });
  });

  it("falls back to defaults for NaN fields", () => {
    expect(sanitizeMatrix({ columns: NaN, gapX: NaN, gapY: NaN })).toEqual({ columns: 3, gapX: 8, gapY: 8 });
  });

  it("falls back to defaults for negative fields", () => {
    expect(sanitizeMatrix({ columns: -2, gapX: -5, gapY: -5 })).toEqual({ columns: 3, gapX: 8, gapY: 8 });
  });

  it("rounds columns with Math.round (2.5 -> 3)", () => {
    expect(sanitizeMatrix({ columns: 2.5, gapX: 8, gapY: 8 }).columns).toBe(3);
  });

  it("passes through valid numeric values", () => {
    expect(sanitizeMatrix({ columns: 4, gapX: 10, gapY: 12 })).toEqual({ columns: 4, gapX: 10, gapY: 12 });
  });

  it("allows zero for gaps but not for columns", () => {
    expect(sanitizeMatrix({ columns: 0, gapX: 0, gapY: 0 })).toEqual({ columns: 3, gapX: 0, gapY: 0 });
  });

  it("sanitizes each field independently when only some are present", () => {
    expect(sanitizeMatrix({ columns: 5 })).toEqual({ columns: 5, gapX: 8, gapY: 8 });
  });
});

describe("sanitizeSlideSizePreset", () => {
  it("returns 'auto' for garbage input", () => {
    expect(sanitizeSlideSizePreset(undefined)).toBe("auto");
    expect(sanitizeSlideSizePreset(null)).toBe("auto");
    expect(sanitizeSlideSizePreset(42)).toBe("auto");
    expect(sanitizeSlideSizePreset({})).toBe("auto");
    expect(sanitizeSlideSizePreset("not-a-real-preset")).toBe("auto");
  });

  it("returns the key unchanged for every valid preset", () => {
    for (const key of Object.keys(SLIDE_SIZE_PRESETS)) {
      expect(sanitizeSlideSizePreset(key)).toBe(key);
    }
  });
});

const BASE_PERSISTED = {
  recentColors: [] as string[],
  colorTarget: "fill" as const,
  pasteParts: { fill: true, line: true, font: true, size: true },
  fontSizeStep: 1,
};

describe("toCommandOptions", () => {
  it("resolves a forced preset to its slide size", () => {
    expect(
      toCommandOptions({ matrix: DEFAULT_MATRIX, slideSizePreset: "16x9-legacy", ...BASE_PERSISTED }).slideSize,
    ).toEqual({
      width: 720,
      height: 405,
    });
  });

  it("leaves slideSize undefined for 'auto'", () => {
    expect(
      toCommandOptions({ matrix: DEFAULT_MATRIX, slideSizePreset: "auto", ...BASE_PERSISTED }).slideSize,
    ).toBeUndefined();
  });

  it("passes the matrix through unchanged", () => {
    const matrix = { columns: 4, gapX: 10, gapY: 12 };
    expect(toCommandOptions({ matrix, slideSizePreset: "auto", ...BASE_PERSISTED }).matrix).toEqual(matrix);
  });

  it("includes colorTarget and pasteParts", () => {
    const pasteParts = { fill: true, line: false, font: true, size: false };
    const result = toCommandOptions({
      matrix: DEFAULT_MATRIX,
      slideSizePreset: "auto",
      recentColors: [],
      colorTarget: "line",
      pasteParts,
      fontSizeStep: 1,
    });
    expect(result.colorTarget).toBe("line");
    expect(result.pasteParts).toEqual(pasteParts);
  });

  it("never includes a color field (that's passed at click time)", () => {
    const result = toCommandOptions({
      matrix: DEFAULT_MATRIX,
      slideSizePreset: "auto",
      recentColors: [],
      colorTarget: "fill",
      pasteParts: { fill: true, line: true, font: true, size: true },
      fontSizeStep: 1,
    });
    expect(result.color).toBeUndefined();
  });

  it("includes the fontSizeStep", () => {
    const result = toCommandOptions({ ...DEFAULT_PERSISTED, fontSizeStep: 2.5 });
    expect(result.fontSizeStep).toBe(2.5);
  });
});

describe("sanitizeRecentColors", () => {
  it("returns an empty array for non-array input", () => {
    expect(sanitizeRecentColors(undefined)).toEqual([]);
    expect(sanitizeRecentColors(null)).toEqual([]);
    expect(sanitizeRecentColors("not-an-array")).toEqual([]);
    expect(sanitizeRecentColors({})).toEqual([]);
  });

  it("filters out invalid entries", () => {
    expect(sanitizeRecentColors(["#FF0000", 42, null, {}, "not-a-color", "#00ff00"])).toEqual(["#FF0000", "#00FF00"]);
  });

  it("normalizes each valid color (3-digit shorthand expands and dedups against its 6-digit form)", () => {
    expect(sanitizeRecentColors(["abc", "#AABBCC"])).toEqual(["#AABBCC"]);
  });

  it("normalizes a 3-digit shorthand to its 6-digit uppercase form", () => {
    expect(sanitizeRecentColors(["abc"])).toEqual(["#AABBCC"]);
  });

  it("deduplicates normalized colors", () => {
    expect(sanitizeRecentColors(["#ff0000", "#FF0000", "#F00"])).toEqual(["#FF0000"]);
  });

  it("caps the result at MAX_RECENT_COLORS entries", () => {
    const input = Array.from({ length: 20 }, (_, i) => `#${i.toString(16).padStart(2, "0")}0000`);
    const result = sanitizeRecentColors(input);
    expect(result.length).toBe(MAX_RECENT_COLORS);
  });
});

describe("sanitizeColorTarget", () => {
  it("returns 'fill' for garbage input", () => {
    expect(sanitizeColorTarget(undefined)).toBe("fill");
    expect(sanitizeColorTarget(null)).toBe("fill");
    expect(sanitizeColorTarget(42)).toBe("fill");
    expect(sanitizeColorTarget("not-a-target")).toBe("fill");
  });

  it("passes through 'line' and 'font'", () => {
    expect(sanitizeColorTarget("line")).toBe("line");
    expect(sanitizeColorTarget("font")).toBe("font");
  });

  it("passes through 'fill' explicitly", () => {
    expect(sanitizeColorTarget("fill")).toBe("fill");
  });
});

describe("sanitizePasteParts", () => {
  it("defaults every field to true for garbage input", () => {
    expect(sanitizePasteParts(undefined)).toEqual({ fill: true, line: true, font: true, size: true });
    expect(sanitizePasteParts(null)).toEqual({ fill: true, line: true, font: true, size: true });
    expect(sanitizePasteParts("not-an-object")).toEqual({ fill: true, line: true, font: true, size: true });
  });

  it("passes through valid booleans", () => {
    expect(sanitizePasteParts({ fill: false, line: true, font: false, size: true })).toEqual({
      fill: false,
      line: true,
      font: false,
      size: true,
    });
  });

  it("defaults individual non-boolean fields to true", () => {
    expect(sanitizePasteParts({ fill: false, line: "no", font: undefined })).toEqual({
      fill: false,
      line: true,
      font: true,
      size: true,
    });
  });
});

describe("sanitizeFontSizeStep", () => {
  it("returns 1 for garbage input", () => {
    expect(sanitizeFontSizeStep(undefined)).toBe(1);
    expect(sanitizeFontSizeStep(null)).toBe(1);
    expect(sanitizeFontSizeStep("not-a-number")).toBe(1);
    expect(sanitizeFontSizeStep(NaN)).toBe(1);
    expect(sanitizeFontSizeStep(Infinity)).toBe(1);
    expect(sanitizeFontSizeStep(-Infinity)).toBe(1);
  });

  it("clamps finite values outside [0.5, 20] instead of resetting to 1", () => {
    expect(sanitizeFontSizeStep(0)).toBe(0.5);
    expect(sanitizeFontSizeStep(0.1)).toBe(0.5);
    expect(sanitizeFontSizeStep(-3)).toBe(0.5);
    expect(sanitizeFontSizeStep(25)).toBe(20);
    expect(sanitizeFontSizeStep(20.1)).toBe(20);
  });

  it("passes through valid values unchanged", () => {
    expect(sanitizeFontSizeStep(0.5)).toBe(0.5);
    expect(sanitizeFontSizeStep(2)).toBe(2);
    expect(sanitizeFontSizeStep(20)).toBe(20);
  });

  it("rounds to the nearest 0.5", () => {
    expect(sanitizeFontSizeStep(1.2)).toBe(1);
    expect(sanitizeFontSizeStep(1.3)).toBe(1.5);
    expect(sanitizeFontSizeStep(1.74)).toBe(1.5);
    expect(sanitizeFontSizeStep(1.76)).toBe(2);
  });
});
