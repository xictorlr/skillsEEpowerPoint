import { describe, expect, it } from "vitest";
import {
  DEFAULT_MATRIX,
  SLIDE_SIZE_PRESETS,
  sanitizeMatrix,
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

describe("toCommandOptions", () => {
  it("resolves a forced preset to its slide size", () => {
    expect(toCommandOptions({ matrix: DEFAULT_MATRIX, slideSizePreset: "16x9-legacy" }).slideSize).toEqual({
      width: 720,
      height: 405,
    });
  });

  it("leaves slideSize undefined for 'auto'", () => {
    expect(toCommandOptions({ matrix: DEFAULT_MATRIX, slideSizePreset: "auto" }).slideSize).toBeUndefined();
  });

  it("passes the matrix through unchanged", () => {
    const matrix = { columns: 4, gapX: 10, gapY: 12 };
    expect(toCommandOptions({ matrix, slideSizePreset: "auto" }).matrix).toEqual(matrix);
  });
});
