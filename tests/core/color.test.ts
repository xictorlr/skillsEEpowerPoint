import { describe, expect, it } from "vitest";
import { DEFAULT_PALETTE, normalizeHex, pushRecent } from "../../src/core/color";

describe("normalizeHex", () => {
  it("accepts 6-digit hex with or without #, upper-casing it", () => {
    expect(normalizeHex("#ffa500")).toBe("#FFA500");
    expect(normalizeHex("ffa500")).toBe("#FFA500");
  });
  it("expands 3-digit hex", () => {
    expect(normalizeHex("#abc")).toBe("#AABBCC");
  });
  it("rejects anything else", () => {
    expect(normalizeHex("orange")).toBeNull();
    expect(normalizeHex("#12345")).toBeNull();
    expect(normalizeHex("")).toBeNull();
  });
});

describe("DEFAULT_PALETTE", () => {
  it("has 16 unique normalized colors", () => {
    expect(DEFAULT_PALETTE).toHaveLength(16);
    expect(new Set(DEFAULT_PALETTE).size).toBe(16);
    expect(DEFAULT_PALETTE.every((c) => normalizeHex(c) === c)).toBe(true);
  });
});

describe("pushRecent", () => {
  it("prepends, de-duplicates and caps the list", () => {
    expect(pushRecent(["#111111", "#222222"], "#222222", 8)).toEqual(["#222222", "#111111"]);
    const eight = ["#000001", "#000002", "#000003", "#000004", "#000005", "#000006", "#000007", "#000008"];
    expect(pushRecent(eight, "#000009", 8)).toEqual(["#000009", ...eight.slice(0, 7)]);
  });
  it("ignores invalid colors", () => {
    expect(pushRecent(["#111111"], "nope", 8)).toEqual(["#111111"]);
  });
});
