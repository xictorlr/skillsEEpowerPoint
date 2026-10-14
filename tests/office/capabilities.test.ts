import { afterEach, describe, expect, it, vi } from "vitest";
import { supportsApi } from "../../src/office/capabilities";

describe("supportsApi", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("returns false when the Office global isn't available (e.g. in Node)", () => {
    expect(supportsApi("1.10")).toBe(false);
  });

  it("returns true only when the requirement set and version are both supported", () => {
    vi.stubGlobal("Office", {
      context: {
        requirements: {
          isSetSupported: (name: string, version: string) => name === "PowerPointApi" && version === "1.10",
        },
      },
    });
    expect(supportsApi("1.10")).toBe(true);
    expect(supportsApi("1.11")).toBe(false);
  });

  it("returns false if Office.context.requirements.isSetSupported throws", () => {
    vi.stubGlobal("Office", {
      context: {
        requirements: {
          isSetSupported: () => {
            throw new Error("boom");
          },
        },
      },
    });
    expect(supportsApi("1.10")).toBe(false);
  });

  it("returns false if Office.context is missing", () => {
    vi.stubGlobal("Office", {});
    expect(supportsApi("1.10")).toBe(false);
  });
});
