import { describe, expect, it } from "vitest";
import { pngDimensions, slideSizeFromAspect } from "../../src/core/png";

// PNG 1x1 válido (cabecera estándar)
const PNG_1x1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

describe("pngDimensions", () => {
  it("reads width and height from the IHDR chunk", () => {
    expect(pngDimensions(PNG_1x1)).toEqual({ width: 1, height: 1 });
  });

  it("throws on non-PNG data", () => {
    expect(() => pngDimensions(btoa("not a png at all, really"))).toThrow(/PNG/);
  });
});

describe("slideSizeFromAspect", () => {
  it("assumes 540pt height and scales width by aspect ratio", () => {
    expect(slideSizeFromAspect({ width: 160, height: 90 })).toEqual({ width: 960, height: 540 });
    expect(slideSizeFromAspect({ width: 40, height: 30 })).toEqual({ width: 720, height: 540 });
  });
});
