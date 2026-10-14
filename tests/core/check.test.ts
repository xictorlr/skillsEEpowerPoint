import { describe, expect, it } from "vitest";
import { agendaLayout, distinctFonts, isOffSlide, parseAgendaSections } from "../../src/core/check";
import { shape, slide16x9 } from "./fixtures";

describe("isOffSlide", () => {
  it("returns false for a shape partially inside the slide (hanging off the right edge)", () => {
    const s = shape({ left: 900, top: 100, width: 100, height: 50 });
    expect(isOffSlide(s, slide16x9)).toBe(false);
  });

  it("returns true for a shape fully off the right edge", () => {
    const s = shape({ left: 970, top: 100, width: 50, height: 50 });
    expect(isOffSlide(s, slide16x9)).toBe(true);
  });

  it("returns true for a shape fully off the left edge (negative coordinates)", () => {
    const s = shape({ left: -200, top: 100, width: 50, height: 50 });
    expect(isOffSlide(s, slide16x9)).toBe(true);
  });

  it("returns true for a shape fully above or fully below the slide", () => {
    const above = shape({ left: 100, top: -100, width: 50, height: 50 });
    const below = shape({ left: 100, top: 600, width: 50, height: 50 });
    expect(isOffSlide(above, slide16x9)).toBe(true);
    expect(isOffSlide(below, slide16x9)).toBe(true);
  });

  it("returns false for a shape fully contained within the slide", () => {
    const s = shape({ left: 10, top: 10, width: 50, height: 50 });
    expect(isOffSlide(s, slide16x9)).toBe(false);
  });
});

describe("distinctFonts", () => {
  it("deduplicates, trims, drops empty/null/undefined, and sorts with localeCompare('es')", () => {
    expect(
      distinctFonts(["Arial", " Calibri ", "Arial", "", null, undefined, "  ", "Ábaco", "Zebra", "arial"]),
    ).toEqual(["Ábaco", "arial", "Arial", "Calibri", "Zebra"]);
  });

  it("returns [] for an empty or all-blank input", () => {
    expect(distinctFonts([])).toEqual([]);
    expect(distinctFonts([null, undefined, "", "   "])).toEqual([]);
  });
});

describe("parseAgendaSections", () => {
  it("splits on newlines, trims, and drops empty lines", () => {
    expect(parseAgendaSections("Intro\r\n  Market overview  \n\nStrategy\n")).toEqual([
      "Intro",
      "Market overview",
      "Strategy",
    ]);
  });

  it("caps at 20 sections", () => {
    const lines = Array.from({ length: 25 }, (_, i) => `Section ${i + 1}`);
    const result = parseAgendaSections(lines.join("\n"));
    expect(result).toHaveLength(20);
    expect(result[0]).toBe("Section 1");
    expect(result[19]).toBe("Section 20");
  });

  it("returns [] for empty/blank text", () => {
    expect(parseAgendaSections("")).toEqual([]);
    expect(parseAgendaSections("\n\n  \n")).toEqual([]);
  });
});

describe("agendaLayout", () => {
  it("lays out a single-column agenda for 960x540 and 3 sections", () => {
    const spec = agendaLayout(slide16x9, { title: "Agenda", sections: ["Intro", "Market", "Plan"] }, 1);

    expect(spec.title.text).toBe("Agenda");
    expect(spec.title.rect.left).toBeCloseTo(76.8);
    expect(spec.title.rect.top).toBeCloseTo(54);
    expect(spec.title.rect.width).toBeCloseTo(806.4);
    expect(spec.title.rect.height).toBeCloseTo(64.8);

    expect(spec.items).toHaveLength(3);
    const tops = spec.items.map((item) => item.rect.top);
    expect(tops[0]).toBeCloseTo(151.2);
    expect(tops[1]).toBeCloseTo(194.4);
    expect(tops[2]).toBeCloseTo(237.6);

    for (const item of spec.items) {
      expect(item.rect.left).toBeCloseTo(76.8);
      expect(item.rect.width).toBeCloseTo(806.4);
      expect(item.rect.height).toBeCloseTo(43.2);
    }

    expect(spec.items.map((item) => item.highlighted)).toEqual([false, true, false]);
    expect(spec.items.map((item) => item.text)).toEqual(["Intro", "Market", "Plan"]);
  });

  it("splits into two columns of 6 for 12 sections, restarting tops at 151.2", () => {
    const sections = Array.from({ length: 12 }, (_, i) => `Section ${i + 1}`);
    const spec = agendaLayout(slide16x9, { title: "Agenda", sections }, 7);

    expect(spec.items).toHaveLength(12);

    const first = spec.items[0];
    const seventh = spec.items[6];
    expect(first).toBeDefined();
    expect(seventh).toBeDefined();
    expect(seventh!.rect.left).toBeGreaterThan(first!.rect.left);
    expect(seventh!.rect.top).toBeCloseTo(151.2);
    expect(first!.rect.top).toBeCloseTo(151.2);

    // Columns must stay within the slide's lateral margins (8% each side).
    expect(seventh!.rect.left).toBeCloseTo(489.6);
    expect(seventh!.rect.width).toBeCloseTo(393.6);
    expect(seventh!.rect.left + seventh!.rect.width).toBeCloseTo(883.2);

    // Each column holds at most 10 rows; with 12 sections, 6 and 6.
    const leftColLeft = first!.rect.left;
    const rightColLeft = seventh!.rect.left;
    const leftCount = spec.items.filter((item) => item.rect.left === leftColLeft).length;
    const rightCount = spec.items.filter((item) => item.rect.left === rightColLeft).length;
    expect(leftCount).toBe(6);
    expect(rightCount).toBe(6);

    expect(spec.items[7]!.highlighted).toBe(true);
    expect(spec.items.filter((item) => item.highlighted)).toHaveLength(1);
  });

  it("keeps the full 43.2 row height for a small (3-section) deck", () => {
    const spec = agendaLayout(slide16x9, { title: "Agenda", sections: ["Intro", "Market", "Plan"] }, 0);
    for (const item of spec.items) {
      expect(item.rect.height).toBeCloseTo(43.2);
    }
  });

  it("shrinks row height so 10 sections never overflow past the bottom margin (92% of H)", () => {
    const sections = Array.from({ length: 10 }, (_, i) => `Section ${i + 1}`);
    const spec = agendaLayout(slide16x9, { title: "Agenda", sections }, 0);

    expect(spec.items).toHaveLength(10);
    const bottomLimit = slide16x9.height * 0.92;
    for (const item of spec.items) {
      expect(item.rect.top + item.rect.height).toBeLessThanOrEqual(bottomLimit + 0.01);
    }
  });

  it("shrinks row height so 20 sections (two columns) never overflow past the bottom margin (92% of H)", () => {
    const sections = Array.from({ length: 20 }, (_, i) => `Section ${i + 1}`);
    const spec = agendaLayout(slide16x9, { title: "Agenda", sections }, 0);

    expect(spec.items).toHaveLength(20);
    const bottomLimit = slide16x9.height * 0.92;
    for (const item of spec.items) {
      expect(item.rect.top + item.rect.height).toBeLessThanOrEqual(bottomLimit + 0.01);
    }
  });

  it("clamps sections to a maximum of 20, even with more input sections", () => {
    const sections = Array.from({ length: 25 }, (_, i) => `Section ${i + 1}`);
    const spec = agendaLayout(slide16x9, { title: "Agenda", sections }, 0);
    expect(spec.items).toHaveLength(20);
    expect(spec.items[19]!.text).toBe("Section 20");
  });
});
