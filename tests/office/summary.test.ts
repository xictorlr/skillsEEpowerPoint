import { describe, expect, it } from "vitest";
import { readSelectionSummary } from "../../src/office/summary";
import { MockContext, MockShape } from "./mockPowerPoint";

describe("readSelectionSummary", () => {
  it("summarizes count, active slide (2nd of 3) and the bbox union of 2 selected shapes", async () => {
    const a = new MockShape({ id: "a", left: 10, top: 20, width: 30, height: 40 });
    const b = new MockShape({ id: "b", left: 50, top: 10, width: 10, height: 10 });
    const ctx = new MockContext([a, b], "", {
      slides: [{ id: "s1" }, { id: "s2" }, { id: "s3" }],
      activeSlideIndex: 1,
    });

    const summary = await readSelectionSummary(ctx as never);

    expect(summary).toEqual({
      count: 2,
      slideId: "s2",
      slideIndex: 1,
      slideCount: 3,
      bbox: { left: 10, top: 10, width: 50, height: 50 },
    });
    expect(ctx.syncCount).toBe(1);
  });

  it("returns bbox null for an empty selection, keeping slide navigation info", async () => {
    const ctx = new MockContext([], "", { slides: [{ id: "s1" }, { id: "s2" }], activeSlideIndex: 0 });

    const summary = await readSelectionSummary(ctx as never);

    expect(summary).toEqual({ count: 0, slideId: "s1", slideIndex: 0, slideCount: 2, bbox: null });
  });

  it("returns slideId null and slideIndex 0 when the deck has no slides", async () => {
    const ctx = new MockContext([], "", { slides: [] });

    const summary = await readSelectionSummary(ctx as never);

    expect(summary).toEqual({ count: 0, slideId: null, slideIndex: 0, slideCount: 0, bbox: null });
  });
});
