import { describe, expect, it } from "vitest";
import { applyPatches } from "../../src/office/apply";
import { MockShape } from "./mockPowerPoint";

describe("applyPatches", () => {
  it("writes only the fields present in each patch", () => {
    const a = new MockShape({ id: "a", left: 0, top: 0, width: 10, height: 10 });
    const b = new MockShape({ id: "b", left: 0, top: 0, width: 10, height: 10 });
    const proxies = new Map<string, MockShape>([["a", a], ["b", b]]);
    applyPatches(proxies, [{ id: "a", left: 5 }, { id: "b", top: 7, height: 20 }]);
    expect(a.writes).toEqual([{ prop: "left", value: 5 }]);
    expect(b.writes).toEqual([{ prop: "top", value: 7 }, { prop: "height", value: 20 }]);
  });

  it("ignores patches whose id is not in the map", () => {
    const a = new MockShape({ id: "a", left: 0, top: 0, width: 10, height: 10 });
    applyPatches(new Map([["a", a]]), [{ id: "zzz", left: 1 }]);
    expect(a.writes).toEqual([]);
  });

  it("skips non-finite values (NaN, Infinity) instead of writing them", () => {
    const a = new MockShape({ id: "a", left: 0, top: 0, width: 10, height: 10 });
    applyPatches(new Map([["a", a]]), [{ id: "a", left: NaN, top: Infinity, width: -Infinity, height: 5 }]);
    expect(a.writes).toEqual([{ prop: "height", value: 5 }]);
  });
});
