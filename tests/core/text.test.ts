import { describe, expect, it } from "vitest";
import {
  autoSizePatches, casePatches, changeCase, cleanPatches, cleanText, computeTextEdits, fontSizePatches,
  marginPatches, paragraphAlignPatches, splitLines, verticalAlignPatches, wordWrapPatches,
  STANDARD_MARGINS, ZERO_MARGINS, type TextShape,
} from "../../src/core/text";

function ts(id: string, over: Partial<TextShape> = {}): TextShape {
  return {
    id, text: "hola mundo", fontSize: 12, horizontalAlignment: "Left", verticalAlignment: "Top",
    autoSize: "AutoSizeNone", wordWrap: true, margins: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 }, ...over,
  };
}

/** Aplica ediciones de sub-rango (en el orden del arreglo, como haría el adaptador Office). */
function applyEdits(text: string, edits: { start: number; length: number; text: string }[]): string {
  let result = text;
  for (const e of edits) result = result.slice(0, e.start) + e.text + result.slice(e.start + e.length);
  return result;
}

describe("splitLines", () => {
  it("splits on \\r, \\n and \\v keeping the separators", () => {
    expect(splitLines("a\rb\nc\vd")).toEqual([
      { text: "a", sep: "\r" }, { text: "b", sep: "\n" }, { text: "c", sep: "\v" }, { text: "d", sep: "" },
    ]);
  });
});

describe("changeCase", () => {
  it("upper and lower are locale-aware", () => {
    expect(changeCase("año nuevo", "upper")).toBe("AÑO NUEVO");
    expect(changeCase("AÑO NUEVO", "lower")).toBe("año nuevo");
  });
  it("sentence case capitalizes the first letter of each sentence and lowercases the rest", () => {
    expect(changeCase("hOLA mundo. qué TAL? bien! ok", "sentence")).toBe("Hola mundo. Qué tal? Bien! Ok");
  });
  it("title case capitalizes every word", () => {
    expect(changeCase("resultados del tercer trimestre", "title")).toBe("Resultados Del Tercer Trimestre");
  });
  it("preserves paragraph separators", () => {
    expect(changeCase("uno\rdos\vtres", "upper")).toBe("UNO\rDOS\vTRES");
    expect(changeCase("uno. dos\rtres", "sentence")).toBe("Uno. Dos\rTres");
  });
  it("sentence case only ends a sentence when a terminator is followed by whitespace (quotes/brackets may sit between)", () => {
    expect(changeCase("versión 2.5 del plan", "sentence")).toBe("Versión 2.5 del plan");
    expect(changeCase("3 puntos. ok", "sentence")).toBe("3 puntos. Ok");
    expect(changeCase('dijo "hola." luego', "sentence")).toBe('Dijo "hola." Luego');
    expect(changeCase("espera... vale", "sentence")).toBe("Espera... Vale");
    expect(changeCase("¿qué tal? ¡bien!", "sentence")).toBe("¿Qué tal? ¡Bien!");
  });
  it("sentence case carries capitalization state across soft breaks (\\v) but resets on hard breaks (\\r, \\n)", () => {
    expect(changeCase("uno\vdos. tres", "sentence")).toBe("Uno\vdos. Tres");
  });
  it("title case capitalizes after apostrophes only at word start, and after hyphens/quotes/brackets", () => {
    expect(changeCase("the cat's hat", "title")).toBe("The Cat's Hat");
    expect(changeCase("e-mail", "title")).toBe("E-Mail");
    expect(changeCase("«cita»", "title")).toBe("«Cita»");
    expect(changeCase("(nota)", "title")).toBe("(Nota)");
  });
});

describe("cleanText", () => {
  it("collapses repeated spaces, trims line ends and drops empty lines", () => {
    expect(cleanText("  hola   mundo  \r\r  qué  tal \r")).toBe("hola mundo\rqué tal");
  });
  it("returns the same text when already clean", () => {
    expect(cleanText("hola\rmundo")).toBe("hola\rmundo");
  });
  it("keeps tabs and non-breaking spaces untouched, only collapses runs of ordinary spaces", () => {
    expect(cleanText("a\tb")).toBe("a\tb");
    expect(cleanText("10 %")).toBe("10 %");
    expect(cleanText("a\v\vb")).toBe("a\vb");
  });
});

describe("computeTextEdits", () => {
  it("returns [] for identical texts", () => {
    expect(computeTextEdits("hola mundo", "hola mundo")).toEqual([]);
  });
  it("emits word-level edits only for changed words, sorted by start descending", () => {
    expect(computeTextEdits("hola MUNDO bien", "HOLA MUNDO BIEN")).toEqual([
      { start: 11, length: 4, text: "BIEN" },
      { start: 0, length: 4, text: "HOLA" },
    ]);
  });
  it("emits a whole-line edit when the line length changes", () => {
    expect(computeTextEdits("ß", "SS")).toEqual([{ start: 0, length: 1, text: "SS" }]);
  });
  it("in a multi-paragraph text, only edits the changed paragraph with correct absolute offsets", () => {
    expect(computeTextEdits("uno\rdos tres\rcuatro", "uno\rDOS tres\rcuatro")).toEqual([
      { start: 4, length: 3, text: "DOS" },
    ]);
  });
  it("falls back to a single whole-text edit when paragraph counts differ", () => {
    expect(computeTextEdits("hola\r\rmundo", "hola\rmundo")).toEqual([
      { start: 0, length: 11, text: "hola\rmundo" },
    ]);
  });
  it("falls back to a whole-line edit when word tokens are not aligned (count/position/length), even if the line length matches", () => {
    expect(computeTextEdits("a b", "ab ")).toEqual([{ start: 0, length: 3, text: "ab " }]);
  });
  it("falls back to a single whole-text edit when a separator changes, even with matching line counts", () => {
    expect(computeTextEdits("a\rb", "a\vb")).toEqual([{ start: 0, length: 3, text: "a\vb" }]);
  });
});

describe("computeTextEdits + applyEdits round-trip", () => {
  const cases: Array<[string, string]> = [
    ["hola mundo", "hola mundo"],
    ["hola MUNDO bien", "HOLA MUNDO BIEN"],
    ["ß", "SS"],
    ["uno\rdos tres\rcuatro", "uno\rDOS tres\rcuatro"],
    ["hola\r\rmundo", "hola\rmundo"],
    ["a b", "ab "],
    ["a b", "A\tb"],
    ["a\rb", "a\vb"],
  ];
  it.each(cases)("applying computeTextEdits(%j, %j) to the old text reproduces the new text", (oldText, newText) => {
    expect(applyEdits(oldText, computeTextEdits(oldText, newText))).toBe(newText);
  });
});

describe("patch generators", () => {
  it("casePatches only emits sub-range edits for shapes that actually change", () => {
    expect(casePatches([ts("a"), ts("b", { text: "HOLA MUNDO" })], "upper")).toEqual([
      { id: "a", edits: [{ start: 5, length: 5, text: "MUNDO" }, { start: 0, length: 4, text: "HOLA" }] },
    ]);
  });
  it("fontSizePatches steps and clamps, skipping mixed sizes", () => {
    expect(fontSizePatches([ts("a"), ts("b", { fontSize: null }), ts("c", { fontSize: 1 })], -1)).toEqual([{ id: "a", fontSize: 11 }]);
    expect(fontSizePatches([ts("a", { fontSize: 400 })], 1)).toEqual([]);
  });
  it("fontSizePatches rejects a no-op or non-finite delta and rounds to 0.1", () => {
    expect(fontSizePatches([ts("a")], 0)).toEqual([]);
    expect(fontSizePatches([ts("a")], NaN)).toEqual([]);
    expect(fontSizePatches([ts("a", { fontSize: 10.5 })], 1)).toEqual([{ id: "a", fontSize: 11.5 }]);
    expect(fontSizePatches([ts("a", { fontSize: 399.5 })], 1)).toEqual([{ id: "a", fontSize: 400 }]);
  });
  it("paragraph and vertical alignment skip shapes already aligned", () => {
    expect(paragraphAlignPatches([ts("a"), ts("b", { horizontalAlignment: "Center" })], "Center")).toEqual([{ id: "a", horizontalAlignment: "Center" }]);
    expect(verticalAlignPatches([ts("a")], "Middle")).toEqual([{ id: "a", verticalAlignment: "Middle" }]);
    expect(verticalAlignPatches([ts("a")], "Top")).toEqual([]);
  });
  it("verticalAlignPatches compares on the base value and preserves a *Centered suffix", () => {
    expect(verticalAlignPatches([ts("a", { verticalAlignment: "MiddleCentered" })], "Middle")).toEqual([]);
    expect(verticalAlignPatches([ts("a", { verticalAlignment: "TopCentered" })], "Middle")).toEqual([
      { id: "a", verticalAlignment: "MiddleCentered" },
    ]);
  });
  it("paragraphAlignPatches treats a mixed (null) alignment as needing a patch", () => {
    expect(paragraphAlignPatches([ts("a", { horizontalAlignment: null })], "Center")).toEqual([{ id: "a", horizontalAlignment: "Center" }]);
  });
  it("marginPatches applies zero or standard presets", () => {
    expect(marginPatches([ts("a")], "zero")).toEqual([{ id: "a", margins: ZERO_MARGINS }]);
    expect(marginPatches([ts("a")], "standard")).toEqual([]);
    expect(STANDARD_MARGINS).toEqual({ left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 });
  });
  it("autoSize and wordWrap patches", () => {
    expect(autoSizePatches([ts("a")], "AutoSizeShapeToFitText")).toEqual([{ id: "a", autoSize: "AutoSizeShapeToFitText" }]);
    expect(wordWrapPatches([ts("a")], false)).toEqual([{ id: "a", wordWrap: false }]);
    expect(wordWrapPatches([ts("a")], true)).toEqual([]);
  });
  it("cleanPatches only emits sub-range edits when the text changes", () => {
    expect(cleanPatches([ts("a", { text: "a  b" }), ts("b")])).toEqual([
      { id: "a", edits: [{ start: 0, length: 4, text: "a b" }] },
    ]);
  });
  it("all generators return [] for empty input", () => {
    expect(casePatches([], "upper")).toEqual([]);
    expect(cleanPatches([])).toEqual([]);
  });
});
