import { describe, expect, it } from "vitest";
import {
  BASIC_SHAPE_RECT,
  chevronChainLayout,
  columnsLayout,
  contentArea,
  gridLayout,
  numberedCirclesLayout,
  stickerRect,
} from "../../src/core/insert";
import type { Rect } from "../../src/core/types";
import { slide16x9 } from "./fixtures";

/** Compara dos rects con tolerancia de punto flotante (precisión de `toBeCloseTo`). */
function expectRect(actual: Rect, expected: Rect, precision = 9): void {
  expect(actual.left).toBeCloseTo(expected.left, precision);
  expect(actual.top).toBeCloseTo(expected.top, precision);
  expect(actual.width).toBeCloseTo(expected.width, precision);
  expect(actual.height).toBeCloseTo(expected.height, precision);
}

describe("contentArea", () => {
  it("descuenta 8% lateral, 20% superior y 8% inferior", () => {
    expectRect(contentArea(slide16x9), { left: 76.8, top: 108, width: 806.4, height: 388.8 });
  });
});

describe("columnsLayout", () => {
  it("reparte count columnas iguales con huecos de gapRatio*width entre ellas", () => {
    const rects = columnsLayout(slide16x9, 3);
    // area.width = 806.4, gap = 960*0.02 = 19.2, width = (806.4 - 2*19.2)/3 = 256
    expect(rects).toHaveLength(3);
    expectRect(rects[0]!, { left: 76.8, top: 108, width: 256, height: 388.8 });
    expectRect(rects[1]!, { left: 76.8 + 256 + 19.2, top: 108, width: 256, height: 388.8 });
    expectRect(rects[2]!, { left: 76.8 + 2 * (256 + 19.2), top: 108, width: 256, height: 388.8 });
  });

  it("acota count a 1..12 (por debajo)", () => {
    expect(columnsLayout(slide16x9, 0)).toHaveLength(1);
    expect(columnsLayout(slide16x9, -5)).toHaveLength(1);
  });

  it("acota count a 1..12 (por encima)", () => {
    expect(columnsLayout(slide16x9, 20)).toHaveLength(12);
  });

  it("acepta un gapRatio distinto", () => {
    const rects = columnsLayout(slide16x9, 2, 0.05);
    const gap = 960 * 0.05;
    const width = (806.4 - gap) / 2;
    expectRect(rects[0]!, { left: 76.8, top: 108, width, height: 388.8 });
    expectRect(rects[1]!, { left: 76.8 + width + gap, top: 108, width, height: 388.8 });
  });
});

describe("gridLayout", () => {
  it("reparte rows x cols celdas fila a fila", () => {
    const rects = gridLayout(slide16x9, 2, 2);
    // hGap = 960*0.02=19.2, vGap = 540*0.02=10.8; width=(806.4-19.2)/2=393.6; height=(388.8-10.8)/2=189
    expect(rects).toHaveLength(4);
    expectRect(rects[0]!, { left: 76.8, top: 108, width: 393.6, height: 189 });
    expectRect(rects[1]!, { left: 76.8 + 393.6 + 19.2, top: 108, width: 393.6, height: 189 });
    expectRect(rects[2]!, { left: 76.8, top: 108 + 189 + 10.8, width: 393.6, height: 189 });
    expectRect(rects[3]!, { left: 76.8 + 393.6 + 19.2, top: 108 + 189 + 10.8, width: 393.6, height: 189 });
  });

  it("acota rows y cols a 1..12 de forma independiente", () => {
    expect(gridLayout(slide16x9, 0, 20)).toHaveLength(12);
    expect(gridLayout(slide16x9, 3, -1)).toHaveLength(3);
  });
});

describe("chevronChainLayout", () => {
  it("reparte count chevrones de igual ancho, sin solape, centrados verticalmente", () => {
    const rects = chevronChainLayout(slide16x9, 4);
    // height = 540*0.12 = 64.8; width = 806.4/4 = 201.6; top = 108 + (388.8-64.8)/2 = 270
    expect(rects).toHaveLength(4);
    for (const r of rects) {
      expect(r.width).toBeCloseTo(201.6, 9);
      expect(r.height).toBeCloseTo(64.8, 9);
      expect(r.top).toBeCloseTo(270, 9);
    }
    expect(rects[0]!.left).toBeCloseTo(76.8, 9);
    expect(rects[1]!.left).toBeCloseTo(76.8 + 201.6, 9);
    expect(rects[2]!.left).toBeCloseTo(76.8 + 2 * 201.6, 9);
    expect(rects[3]!.left).toBeCloseTo(76.8 + 3 * 201.6, 9);
  });

  it("no se solapan entre sí (cada uno empieza donde termina el anterior)", () => {
    const rects = chevronChainLayout(slide16x9, 5);
    for (let i = 1; i < rects.length; i++) {
      expect(rects[i]!.left).toBeCloseTo(rects[i - 1]!.left + rects[i - 1]!.width, 9);
    }
  });

  it("acota count a 1..12", () => {
    expect(chevronChainLayout(slide16x9, 50)).toHaveLength(12);
    expect(chevronChainLayout(slide16x9, 0)).toHaveLength(1);
  });
});

describe("numberedCirclesLayout", () => {
  it("reparte count círculos de diámetro 7% del ancho, en una fila centrada verticalmente", () => {
    const rects = numberedCirclesLayout(slide16x9, 5);
    // diameter = 960*0.07 = 67.2; slot = 806.4/5 = 161.28; top = 108+(388.8-67.2)/2 = 268.8
    expect(rects).toHaveLength(5);
    for (const r of rects) {
      expect(r.width).toBeCloseTo(67.2, 9);
      expect(r.height).toBeCloseTo(67.2, 9);
      expect(r.top).toBeCloseTo(268.8, 9);
    }
    const slot = 806.4 / 5;
    expect(rects[0]!.left).toBeCloseTo(76.8 + (slot - 67.2) / 2, 9);
    expect(rects[4]!.left).toBeCloseTo(76.8 + 4 * slot + (slot - 67.2) / 2, 9);
  });

  it("acota count a 1..12", () => {
    expect(numberedCirclesLayout(slide16x9, 50)).toHaveLength(12);
    expect(numberedCirclesLayout(slide16x9, -3)).toHaveLength(1);
  });
});

describe("stickerRect", () => {
  it("coloca una etiqueta 18% x 7% en la esquina superior derecha, con 3% de margen", () => {
    expectRect(stickerRect(slide16x9), { left: 960 * (1 - 0.03 - 0.18), top: 540 * 0.03, width: 172.8, height: 37.8 });
  });
});

describe("BASIC_SHAPE_RECT", () => {
  it("devuelve un cuadrado centrado, con lado 25% del lado menor de la diapositiva", () => {
    // min(960,540)=540; side=135; left=(960-135)/2=412.5; top=(540-135)/2=202.5
    expectRect(BASIC_SHAPE_RECT(slide16x9), { left: 412.5, top: 202.5, width: 135, height: 135 });
  });

  it("usa el ancho cuando es el lado menor", () => {
    const portrait = { width: 540, height: 960 };
    const side = 540 * 0.25;
    expectRect(BASIC_SHAPE_RECT(portrait), { left: (540 - side) / 2, top: (960 - side) / 2, width: side, height: side });
  });
});
