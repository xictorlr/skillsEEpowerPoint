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
import type { Rect, SlideSize } from "../../src/core/types";
import { slide16x9 } from "./fixtures";

/** Compara dos rects con tolerancia de punto flotante (precisión de `toBeCloseTo`). */
function expectRect(actual: Rect, expected: Rect, precision = 9): void {
  expect(actual.left).toBeCloseTo(expected.left, precision);
  expect(actual.top).toBeCloseTo(expected.top, precision);
  expect(actual.width).toBeCloseTo(expected.width, precision);
  expect(actual.height).toBeCloseTo(expected.height, precision);
}

/** Diapositiva vertical (más alta que ancha), para probar que los layouts no asumen 16:9. */
const portrait: SlideSize = { width: 540, height: 960 };

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

  it("con count=1 no hay hueco: la única columna ocupa toda el área de contenido", () => {
    const rects = columnsLayout(slide16x9, 1);
    expect(rects).toHaveLength(1);
    expectRect(rects[0]!, contentArea(slide16x9));
  });

  it("con gapRatio=0 las columnas se tocan y rellenan el área exactamente", () => {
    const rects = columnsLayout(slide16x9, 4, 0);
    const area = contentArea(slide16x9);
    expect(rects).toHaveLength(4);
    const totalWidth = rects.reduce((sum, r) => sum + r.width, 0);
    expect(totalWidth).toBeCloseTo(area.width, 9);
    for (let i = 1; i < rects.length; i++) {
      expect(rects[i]!.left).toBeCloseTo(rects[i - 1]!.left + rects[i - 1]!.width, 9);
    }
    expect(rects[0]!.left).toBeCloseTo(area.left, 9);
  });

  it("con un hueco sobredimensionado (gapRatio grande, muchas columnas) ningún ancho es negativo", () => {
    const rects = columnsLayout(slide16x9, 12, 0.1);
    expect(rects).toHaveLength(12);
    for (const r of rects) {
      expect(Number.isFinite(r.width)).toBe(true);
      // El hueco se acota a area.width/(n-1): en el límite, el ancho da 0 salvo ruido de punto
      // flotante (del orden de 1e-14), nunca negativo de verdad.
      expect(r.width).toBeGreaterThan(-1e-9);
    }
  });

  it("un gapRatio negativo o NaN se trata como 0 (mismo resultado que gapRatio=0)", () => {
    const withZero = columnsLayout(slide16x9, 4, 0);
    const withNegative = columnsLayout(slide16x9, 4, -1);
    const withNaN = columnsLayout(slide16x9, 4, NaN);
    for (let i = 0; i < 4; i++) {
      expectRect(withNegative[i]!, withZero[i]!);
      expectRect(withNaN[i]!, withZero[i]!);
    }
  });

  it("acota count a 1..12 (por debajo)", () => {
    expect(columnsLayout(slide16x9, 0)).toHaveLength(1);
    expect(columnsLayout(slide16x9, -5)).toHaveLength(1);
  });

  it("acota count a 1..12 (por encima)", () => {
    expect(columnsLayout(slide16x9, 20)).toHaveLength(12);
  });

  it("un count no finito (NaN) cae al mínimo (1 columna)", () => {
    expect(columnsLayout(slide16x9, NaN)).toHaveLength(1);
  });

  it("acepta un gapRatio distinto", () => {
    const rects = columnsLayout(slide16x9, 2, 0.05);
    const gap = 960 * 0.05;
    const width = (806.4 - gap) / 2;
    expectRect(rects[0]!, { left: 76.8, top: 108, width, height: 388.8 });
    expectRect(rects[1]!, { left: 76.8 + width + gap, top: 108, width, height: 388.8 });
  });

  it("funciona igual en una diapositiva vertical (no asume 16:9)", () => {
    const rects = columnsLayout(portrait, 3);
    const area = contentArea(portrait);
    const gap = portrait.width * 0.02;
    const width = (area.width - 2 * gap) / 3;
    expect(rects).toHaveLength(3);
    expectRect(rects[0]!, { left: area.left, top: area.top, width, height: area.height });
    expectRect(rects[1]!, { left: area.left + width + gap, top: area.top, width, height: area.height });
    expectRect(rects[2]!, { left: area.left + 2 * (width + gap), top: area.top, width, height: area.height });
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

  it("con gapRatio=0 las celdas se tocan y rellenan el área exactamente", () => {
    const rects = gridLayout(slide16x9, 3, 3, 0);
    const area = contentArea(slide16x9);
    const totalWidth = rects.slice(0, 3).reduce((sum, r) => sum + r.width, 0);
    expect(totalWidth).toBeCloseTo(area.width, 9);
    expect(rects[0]!.left).toBeCloseTo(area.left, 9);
    expect(rects[0]!.top).toBeCloseTo(area.top, 9);
  });

  it("con huecos sobredimensionados (horizontal y vertical) ningún tamaño es negativo", () => {
    const rects = gridLayout(slide16x9, 12, 12, 0.1);
    expect(rects).toHaveLength(144);
    for (const r of rects) {
      expect(Number.isFinite(r.width)).toBe(true);
      expect(Number.isFinite(r.height)).toBe(true);
      // Mismo límite que en columnsLayout: en el peor caso da 0, salvo ruido de punto flotante.
      expect(r.width).toBeGreaterThan(-1e-9);
      expect(r.height).toBeGreaterThan(-1e-9);
    }
  });

  it("un gapRatio negativo o NaN se trata como 0, tanto horizontal como verticalmente", () => {
    const withZero = gridLayout(slide16x9, 2, 2, 0);
    const withNegative = gridLayout(slide16x9, 2, 2, -1);
    const withNaN = gridLayout(slide16x9, 2, 2, NaN);
    for (let i = 0; i < 4; i++) {
      expectRect(withNegative[i]!, withZero[i]!);
      expectRect(withNaN[i]!, withZero[i]!);
    }
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

  it("funciona igual en una diapositiva vertical (no asume 16:9)", () => {
    const rects = chevronChainLayout(portrait, 3);
    const area = contentArea(portrait);
    const height = portrait.height * 0.12;
    const width = area.width / 3;
    const top = area.top + (area.height - height) / 2;
    expect(rects).toHaveLength(3);
    for (const r of rects) {
      expect(r.width).toBeCloseTo(width, 9);
      expect(r.height).toBeCloseTo(height, 9);
      expect(r.top).toBeCloseTo(top, 9);
    }
    expect(rects[0]!.left).toBeCloseTo(area.left, 9);
    expect(rects[2]!.left).toBeCloseTo(area.left + 2 * width, 9);
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

  it("acota el diámetro al ancho de la franja en una diapositiva estrecha, para que no se solapen", () => {
    // Diapositiva estrecha + el máximo de 12 círculos: la franja (slot) mide justo lo mismo que el
    // diámetro sin acotar (7%/12 == (1 - 2*8%)/12), así que aquí el acotado coincide con el límite:
    // sirve para comprobar que nunca se supera.
    const narrow: SlideSize = { width: 300, height: 900 };
    const rects = numberedCirclesLayout(narrow, 12);
    const area = contentArea(narrow);
    const slot = area.width / 12;

    expect(rects).toHaveLength(12);
    for (const r of rects) {
      expect(r.width).toBeLessThanOrEqual(slot + 1e-9);
      expect(r.width).toBeCloseTo(slot, 9);
      expect(r.height).toBeCloseTo(slot, 9);
    }
    // Se tocan exactamente (diameter === slot): el siguiente empieza donde termina el anterior.
    for (let i = 1; i < rects.length; i++) {
      expect(rects[i]!.left).toBeCloseTo(rects[i - 1]!.left + rects[i - 1]!.width, 9);
    }
  });
});

describe("stickerRect", () => {
  it("coloca una etiqueta 18% x 7% en la esquina superior derecha, con margen = 3% del lado menor", () => {
    // min(960,540)=540; margin=540*0.03=16.2; width=172.8; height=37.8; left=960-16.2-172.8=771
    expectRect(stickerRect(slide16x9), { left: 771, top: 16.2, width: 172.8, height: 37.8 });
  });

  it("usa el mismo margen absoluto (del lado menor) en una diapositiva vertical", () => {
    // min(540,960)=540 (igual que en slide16x9): margin=16.2 en ambos ejes también aquí.
    const width = portrait.width * 0.18;
    const height = portrait.height * 0.07;
    expectRect(stickerRect(portrait), { left: portrait.width - 16.2 - width, top: 16.2, width, height });
  });
});

describe("BASIC_SHAPE_RECT", () => {
  it("devuelve un cuadrado centrado, con lado 25% del lado menor de la diapositiva", () => {
    // min(960,540)=540; side=135; left=(960-135)/2=412.5; top=(540-135)/2=202.5
    expectRect(BASIC_SHAPE_RECT(slide16x9), { left: 412.5, top: 202.5, width: 135, height: 135 });
  });

  it("usa el ancho cuando es el lado menor", () => {
    const side = portrait.width * 0.25;
    expectRect(BASIC_SHAPE_RECT(portrait), { left: (portrait.width - side) / 2, top: (portrait.height - side) / 2, width: side, height: side });
  });
});
