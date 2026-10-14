/// <reference types="office-js" />
/**
 * Mock mínimo de los proxies de PowerPoint que usan selection.ts, apply.ts y format.ts.
 * Registra escrituras en `writes` (geometría) y `formatWrites` (fill/lineFormat/font), y cuenta
 * llamadas a sync.
 */
export interface MockShapeInit {
  id: string;
  name?: string;
  type?: string;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
  zOrderPosition?: number;
  level?: number;
  fill?: { type: string; foregroundColor: string; transparency: number | null };
  line?: {
    visible: boolean | null;
    color: string | null;
    weight: number | null;
    dashStyle: string | null;
    transparency: number | null;
  };
  /** `undefined` = forma de texto con valores por defecto; `null` = sin marco de texto (p. ej. una imagen). */
  font?:
    | {
        name: string | null;
        size: number | null;
        color: string | null;
        bold: boolean | null;
        italic: boolean | null;
        underline: string | null;
      }
    | null;
}

export interface MockFill {
  type: string;
  foregroundColor: string;
  transparency: number | null;
  load: () => void;
  setSolidColor: (c: string) => void;
  clear: () => void;
}

export interface MockLineFormat {
  visible: boolean | null;
  color: string | null;
  weight: number | null;
  dashStyle: string | null;
  transparency: number | null;
  load: () => void;
}

export type MockFont = Record<string, unknown> & { load: (props: string) => void };

const DEFAULT_FILL = { type: "Solid", foregroundColor: "#FF0000", transparency: 0 };
const DEFAULT_LINE = { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", transparency: 0 };
const DEFAULT_FONT = { name: "Calibri", size: 18, color: "#000000", bold: false, italic: false, underline: "None" };

export class MockShape {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation: number;
  zOrderPosition: number;
  level: number;
  writes: Array<{ prop: string; value: number }> = [];
  formatWrites: Array<{ path: string; value: unknown }> = [];
  fontLoaded = false;

  fill: MockFill;
  lineFormat: MockLineFormat;

  constructor(init: MockShapeInit) {
    this.id = init.id;
    this.name = init.name ?? init.id;
    this.type = init.type ?? "GeometricShape";
    this.left = init.left;
    this.top = init.top;
    this.width = init.width;
    this.height = init.height;
    this.rotation = init.rotation ?? 0;
    this.zOrderPosition = init.zOrderPosition ?? 0;
    this.level = init.level ?? 0;
    for (const prop of ["left", "top", "width", "height"] as const) {
      let value = this[prop];
      Object.defineProperty(this, prop, {
        get: () => value,
        set: (v: number) => {
          value = v;
          this.writes.push({ prop, value: v });
        },
      });
    }

    // --- fill ---
    const fillInit = init.fill ?? DEFAULT_FILL;
    const fill: MockFill = {
      type: fillInit.type,
      foregroundColor: fillInit.foregroundColor,
      transparency: fillInit.transparency,
      load: () => {
        /* no-op: la carga real pasa por collection.load(FORMAT_LOAD_PROPS) */
      },
      setSolidColor: (c: string) => {
        fill.type = "Solid";
        fill.foregroundColor = c;
        this.formatWrites.push({ path: "fill.setSolidColor", value: c });
      },
      clear: () => {
        fill.type = "NoFill";
        this.formatWrites.push({ path: "fill.clear", value: null });
      },
    };
    let fillTransparency = fillInit.transparency;
    Object.defineProperty(fill, "transparency", {
      get: () => fillTransparency,
      set: (v: number | null) => {
        fillTransparency = v;
        this.formatWrites.push({ path: "fill.transparency", value: v });
      },
      enumerable: true,
    });
    this.fill = fill;

    // --- lineFormat ---
    const lineInit = init.line ?? DEFAULT_LINE;
    const lineFormat = {
      load: () => {
        /* no-op */
      },
    } as MockLineFormat;
    for (const key of ["visible", "color", "weight", "dashStyle", "transparency"] as const) {
      let value: unknown = lineInit[key];
      Object.defineProperty(lineFormat, key, {
        get: () => value,
        set: (v: unknown) => {
          value = v;
          this.formatWrites.push({ path: `lineFormat.${key}`, value: v });
        },
        enumerable: true,
      });
    }
    this.lineFormat = lineFormat;

    // --- textFrame / font ---
    if (init.font !== null) {
      const fontInit = init.font ?? DEFAULT_FONT;
      const font = {
        load: (_props: string) => {
          this.fontLoaded = true;
        },
      } as MockFont;
      for (const key of ["name", "size", "color", "bold", "italic", "underline"] as const) {
        let value: unknown = fontInit[key];
        Object.defineProperty(font, key, {
          get: () => value,
          set: (v: unknown) => {
            value = v;
            this.formatWrites.push({ path: `font.${key}`, value: v });
          },
          enumerable: true,
        });
      }
      const textFrame = { textRange: { font } };
      Object.defineProperty(this, "textFrame", {
        get: () => textFrame,
        enumerable: true,
      });
    } else {
      Object.defineProperty(this, "textFrame", {
        get: () => {
          throw new Error("InvalidArgument");
        },
        enumerable: true,
      });
    }
  }
}

export class MockContext {
  syncCount = 0;
  /** Cadena de propiedades pasada al último `load()` de la colección de formas. */
  loadedProps?: string;
  /** Opciones pasadas al último `getImageAsBase64()`; queda `undefined` si nunca se llamó. */
  imageOptions?: Record<string, unknown>;
  constructor(public shapes: MockShape[], public slidePng = "") {}
  presentation = {
    getSelectedShapes: () => ({
      items: this.shapes,
      load: (props: string) => {
        this.loadedProps = props;
      },
    }),
    getSelectedSlides: () => ({
      getItemAt: () => ({
        getImageAsBase64: (options: Record<string, unknown>) => {
          this.imageOptions = options;
          return { value: this.slidePng };
        },
      }),
    }),
  };
  async sync(): Promise<void> {
    this.syncCount += 1;
  }
}
