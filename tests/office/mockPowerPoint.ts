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
  /** Contenido de un marcador de posición ("Placeholder"); `undefined` se trata como `null` (texto). */
  containedType?: string | null;
  /** `false` simula una forma que Office.js rechaza al tocar `fill`/`lineFormat` (ver `MockContext`). */
  supportsFormat?: boolean;
  /** Propiedades de `textFrame`/`textRange`/`paragraphFormat` para formas con marco de texto. */
  textProps?: {
    text: string;
    horizontalAlignment: string | null;
    verticalAlignment: string;
    autoSizeSetting: string;
    wordWrap: boolean;
    leftMargin: number;
    rightMargin: number;
    topMargin: number;
    bottomMargin: number;
  };
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
const DEFAULT_TEXT_PROPS = {
  text: "hola mundo",
  horizontalAlignment: "Left" as string | null,
  verticalAlignment: "Top",
  autoSizeSetting: "AutoSizeNone",
  wordWrap: true,
  leftMargin: 7.2,
  rightMargin: 7.2,
  topMargin: 3.6,
  bottomMargin: 3.6,
};

export class MockShape {
  id: string;
  /** Definida vía `Object.defineProperty` en el constructor (para registrar escrituras en `formatWrites`). */
  name!: string;
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
  fillLoaded = false;
  lineLoaded = false;
  textFrameLoaded = false;
  textRangeLoaded = false;
  /** Cada `load(props)` llamado sobre textFrame/textRange/paragraphFormat/font, con su ruta y props. */
  loads: Array<{ path: string; props: string }> = [];
  supportsFormat: boolean;
  /** Referencia al contexto que la contiene; la fija `MockContext` al construirse. */
  ctx?: MockContext;
  /** `shape.group.ungroup()`: registra `{ path: "ungroup", value: null }` en `formatWrites`. */
  readonly group = {
    ungroup: () => {
      this.formatWrites.push({ path: "ungroup", value: null });
    },
  };

  constructor(init: MockShapeInit) {
    this.id = init.id;
    this.type = init.type ?? "GeometricShape";
    this.left = init.left;
    this.top = init.top;
    this.width = init.width;
    this.height = init.height;
    this.rotation = init.rotation ?? 0;
    this.zOrderPosition = init.zOrderPosition ?? 0;
    this.level = init.level ?? 0;
    this.supportsFormat = init.supportsFormat ?? true;
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

    // --- name ---
    let nameValue = init.name ?? init.id;
    Object.defineProperty(this, "name", {
      get: () => nameValue,
      set: (v: string) => {
        nameValue = v;
        this.formatWrites.push({ path: "name", value: v });
      },
      enumerable: true,
    });

    // --- fill ---
    const fillInit = init.fill ?? DEFAULT_FILL;
    const fillImpl: MockFill = {
      type: fillInit.type,
      foregroundColor: fillInit.foregroundColor,
      transparency: fillInit.transparency,
      load: () => {
        this.fillLoaded = true;
      },
      setSolidColor: (c: string) => {
        fillImpl.type = "Solid";
        fillImpl.foregroundColor = c;
        this.formatWrites.push({ path: "fill.setSolidColor", value: c });
      },
      clear: () => {
        fillImpl.type = "NoFill";
        this.formatWrites.push({ path: "fill.clear", value: null });
      },
    };
    let fillTransparency = fillInit.transparency;
    Object.defineProperty(fillImpl, "transparency", {
      get: () => fillTransparency,
      set: (v: number | null) => {
        fillTransparency = v;
        this.formatWrites.push({ path: "fill.transparency", value: v });
      },
      enumerable: true,
    });
    Object.defineProperty(this, "fill", {
      get: () => {
        if (!this.supportsFormat && this.ctx) this.ctx.invalidTouched = true;
        return fillImpl;
      },
      enumerable: true,
    });

    // --- lineFormat ---
    const lineInit = init.line ?? DEFAULT_LINE;
    const lineFormatImpl = {
      load: () => {
        this.lineLoaded = true;
      },
    } as MockLineFormat;
    for (const key of ["visible", "color", "weight", "dashStyle", "transparency"] as const) {
      let value: unknown = lineInit[key];
      Object.defineProperty(lineFormatImpl, key, {
        get: () => value,
        set: (v: unknown) => {
          value = v;
          this.formatWrites.push({ path: `lineFormat.${key}`, value: v });
        },
        enumerable: true,
      });
    }
    Object.defineProperty(this, "lineFormat", {
      get: () => {
        if (!this.supportsFormat && this.ctx) this.ctx.invalidTouched = true;
        return lineFormatImpl;
      },
      enumerable: true,
    });

    // --- placeholderFormat ---
    const placeholderFormatImpl = {
      containedType: init.containedType ?? null,
      load: () => {
        /* no-op: el valor ya está disponible en el mock */
      },
    };
    Object.defineProperty(this, "placeholderFormat", {
      get: () => placeholderFormatImpl,
      enumerable: true,
    });

    // --- textFrame / font ---
    if (init.font !== null) {
      const fontInit = init.font ?? DEFAULT_FONT;
      const font = {
        load: (props: string) => {
          this.fontLoaded = true;
          this.loads.push({ path: "font", props });
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
      const textInit = { ...DEFAULT_TEXT_PROPS, ...init.textProps };

      const paragraphFormat = {
        load: (props: string) => {
          this.loads.push({ path: "paragraphFormat", props });
        },
      } as { horizontalAlignment: string | null; load: (props: string) => void };
      let horizontalAlignment: string | null = textInit.horizontalAlignment;
      Object.defineProperty(paragraphFormat, "horizontalAlignment", {
        get: () => horizontalAlignment,
        set: (v: string | null) => {
          horizontalAlignment = v;
          this.formatWrites.push({ path: "paragraphFormat.horizontalAlignment", value: v });
        },
        enumerable: true,
      });

      let textValue = textInit.text;
      const textRange = {
        font,
        paragraphFormat,
        load: (props: string) => {
          this.textRangeLoaded = true;
          this.loads.push({ path: "textRange", props });
        },
        getSubstring: (start: number, length: number) => {
          const substring = {} as { text: string };
          Object.defineProperty(substring, "text", {
            get: () => textValue.slice(start, start + length),
            set: (v: string) => {
              textValue = textValue.slice(0, start) + v + textValue.slice(start + length);
              this.formatWrites.push({ path: "textRange.substring", value: { start, length, text: v } });
            },
            enumerable: true,
          });
          return substring;
        },
      } as {
        text: string;
        font: MockFont;
        paragraphFormat: typeof paragraphFormat;
        load: (props: string) => void;
        getSubstring: (start: number, length: number) => { text: string };
      };
      Object.defineProperty(textRange, "text", {
        get: () => textValue,
        set: (v: string) => {
          textValue = v;
          this.formatWrites.push({ path: "textRange.text", value: v });
        },
        enumerable: true,
      });

      const textFrame = {
        textRange,
        load: (props: string) => {
          this.textFrameLoaded = true;
          this.loads.push({ path: "textFrame", props });
        },
      } as Record<string, unknown> & { textRange: typeof textRange; load: (props: string) => void };
      for (const key of ["verticalAlignment", "autoSizeSetting", "wordWrap", "leftMargin", "rightMargin", "topMargin", "bottomMargin"] as const) {
        let value: unknown = textInit[key];
        Object.defineProperty(textFrame, key, {
          get: () => value,
          set: (v: unknown) => {
            value = v;
            this.formatWrites.push({ path: `textFrame.${key}`, value: v });
          },
          enumerable: true,
        });
      }
      Object.defineProperty(this, "textFrame", {
        get: () => textFrame,
        enumerable: true,
      });
    } else {
      // Office.js no lanza de forma síncrona al leer la propiedad: el error real aparece recién en
      // el siguiente sync(). El mock imita eso marcando el contexto en vez de lanzar aquí.
      const dummyFont = { load: () => {} } as MockFont;
      Object.defineProperty(this, "textFrame", {
        get: () => {
          if (this.ctx) this.ctx.invalidTouched = true;
          return { textRange: { font: dummyFont } };
        },
        enumerable: true,
      });
    }
  }

  /** `shape.setZOrder(op)`: registra `{ path: "setZOrder", value: op }` en `formatWrites`. */
  setZOrder(op: string): void {
    this.formatWrites.push({ path: "setZOrder", value: op });
  }
}

export interface MockSlideInit {
  id: string;
  shapes?: MockShape[];
}

/**
 * Mock mínimo de `PowerPoint.Slide`: expone las formas de la diapositiva (`shapes`), permite
 * seleccionarlas (`setSelectedShapes`) y conserva el comportamiento previo de `getItemAt(0)`
 * (`getImageAsBase64`, `themeColorScheme`), usado por `slideSize.ts`.
 */
export class MockSlide {
  id: string;
  index: number;
  shapesList: MockShape[];
  ctx: MockContext;

  constructor(ctx: MockContext, id: string, index: number, shapes: MockShape[]) {
    this.ctx = ctx;
    this.id = id;
    this.index = index;
    this.shapesList = shapes;
    for (const s of shapes) s.ctx = ctx;
  }

  get shapes() {
    return {
      items: this.shapesList,
      load: (_props: string) => {
        /* no-op: los valores ya están disponibles en el mock */
      },
    };
  }

  /** `slide.setSelectedShapes(ids)`: registra `ids` en `ctx.selectedShapeIds`. */
  setSelectedShapes(ids: string[]): void {
    this.ctx.selectedShapeIds = ids;
  }

  getImageAsBase64(options: Record<string, unknown>): { value: string } {
    this.ctx.imageOptions = options;
    return { value: this.ctx.slidePng };
  }

  get themeColorScheme() {
    return {
      getThemeColor: (name: string) => ({ value: this.ctx.opts.themeColors?.[name] ?? "" }),
    };
  }
}

export interface MockContextOptions {
  /** Si se da, `presentation.pageSetup` expone estos valores (simula PowerPointApi 1.10). */
  pageSetup?: { slideWidth: number; slideHeight: number };
  /** Mapa nombre de color de tema -> valor devuelto por `themeColorScheme.getThemeColor(name)`. */
  themeColors?: Record<string, string>;
  /** Diapositivas de la presentación. Si se omite, se sintetiza una única diapositiva que contiene `shapes`. */
  slides?: MockSlideInit[];
  /** Índice (en `slides`, o en la diapositiva sintetizada) de la diapositiva activa. Por defecto 0. */
  activeSlideIndex?: number;
}

export class MockContext {
  syncCount = 0;
  /** Cadena de propiedades pasada al último `load()` de la colección de formas. */
  loadedProps?: string;
  /** Opciones pasadas al último `getImageAsBase64()`; queda `undefined` si nunca se llamó. */
  imageOptions?: Record<string, unknown>;
  /** Se pone a true si se accede a `fill`/`lineFormat`/`textFrame` de una forma que no lo admite,
   *  o si un test quiere forzar que el próximo `sync()` falle (p. ej. error de lectura del tema);
   *  el siguiente `sync()` lo traduce en un rechazo y lo limpia. */
  invalidTouched = false;
  /** `presentation.setSelectedSlides(ids)`: ids registrados por la última llamada. */
  selectedSlideIds?: string[];
  /** `slide.setSelectedShapes(ids)`: ids registrados por la última llamada. */
  selectedShapeIds?: string[];
  /** `presentation.getSelectedShapes().group()`: true si se llamó. */
  grouped = false;
  readonly slidesList: MockSlide[];
  activeSlideIndex: number;

  constructor(public shapes: MockShape[], public slidePng = "", public opts: MockContextOptions = {}) {
    for (const s of shapes) s.ctx = this;
    this.slidesList = opts.slides
      ? opts.slides.map((init, i) => new MockSlide(this, init.id, i, init.shapes ?? []))
      : [new MockSlide(this, "slide-1", 0, shapes)];
    this.activeSlideIndex = opts.activeSlideIndex ?? 0;
  }
  /**
   * Getter (en vez de propiedad de instancia con `=`): los inicializadores de campo de clase se
   * evalúan antes que las asignaciones de "parameter properties" del constructor (p. ej. `opts`),
   * así que leer `this.opts` de forma anticipada en un campo devolvería `undefined`. Como getter,
   * se evalúa de forma perezosa en cada acceso a `ctx.presentation`, ya con el constructor terminado.
   */
  get presentation() {
    const activeSlide = this.slidesList[this.activeSlideIndex];
    return {
      getSelectedShapes: () => ({
        items: this.shapes,
        load: (props: string) => {
          this.loadedProps = props;
        },
        group: () => {
          this.grouped = true;
        },
      }),
      getSelectedSlides: () => ({
        items: activeSlide ? [activeSlide] : [],
        load: (_props: string) => {
          /* no-op: los valores ya están disponibles en el mock */
        },
        getItemAt: (_index: number) => activeSlide,
      }),
      setSelectedSlides: (ids: string[]) => {
        this.selectedSlideIds = ids;
      },
      slides: {
        items: this.slidesList,
        load: (_props: string) => {
          /* no-op: los valores ya están disponibles en el mock */
        },
        getItem: (id: string) => {
          const found = this.slidesList.find((s) => s.id === id);
          if (!found) throw { code: "ItemNotFound", message: `No se encontró la diapositiva "${id}".` };
          return found;
        },
        getCount: () => ({ value: this.slidesList.length }),
      },
      pageSetup: {
        slideWidth: this.opts.pageSetup?.slideWidth ?? 0,
        slideHeight: this.opts.pageSetup?.slideHeight ?? 0,
        load: (_props?: string) => {
          /* no-op: los valores ya están disponibles en el mock */
        },
      },
    };
  }
  async sync(): Promise<void> {
    this.syncCount += 1;
    if (this.invalidTouched) {
      this.invalidTouched = false;
      throw { code: "InvalidArgument", message: "La forma no admite esta propiedad." };
    }
  }
}
