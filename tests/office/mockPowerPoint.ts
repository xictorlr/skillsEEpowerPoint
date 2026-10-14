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
  textProps?: Partial<{
    text: string;
    horizontalAlignment: string | null;
    verticalAlignment: string;
    autoSizeSetting: string;
    wordWrap: boolean;
    leftMargin: number;
    rightMargin: number;
    topMargin: number;
    bottomMargin: number;
    /** Valor de `textFrame.hasText` (API 1.4). Por defecto `true`. */
    hasText: boolean;
    /** Valor inicial de `paragraphFormat.bulletFormat.visible`. Por defecto `true`. */
    bulletsVisible: boolean;
    /** Valor inicial de `paragraphFormat.indentLevel`. Por defecto `0`. */
    indentLevel: number;
  }>;
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

/** Layouts del primer slide master, por defecto, cuando `MockContextOptions.layouts` no se da. */
const DEFAULT_LAYOUTS: Array<{ id: string; name: string }> = [
  { id: "lay-title", name: "Title Slide" },
  { id: "lay-blank", name: "Blank" },
];

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
  bulletsVisible: true,
  indentLevel: 0,
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
      this.recordFormatWrite("ungroup", null);
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
        this.recordFormatWrite("name", v);
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
        this.recordFormatWrite("fill.setSolidColor", c);
      },
      clear: () => {
        fillImpl.type = "NoFill";
        this.recordFormatWrite("fill.clear", null);
      },
    };
    let fillTransparency = fillInit.transparency;
    Object.defineProperty(fillImpl, "transparency", {
      get: () => fillTransparency,
      set: (v: number | null) => {
        fillTransparency = v;
        this.recordFormatWrite("fill.transparency", v);
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
          this.recordFormatWrite(`lineFormat.${key}`, v);
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
            this.recordFormatWrite(`font.${key}`, v);
          },
          enumerable: true,
        });
      }
      const textInit = { ...DEFAULT_TEXT_PROPS, ...init.textProps };

      const bulletFormat = {
        load: (props: string) => {
          this.loads.push({ path: "bulletFormat", props });
        },
      } as { visible: boolean; load: (props: string) => void };
      let bulletsVisible = textInit.bulletsVisible;
      Object.defineProperty(bulletFormat, "visible", {
        get: () => bulletsVisible,
        set: (v: boolean) => {
          bulletsVisible = v;
          this.recordFormatWrite("bulletFormat.visible", v);
        },
        enumerable: true,
      });

      const paragraphFormat = {
        bulletFormat,
        load: (props: string) => {
          this.loads.push({ path: "paragraphFormat", props });
        },
      } as { horizontalAlignment: string | null; indentLevel: number; bulletFormat: typeof bulletFormat; load: (props: string) => void };
      let horizontalAlignment: string | null = textInit.horizontalAlignment;
      Object.defineProperty(paragraphFormat, "horizontalAlignment", {
        get: () => horizontalAlignment,
        set: (v: string | null) => {
          horizontalAlignment = v;
          this.recordFormatWrite("paragraphFormat.horizontalAlignment", v);
        },
        enumerable: true,
      });
      let indentLevel = textInit.indentLevel;
      Object.defineProperty(paragraphFormat, "indentLevel", {
        get: () => indentLevel,
        set: (v: number) => {
          indentLevel = v;
          this.recordFormatWrite("paragraphFormat.indentLevel", v);
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
              this.recordFormatWrite("textRange.substring", { start, length, text: v });
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
          this.recordFormatWrite("textRange.text", v);
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
            this.recordFormatWrite(`textFrame.${key}`, v);
          },
          enumerable: true,
        });
      }
      // `hasText` (API 1.4) es de solo lectura: no se escribe desde el código bajo prueba.
      Object.defineProperty(textFrame, "hasText", {
        get: () => textInit.hasText ?? true,
        enumerable: true,
      });
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

  /** `shape.load(props)` (p. ej. `"id"` sobre una forma recién creada, ver `commands/insert.ts`):
   *  registra `{ path: "shape", props }` en `loads`. Los valores ya están disponibles en el mock. */
  load(props: string): void {
    this.loads.push({ path: "shape", props });
  }

  /** `shape.setZOrder(op)`: registra `{ path: "setZOrder", value: op }` en `formatWrites`. */
  setZOrder(op: string): void {
    this.recordFormatWrite("setZOrder", op);
  }

  /** `shape.delete()`: registra `{ path: "delete", value: null }` en `formatWrites`. */
  delete(): void {
    this.recordFormatWrite("delete", null);
  }

  /**
   * Registra una escritura de formato: en `this.formatWrites` (como antes) y, si la forma
   * pertenece a un contexto, también en `ctx.calls` (con el id de la forma) para poder verificar
   * el orden de llamadas entre varias formas (p. ej. `applyZOrder` sobre varios ids).
   */
  private recordFormatWrite(path: string, value: unknown): void {
    this.formatWrites.push({ path, value });
    this.ctx?.calls.push({ id: this.id, path, value });
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

  /**
   * Crea una forma nueva de tipo `type`, la añade a la diapositiva y registra
   * `{ id, path, value }` en `ctx.calls`. Compartido por `addTextBox`, `addGeometricShape` y
   * `addLine` (mismo comportamiento; solo cambian el tipo de forma y los datos registrados en
   * `value`).
   */
  private addShape(
    type: string,
    path: string,
    value: Record<string, unknown>,
    opts: { left: number; top: number; width: number; height: number },
  ): MockShape {
    this.ctx.addedShapes += 1;
    const shape = new MockShape({ id: `new-shape-${this.ctx.addedShapes}`, type, ...opts });
    shape.ctx = this.ctx;
    this.shapesList.push(shape);
    this.ctx.calls.push({ id: shape.id, path, value });
    return shape;
  }

  get shapes() {
    return {
      items: this.shapesList,
      load: (_props: string) => {
        /* no-op: los valores ya están disponibles en el mock */
      },
      /** `shapes.addTextBox(text, opts)`: crea un `MockShape` de tipo "TextBox", lo añade a la
       *  diapositiva y registra `{ id, path: "addTextBox", value: { text, ...opts } }` en `ctx.calls`. */
      addTextBox: (text: string, opts: { left: number; top: number; width: number; height: number }): MockShape =>
        this.addShape("TextBox", "addTextBox", { text, ...opts }, opts),
      /** `shapes.addGeometricShape(type, opts)`: crea un `MockShape` de tipo "GeometricShape", lo
       *  añade a la diapositiva y registra `{ id, path: "addGeometricShape", value: { type, ...opts } }`
       *  en `ctx.calls`. */
      addGeometricShape: (type: string, opts: { left: number; top: number; width: number; height: number }): MockShape =>
        this.addShape("GeometricShape", "addGeometricShape", { type, ...opts }, opts),
      /** `shapes.addLine(kind, opts)`: crea un `MockShape` de tipo "Line", lo añade a la
       *  diapositiva y registra `{ id, path: "addLine", value: { kind, ...opts } }` en `ctx.calls`. */
      addLine: (kind: string, opts: { left: number; top: number; width: number; height: number }): MockShape =>
        this.addShape("Line", "addLine", { kind, ...opts }, opts),
    };
  }

  /** `slide.setSelectedShapes(ids)`: registra `ids` en `ctx.selectedShapeIds`. */
  setSelectedShapes(ids: string[]): void {
    this.ctx.selectedShapeIds = ids;
  }

  /** `slide.moveTo(index)`: registra `{ id: slide.id, path: "moveTo", value: index }` en `ctx.calls`. */
  moveTo(index: number): void {
    this.ctx.calls.push({ id: this.id, path: "moveTo", value: index });
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
  /** Layouts de `presentation.slideMasters.getItemAt(0).layouts`. Por defecto `DEFAULT_LAYOUTS`. */
  layouts?: Array<{ id: string; name: string }>;
  /** Id de `presentation.slideMasters.getItemAt(0)`. Por defecto `"master-1"`. */
  masterId?: string;
  /** Cantidad de marcadores de posición ("Placeholder") con los que nace cada diapositiva creada
   *  vía `slides.add()`, para poder probar `deletePlaceholders`. Por defecto 0. */
  newSlidePlaceholders?: number;
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
  /** Cada escritura de formato de cualquier `MockShape` de este contexto, en el orden en que
   *  ocurrió (incluye el id de la forma); ver `MockShape.recordFormatWrite`. Permite verificar el
   *  orden de llamadas entre varias formas (p. ej. `applyZOrder` sobre varios ids). */
  calls: Array<{ id: string; path: string; value: unknown }> = [];
  readonly slidesList: MockSlide[];
  activeSlideIndex: number;
  /** Diapositivas añadidas vía `slides.add()`, usado para generar ids únicos. */
  addedSlides = 0;
  /** Formas añadidas vía `shapes.addTextBox()`, usado para generar ids únicos. */
  addedShapes = 0;
  /** Cadena de propiedades pasada a cada `presentation.slides.load()`, en orden. */
  slideLoads: string[] = [];

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
        load: (props: string) => {
          this.slideLoads.push(props);
        },
        getItem: (id: string) => {
          const found = this.slidesList.find((s) => s.id === id);
          if (!found) throw { code: "ItemNotFound", message: `No se encontró la diapositiva "${id}".` };
          return found;
        },
        getItemAt: (index: number) => {
          const found = this.slidesList[index];
          if (!found) throw { code: "ItemNotFound", message: `No se encontró la diapositiva en el índice ${index}.` };
          return found;
        },
        getCount: () => ({ value: this.slidesList.length }),
        /** `slides.add(opts)`: registra `{ id: "presentation", path: "slides.add", value: opts }`
         *  en `ctx.calls` y crea un `MockSlide` nuevo al final (con `opts.newSlidePlaceholders`
         *  marcadores de posición, para poder probar `deletePlaceholders`); no devuelve nada,
         *  igual que en Office.js real (ver `addSlideAfter` en `deck.ts`). */
        add: (addOptions?: { layoutId?: string }) => {
          this.calls.push({ id: "presentation", path: "slides.add", value: addOptions });
          this.addedSlides += 1;
          const placeholderCount = this.opts.newSlidePlaceholders ?? 0;
          const placeholders: MockShape[] = [];
          for (let i = 0; i < placeholderCount; i += 1) {
            placeholders.push(new MockShape({ id: `new-slide-${this.addedSlides}-ph-${i + 1}`, type: "Placeholder", left: 0, top: 0, width: 1, height: 1 }));
          }
          this.slidesList.push(new MockSlide(this, `new-slide-${this.addedSlides}`, this.slidesList.length, placeholders));
        },
        exportAsBase64Presentation: (ids: string[]) => ({ value: `BASE64:${ids.join(",")}` }),
      },
      slideMasters: {
        getItemAt: (_index: number) => ({
          id: this.opts.masterId ?? "master-1",
          load: (_props: string) => {
            /* no-op: los valores ya están disponibles en el mock */
          },
          layouts: {
            items: this.opts.layouts ?? DEFAULT_LAYOUTS,
            load: (_props: string) => {
              /* no-op: los valores ya están disponibles en el mock */
            },
          },
        }),
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
