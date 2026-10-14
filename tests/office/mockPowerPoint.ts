/// <reference types="office-js" />
/**
 * Mock mínimo de los proxies de PowerPoint que usan selection.ts y apply.ts.
 * Registra escrituras en `writes` y cuenta llamadas a sync.
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
}

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
  }
}

export class MockContext {
  syncCount = 0;
  constructor(public shapes: MockShape[], public slidePng = "") {}
  presentation = {
    getSelectedShapes: () => ({ items: this.shapes, load: () => undefined }),
    getSelectedSlides: () => ({
      getItemAt: () => ({ getImageAsBase64: () => ({ value: this.slidePng }) }),
    }),
  };
  async sync(): Promise<void> {
    this.syncCount += 1;
  }
}
