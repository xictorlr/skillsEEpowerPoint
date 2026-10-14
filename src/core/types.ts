export interface Rect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface ShapeModel extends Rect {
  id: string;
  name: string;
  type: string;
  rotation: number;
  zOrder: number;
  isGroupChild: boolean;
}

export interface SlideSize {
  width: number;
  height: number;
}

export interface SelectionContext {
  /** En el orden devuelto por getSelectedShapes(): el primero es la referencia. */
  shapes: ShapeModel[];
  slide: SlideSize;
}

/** Cambio parcial de geometría para una forma. Solo se incluyen los campos modificados. */
export type Patch = { id: string } & Partial<Rect>;
