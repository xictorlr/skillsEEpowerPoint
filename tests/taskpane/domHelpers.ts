/// <reference lib="dom" />
import { Panel } from "../../src/taskpane/panel";

/** Construye el DOM mínimo que espera `Panel` (ver `src/taskpane/index.html` / `main.ts`):
 *  un `#notice` oculto y un `<main id="panel">` vacío. */
export function setupPanelDom(): { root: HTMLElement; notice: HTMLElement } {
  document.body.innerHTML = '<div id="notice" hidden></div><main id="panel"></main>';
  const root = document.getElementById("panel");
  const notice = document.getElementById("notice");
  if (!root || !notice) throw new Error("setupPanelDom: #panel/#notice no encontrados");
  return { root, notice };
}

/** Crea y renderiza un `Panel` sobre un DOM nuevo. `requestRefresh` por defecto es un no-op espiable
 *  si el test necesita comprobar que se llamó. */
export function createPanel(requestRefresh: () => void = () => undefined): {
  panel: Panel;
  root: HTMLElement;
  notice: HTMLElement;
} {
  const { root, notice } = setupPanelDom();
  const panel = new Panel(root, notice, requestRefresh);
  panel.render();
  return { panel, root, notice };
}

/** Busca el botón de un comando por su `data-command-id` (ver `Panel.button`). Lanza si no existe. */
export function getCommandButton(root: ParentNode, commandId: string): HTMLButtonElement {
  const el = root.querySelector<HTMLButtonElement>(`[data-command-id="${commandId}"]`);
  if (!el) throw new Error(`getCommandButton: no se encontró el botón del comando "${commandId}"`);
  return el;
}

/** Deja que las microtareas pendientes (la cadena `then`/`finally` de una promesa ya resuelta,
 *  p. ej. el mock de `executeCommand`) terminen de propagarse. Usa únicamente `Promise.resolve()`
 *  (nunca `setTimeout`) para que siga funcionando bajo `vi.useFakeTimers()`. */
export async function flushMicrotasks(): Promise<void> {
  for (let i = 0; i < 10; i++) await Promise.resolve();
}

/** Promesa controlable manualmente desde el test, para simular una respuesta asíncrona en vuelo
 *  (p. ej. `executeCommand`/`getThemeColors` mientras el panel está "busy" o leyendo el tema). */
export function deferred<T>(): { promise: Promise<T>; resolve: (value: T) => void; reject: (reason?: unknown) => void } {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}
