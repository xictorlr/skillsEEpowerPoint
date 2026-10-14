import { readSelectionSummary, type SelectionSummary } from "../office/summary";
import { Panel } from "./panel";

/** Lee el resumen de la selección actual (ver `office/summary.ts`), o `null` si no se pudo leer. */
async function readSummary(): Promise<SelectionSummary | null> {
  try {
    return await PowerPoint.run((context) => readSelectionSummary(context));
  } catch {
    return null;
  }
}

if (typeof Office === "undefined") {
  const panel = document.getElementById("panel");
  if (panel) {
    panel.textContent = "No se pudo cargar Office.js. Revisa la conexión y vuelve a abrir el panel.";
  }
} else {
  Office.onReady((info) => {
    const root = document.getElementById("panel");
    const notice = document.getElementById("notice");
    if (!root || !notice) return;

    if (info.host !== Office.HostType.PowerPoint) {
      root.textContent = "Este panel requiere PowerPoint.";
      return;
    }

    const panel = new Panel(root, notice);
    panel.render();

    let seq = 0;
    let timer: number | undefined;
    const refresh = async () => {
      const id = ++seq;
      const summary = await readSummary();
      if (id === seq) panel.setSummary(summary);
    };
    const scheduleRefresh = () => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => void refresh(), 120);
    };
    void refresh();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, scheduleRefresh, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.warn("[slide-tools] no se pudo suscribir al cambio de selección:", result.error.message);
      }
    });
    window.addEventListener("focus", scheduleRefresh);
  });
}
