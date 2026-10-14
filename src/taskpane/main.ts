import { Panel } from "./panel";

async function selectedShapeCount(): Promise<number> {
  try {
    return await PowerPoint.run(async (context) => {
      const count = context.presentation.getSelectedShapes().getCount();
      await context.sync();
      return count.value;
    });
  } catch {
    return 0;
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

    const refresh = async () => panel.setSelectionCount(await selectedShapeCount());
    void refresh();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => void refresh(), (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.warn("[slide-tools] no se pudo suscribir al cambio de selección:", result.error.message);
      }
    });
  });
}
