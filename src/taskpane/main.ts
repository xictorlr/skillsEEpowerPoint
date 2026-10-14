if (typeof Office === "undefined") {
  const panel = document.getElementById("panel");
  if (panel) {
    panel.textContent = "No se pudo cargar Office.js. Revisa la conexión y vuelve a abrir el panel.";
  }
} else {
  Office.onReady((info) => {
    const panel = document.getElementById("panel");
    if (!panel) return;
    panel.textContent =
      info.host === Office.HostType.PowerPoint ? "Listo en PowerPoint." : "Este panel requiere PowerPoint.";
  });
}
