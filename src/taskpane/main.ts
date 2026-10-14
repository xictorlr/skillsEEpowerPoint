Office.onReady((info) => {
  const panel = document.getElementById("panel");
  if (!panel) return;
  panel.textContent =
    info.host === Office.HostType.PowerPoint ? "Listo en PowerPoint." : "Este panel requiere PowerPoint.";
});
