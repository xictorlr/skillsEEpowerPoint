import { getClipboard } from "../commands/clipboard";
import { executeCommand } from "../commands/execute";
import { SECTION_ORDER, commandsBySection, disabledReason } from "../commands/registry";
import { SECTION_LABELS, type Command, type CommandOptions } from "../commands/types";
import { supportsApi } from "../office/capabilities";
import type { CommandResult } from "../office/runtime";
import type { SelectionSummary } from "../office/summary";
import { getThemeColors } from "../office/theme";
import { iconSvg } from "./icons";
import { ColorSection } from "./sections/colorSection";
import { ExpertSection } from "./sections/expertSection";
import { FormatSection } from "./sections/formatSection";
import { NavigationSection } from "./sections/navigationSection";
import { SettingsSection } from "./sections/settingsSection";
import type { SectionContext } from "./sections/types";
import { toCommandOptions } from "./sanitize";
import { loadPersistedOptions, saveOptions, type PersistedOptions } from "./settings";

/** La transparencia solo tiene sentido para relleno/línea; con destino "texto" no aplica. */
const TRANSPARENCY_COMMAND_IDS: ReadonlySet<string> = new Set(["transparency-up", "transparency-down"]);
const TRANSPARENCY_FONT_REASON = "La transparencia no aplica al texto.";

export class Panel {
  private persisted: PersistedOptions = loadPersistedOptions();
  private options: CommandOptions = toCommandOptions(this.persisted);
  private buttons = new Map<string, { el: HTMLButtonElement; cmd: Command }>();
  private noticeTimer: number | undefined;
  private lastCount: number | null = null;
  private busy = false;
  private colorSection: ColorSection | undefined;
  private navigationSection: NavigationSection | undefined;
  private expertSection: ExpertSection | undefined;
  /** Diapositiva activa del último `setSummary` con resumen no nulo; `undefined` antes del primer
   *  resumen, para forzar la primera lectura de colores del tema aunque esa diapositiva sea `null`. */
  private lastSlideId: string | null | undefined;
  /** Token del último `refreshThemeColors` en curso: descarta una respuesta tardía si mientras
   *  tanto se lanzó otra lectura (p. ej. dos cambios de diapositiva seguidos). */
  private themeSeq = 0;

  constructor(
    private root: HTMLElement,
    private notice: HTMLElement,
    /** Vuelve a pedir un resumen de selección (debounce incluido); ver `main.ts`. Se llama tras
     *  cualquier comando para que Navigation/Expert/Color reflejen su efecto. */
    private requestRefresh: () => void,
  ) {
    notice.setAttribute("role", "status");
    notice.setAttribute("aria-live", "polite");
  }

  render(): void {
    this.root.replaceChildren();
    const ctx = this.sectionContext();
    const formatSection = new FormatSection(ctx);
    const colorSection = new ColorSection(ctx);
    const navigationSection = new NavigationSection(ctx);
    const expertSection = new ExpertSection(ctx);
    const settingsSection = new SettingsSection(ctx);
    this.colorSection = colorSection;
    this.navigationSection = navigationSection;
    this.expertSection = expertSection;
    SECTION_ORDER.forEach((section, i) => {
      if (i > 0) this.root.appendChild(Object.assign(document.createElement("hr"), { className: "st-divider" }));
      const sec = document.createElement("section");
      sec.className = "st-section";
      const title = document.createElement("div");
      title.className = "st-section-title";
      title.textContent = SECTION_LABELS[section];
      sec.appendChild(title);
      const row = document.createElement("div");
      row.className = "st-row";
      for (const cmd of commandsBySection(section)) row.appendChild(this.button(cmd));
      sec.appendChild(row);
      if (section === "format") sec.append(...formatSection.render());
      if (section === "color") sec.append(...colorSection.render());
      if (section === "navigation") sec.append(...navigationSection.render());
      if (section === "expert") sec.append(...expertSection.render());
      this.root.appendChild(sec);
    });
    this.root.appendChild(Object.assign(document.createElement("hr"), { className: "st-divider" }));
    this.root.appendChild(settingsSection.render());
    this.setSelectionCount(0);
  }

  /**
   * Habilita/deshabilita botones según el número de formas seleccionadas.
   * `null` significa que no se pudo determinar la selección: se habilita todo y
   * `executeCommand` revalida la selección real al ejecutar.
   * No toca los botones mientras hay un comando en curso (`busy`); el estado se
   * reaplica al terminar ese comando.
   */
  setSelectionCount(count: number | null): void {
    this.lastCount = count;
    if (this.busy) return;
    const hasClipboard = getClipboard() !== null;
    const transparencyBlocked = this.persisted.colorTarget === "font";
    const noPasteParts = Object.values(this.persisted.pasteParts).every((v) => !v);
    for (const { el, cmd } of this.buttons.values()) {
      let override: string | undefined;
      if (transparencyBlocked && TRANSPARENCY_COMMAND_IDS.has(cmd.id)) override = TRANSPARENCY_FONT_REASON;
      else if (cmd.id === "paste-format" && noPasteParts) override = "Marca al menos una parte.";
      const reason = count === null ? override : (disabledReason(cmd, count, hasClipboard) ?? override);
      el.disabled = reason !== undefined;
      el.title = reason ? `${cmd.label} — ${reason}` : cmd.label;
    }
  }

  /**
   * Recibe el resumen de selección en vivo (ver `office/summary.ts`): actualiza el conteo de
   * formas (como `setSelectionCount`), el indicador de diapositiva/campo "Ir a" de Navigation y el
   * inspector de Expert. Cuando la diapositiva activa cambia respecto a la última vez, también
   * vuelve a pedir los colores del tema (otro `PowerPoint.run`, ver `office/theme.ts`) para la fila
   * "Tema" de Color.
   */
  setSummary(summary: SelectionSummary | null): void {
    this.setSelectionCount(summary?.count ?? null);
    this.navigationSection?.update(summary);
    this.expertSection?.update(summary);
    // Sin resumen (no se pudo leer la selección) no sabemos si la diapositiva cambió: se deja
    // `lastSlideId` como está para no perder la próxima comparación válida.
    if (summary && summary.slideId !== this.lastSlideId) {
      this.lastSlideId = summary.slideId;
      void this.refreshThemeColors();
    }
  }

  private async refreshThemeColors(): Promise<void> {
    const id = ++this.themeSeq;
    // Sin PowerPointApi 1.10 `getThemeColors` siempre devolvería null (ver office/theme.ts): se
    // evita el `PowerPoint.run` de más.
    if (!supportsApi("1.10")) {
      this.colorSection?.setThemeColors(null);
      return;
    }
    try {
      const colors = await PowerPoint.run((context) => getThemeColors(context));
      if (id === this.themeSeq) this.colorSection?.setThemeColors(colors);
    } catch {
      if (id === this.themeSeq) this.colorSection?.setThemeColors(null);
    }
  }

  /** Aplica un parche a las opciones persistidas, recomputa `options` y guarda. */
  private updatePersisted(patch: Partial<PersistedOptions>): void {
    this.persisted = { ...this.persisted, ...patch };
    this.options = toCommandOptions(this.persisted);
    saveOptions(this.persisted);
  }

  showNotice(text: string, isError = false): void {
    this.notice.textContent = text;
    this.notice.classList.toggle("is-error", isError);
    this.notice.hidden = false;
    window.clearTimeout(this.noticeTimer);
    this.noticeTimer = window.setTimeout(() => (this.notice.hidden = true), isError ? 6000 : 3000);
  }

  private button(cmd: Command): HTMLButtonElement {
    const el = document.createElement("button");
    el.className = "st-btn";
    el.type = "button";
    el.innerHTML = iconSvg(cmd.icon);
    el.setAttribute("aria-label", cmd.label);
    el.addEventListener("click", () => void this.run(cmd));
    this.buttons.set(cmd.id, { el, cmd });
    return el;
  }

  private async run(cmd: Command, extra?: Partial<CommandOptions>): Promise<CommandResult | undefined> {
    if (this.busy) return undefined;
    this.busy = true;
    for (const { el } of this.buttons.values()) el.disabled = true;
    this.root.setAttribute("aria-busy", "true");
    try {
      const options = extra ? { ...this.options, ...extra } : this.options;
      const result = await executeCommand(cmd, options);
      if (!result.ok) this.showNotice(result.message, true);
      else if (result.message) this.showNotice(result.message);
      else if (result.changed === 0) this.showNotice("Nada que cambiar.");
      return result;
    } finally {
      this.busy = false;
      this.root.removeAttribute("aria-busy");
      this.setSelectionCount(this.lastCount);
      // Cualquier comando puede haber movido/reescalado/renombrado/reordenado la selección o
      // cambiado de diapositiva: se pide un resumen nuevo para que Navigation/Expert/Color (vía
      // setSummary) lo reflejen, no solo tras un DocumentSelectionChanged del usuario.
      this.requestRefresh();
    }
  }

  /** Construye el contexto que ven las secciones: getters para que siempre lean el estado vigente. */
  private sectionContext(): SectionContext {
    const self = this;
    return {
      get persisted(): PersistedOptions {
        return self.persisted;
      },
      updatePersisted: (patch) => self.updatePersisted(patch),
      run: (cmd, extra) => self.run(cmd, extra),
      refreshButtons: () => self.setSelectionCount(self.lastCount),
      showNotice: (text, isError) => self.showNotice(text, isError),
    };
  }
}
