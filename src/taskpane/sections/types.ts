import type { Command, CommandOptions } from "../../commands/types";
import type { CommandResult } from "../../office/runtime";
import type { PersistedOptions } from "../sanitize";

/** Lo que una sección del panel necesita del núcleo. Las secciones no tocan el DOM fuera de lo que devuelven. */
export interface SectionContext {
  readonly persisted: PersistedOptions;
  updatePersisted(patch: Partial<PersistedOptions>): void;
  run(cmd: Command, extra?: Partial<CommandOptions>): Promise<CommandResult | undefined>;
  /** Vuelve a calcular el estado habilitado/deshabilitado de todos los botones. */
  refreshButtons(): void;
  showNotice(text: string, isError?: boolean): void;
}
