# Fase 7: Robustez y tests de interfaz — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Cerrar los puntos pendientes señalados por las revisiones finales de las fases 1–6: tests DOM del panel, menos round-trips en la agenda y en los comandos de navegación, lectura por diapositiva en la comprobación de cuadros vacíos, tamaño de diapositiva sin caché cuando es exacto, y ocultar los comandos que requieren API 1.10 cuando no hay soporte.

**Architecture:** Sin cambios estructurales. Se añade `happy-dom` como entorno de test por archivo (`// @vitest-environment happy-dom`) solo para `tests/taskpane/panel.test.ts`; el resto sigue en Node.

---

### Task A: Tests DOM del panel

**Files:** `package.json` (devDependency `happy-dom`), `tests/taskpane/panel.test.ts` (new), `tests/taskpane/domHelpers.ts` (new).

- Instalar `happy-dom` (última 15.x/16.x compatible con Vitest 5). El archivo de test empieza con `// @vitest-environment happy-dom`.
- Stub de `Office`/`PowerPoint` con `vi.stubGlobal` como en `tests/commands/*`; `executeCommand` se mockea con `vi.mock("../../src/commands/execute", ...)` para devolver resultados controlados; `getThemeColors` mockeado.
- Casos:
  1. `render()` crea una sección por `SECTION_ORDER` con sus botones (ids como `data-command-id`; añadir ese atributo en `Panel.button` si no existe).
  2. `setSelectionCount(0)` deshabilita los comandos con `minShapes ≥ 1` y habilita los de 0; `null` habilita todo; el título lleva el motivo.
  3. Clipboard: sin formato copiado, los `paste-*` están deshabilitados con "Primero copia el formato de una forma."; tras `setClipboard(...)` + `setSelectionCount(1)` se habilitan.
  4. Target de color "font" deshabilita `transparency-up/down` con el motivo correcto.
  5. Busy: un clic ejecuta; un segundo clic durante la ejecución no llama a `executeCommand`; `aria-busy` se pone y se quita; `requestRefresh` se llama al terminar.
  6. Confirmación: primer clic en `check-empty-text-deck` no ejecuta y pone `is-armed` + título "Pulsa de nuevo para confirmar"; segundo clic ejecuta; clic en otro comando desarma; con `vi.useFakeTimers` el temporizador de 5 s desarma.
  7. Inspector (expertSection): con `setSummary({ bbox })` los campos se rellenan; campo vacío + `change` restaura el valor y no ejecuta; valor válido ejecuta `inspector-apply` con `{ inspector: { left: n } }`.
  8. Tema: `setSummary` con `slideId` nuevo llama a `getThemeColors` una vez; dos cambios rápidos aplican solo el último (token); `supportsApi` falso no llama.
  9. Recientes: `applyColor` exitoso con `changed > 0` añade a recientes y mantiene el foco en la fila.

Commit: `test(taskpane): tests DOM del panel con happy-dom`.

### Task B: Menos round-trips

**Files:** `src/commands/types.ts`, `src/commands/execute.ts`, `src/commands/navigation.ts`, `src/commands/expert.ts`, `src/commands/check.ts`, `src/commands/wizards.ts`, `src/office/deck.ts`, tests.

1. `ActionCommand.needsSelection?: boolean` (por defecto `true`). En `execute.ts`, si `cmd.kind === "action" && cmd.needsSelection === false`, no llamar a `readSelection`: pasar `selection = { shapes: [], slide: DEFAULT_SLIDE_SIZE }`, `proxies = new Map()`. Marcar `needsSelection: false` en: `slide-*`, `select-kind`, `select-by-name`, `check-*`, `new-presentation`, `export-selected-slides`, `agenda-create`. `disabledReason` se sigue evaluando con `count` = 0 (todos tienen `minShapes: 0`). Tests: `slide-next` hace 2 syncs en total (nav read + write) en lugar de 3.
2. `check-empty-text-slide`: usar `readSlideNav` + un nuevo `readSlideShapesWithProxies(context, slideId)` en `deck.ts` (una lectura de `slide.shapes` con `SHAPE_PROPS`, segundo sync solo si hay placeholders) en lugar de `readDeckShapes`. Tests: con 2 diapositivas, solo se cargan las formas de la activa.
3. Agenda en lotes: nuevo `addSlidesAfter(context, index, count, layout?)` en `deck.ts`: encola `count` × `slides.add(...)` y `getCount()`, un sync, obtiene los `count` últimos (`getItemAt`), encola `moveTo(index + 1 + i)` en orden, devuelve los `Slide[]`. `agenda-create` usa esta función (1 sync de alta + 1 final = 2 syncs además de `readSlideNav` y `findBlankLayout`). Si `layout` es undefined, `deletePlaceholders` por diapositiva (acepta el coste). Tests: 3 secciones → `slides.add` ×3 antes del primer sync, `moveTo` 1,2,3.

Commit: `perf(commands): menos syncs en navegación, comprobaciones y agenda`.

### Task C: Tamaño exacto sin caché y comandos ocultos sin API 1.10

**Files:** `src/office/slideSize.ts`, `src/commands/registry.ts`, `src/taskpane/panel.ts`, `src/taskpane/sections/settingsSection.ts`, README, tests.

1. `getSlideSize`: cuando `supportsApi("1.10")`, no usar la caché (lectura de un sync barata) → el cambio de Diseño > Tamaño de diapositiva se refleja de inmediato. Mantener caché solo para la heurística PNG. Actualizar JSDoc, README (quitar la limitación de caché en API 1.10) y tests.
2. `registry.ts`: `requiresApi?: string` en `BaseCommand`; `export-selected-slides` → `"1.10"`. `commandsBySection` excluye los comandos cuyo `requiresApi` no cumple `supportsApi` (import desde office/capabilities; en tests Node sin `Office` se excluyen; añadir un test con `vi.stubGlobal("Office", …)` que los incluye). `disabledReason` devuelve "Requiere PowerPoint reciente (API X)." si se ejecutara igualmente.
3. README: nota de que los comandos que requieren API 1.10 no aparecen en versiones antiguas. Versión 1.5.0 / 1.5.0.0.

Commits: `fix(office): tamaño de diapositiva sin caché con API 1.10 y comandos ocultos sin soporte` y `docs: README Fase 7 y versión 1.5.0`.

---

## Autorrevisión

Cubre las recomendaciones pendientes: tests DOM (F2/F4/F5), N+1 syncs de la agenda (F5), `readSelection` innecesaria en acciones (F4), lectura del mazo para una sola diapositiva (F5), caché del tamaño con 1.10 (F4), ocultar comandos sin API (F5 §6). Tipos: `needsSelection`/`requiresApi` en `BaseCommand`/`ActionCommand`; `addSlidesAfter` devuelve `PowerPoint.Slide[]` consumido por `agenda-create`.
