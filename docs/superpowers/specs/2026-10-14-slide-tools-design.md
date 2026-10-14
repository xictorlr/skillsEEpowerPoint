# Diseño: Slide Tools, panel de productividad para PowerPoint (Office.js)

Fecha: 2026-10-14
Autor: Victor Lopez
Repo: https://github.com/xictorlr/skillsEEpowerPoint
Estado: aprobado por el usuario (diseño); pendiente plan de implementación

## 1. Objetivo

Construir desde cero un add-in de PowerPoint con un panel lateral de productividad, al estilo de
los paneles comerciales para presentaciones: un panel con secciones de botones que aplican
operaciones de alineación, tamaño, formato, color y texto a las formas seleccionadas, y
después el resto de secciones (Navigation & View, Expert Tools, Check, Presentation, Wizards).

Debe funcionar en PowerPoint para macOS y Windows (Microsoft 365) con el mismo código.

## 2. Contexto técnico

Los paneles de productividad comerciales para PowerPoint se construyen históricamente como add-ins
COM/VBA. En macOS, VBA no ofrece un task pane nativo: el panel lateral exige código nativo aparte,
y VBA tampoco es portable a Windows sin duplicar la UI. Por eso se elige Office.js.

## 3. Decisiones tomadas

| Decisión | Elección | Motivo |
|---|---|---|
| Alcance inicial | Herramientas de formato (Alignment, Size, Format, Color, Text) | Uso diario, viable, API suficiente |
| Alcance posterior | Resto de secciones del panel | Pedido explícito del usuario |
| Plataforma | Mac y Windows | Pedido del usuario |
| Arquitectura | Add-in web Office.js con task pane | Única opción con panel lateral en ambas plataformas |
| Lenguaje y build | TypeScript + Vite | Node 26 disponible; tipado del API; build rápido |
| UI | HTML + CSS propios, sin framework | Panel pequeño (ancho ~150-220 px); evita peso y dependencias |
| Tests | Vitest sobre lógica pura | Geometría y formato testeables sin PowerPoint |
| Hosting | GitHub Pages desde el mismo repo | Pedido del usuario; requiere repo público (cuenta gratuita) |
| Marca y autoría | Victor Lopez (nombre del proveedor en manifest, README y UI) | Sin referencias a terceros ni a productos comerciales |
| API mínima | PowerPointApi 1.8 (manifest); runtime check para 1.9/1.10 | Mac 16.110 soporta 1.10; 1.8 da z-order, grupos, tablas, fuente completa |

## 4. Arquitectura

```
Task pane (HTML/CSS/TS)
   │  clic en botón
   ▼
commands/registry  ──►  office/selection (lee selección → ShapeModel[])
                            │
                            ▼
                        core/* (lógica pura: calcula nuevo estado)
                            │
                            ▼
                        office/apply (escribe diferencias, 1 context.sync)
```

Principio: `core/` no importa Office.js. Trabaja sobre modelos planos y es 100 % testeable.
`office/` es un adaptador fino. Un comando = un `PowerPoint.run` = un `context.sync` final,
para que PowerPoint lo registre como un único paso de deshacer.

### 4.1 Modelo de datos

```ts
interface ShapeModel {
  id: string;
  name: string;
  type: string;            // PowerPoint.ShapeType
  left: number; top: number; width: number; height: number;   // puntos
  rotation: number;
  zOrder: number;
  isGroupChild: boolean;
}

interface ShapeFormat {       // para copiar/pegar formato
  fill?: { color: string; transparency: number; visible: boolean };
  line?: { color: string; weight: number; dashStyle: string; transparency: number; visible: boolean };
  font?: { name: string; size: number; color: string; bold: boolean; italic: boolean; underline: string };
  size?: { width: number; height: number };
}

interface SelectionContext {
  shapes: ShapeModel[];      // en orden de selección (orden de la colección devuelta por el API)
  slide: { width: number; height: number };
}
```

### 4.2 Estructura del repo

```
skillsEEpowerPoint/
├── manifest.xml                # producción: apunta a https://xictorlr.github.io/skillsEEpowerPoint/
├── manifest.dev.xml            # desarrollo: https://localhost:3000/
├── package.json, vite.config.ts, tsconfig.json, vitest.config.ts
├── src/
│   ├── taskpane/
│   │   ├── index.html          # panel
│   │   ├── main.ts             # Office.onReady, monta el panel, conecta eventos
│   │   ├── panel.ts            # renderiza secciones y botones desde el registro de comandos
│   │   ├── styles.css
│   │   └── icons/*.svg
│   ├── core/
│   │   ├── geometry.ts         # align, distribute, matchSize, matrix, swap, alignToSlide
│   │   ├── format.ts           # diff de ShapeFormat, qué aplicar según opciones
│   │   └── text.ts             # transformaciones de texto (case, trim, limpiar)
│   ├── office/
│   │   ├── selection.ts        # getSelectedShapes → SelectionContext
│   │   ├── apply.ts            # aplica ShapeModel/ShapeFormat a PowerPoint.Shape
│   │   └── runtime.ts          # wrapper de PowerPoint.run con manejo de errores
│   └── commands/
│       ├── types.ts            # Command { id, section, label, icon, minShapes, run }
│       ├── registry.ts         # lista ordenada por secciones
│       ├── alignment.ts, size.ts, format.ts, color.ts, text.ts, ...
├── tests/
│   ├── core/*.test.ts
│   └── office/*.test.ts        # con mock mínimo de PowerPoint.run
├── scripts/
│   ├── sideload-mac.sh         # copia manifest a ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
│   └── sideload-windows.md     # instrucciones carpeta compartida de confianza
├── docs/superpowers/specs/     # este documento
├── .github/workflows/deploy.yml  # build + deploy a GitHub Pages
└── README.md                   # instalación Mac/Windows, desarrollo, lista de comandos
```

## 5. Fases y comandos

### Fase 1: Alignment + Size

Alignment (requiere 2+ formas salvo "a la diapositiva", que acepta 1+):
- Alinear izquierda, centro H, derecha, arriba, centro V, abajo. Referencia: la primera forma
  seleccionada con opción "a la diapositiva".
- Distribuir horizontal y vertical (espaciado uniforme entre bordes, extremos fijos).
- Alinear en matriz: N columnas, separación configurable.
- Intercambiar posición de dos formas.
- Alinear a la diapositiva: centrar H, centrar V, ambos.

Size (requiere 2+ formas):
- Igualar ancho, alto, ambos, a la primera forma seleccionada.
- Igualar al más ancho / al más alto de la selección.
- Estirar hasta el borde izquierdo/derecho/superior/inferior de la referencia.
- (Ajustar alto al contenido se implementa en la Fase 3 junto con el resto de autofit.)

### Fase 2: Format + Color

- Copiar formato de la forma seleccionada a un portapapeles interno (fill, line, font, size).
- Pegar formato: todo, solo fill, solo line, solo font, solo size.
- Paleta de colores: colores del tema (leídos del slide master cuando el API lo permita; si no,
  paleta configurable en JSON) + recientes. Aplicar a relleno, línea o texto.
- Transparencia de relleno y línea en pasos.
- Sin relleno / sin línea.

### Fase 3: Text

- Cambiar mayúsculas: MAYÚSCULAS, minúsculas, Tipo Oración, Tipo Título.
- Tamaño de fuente +1/-1 en todas las formas seleccionadas.
- Alineación de párrafo: izquierda, centro, derecha, justificado.
- Alineación vertical del texto: arriba, medio, abajo.
- Márgenes internos: a cero, a valores estándar.
- Autofit: none / shrink / resize shape. Wrap on/off.
- Limpiar: quitar espacios dobles y finales, quitar saltos de línea vacíos.

### Fase 4: Navigation & View + Expert Tools

- Ir a diapositiva N, anterior/siguiente (selección de slide vía API 1.6).
- Seleccionar formas por tipo (todas las imágenes, todos los textos...) o por nombre.
- Renombrar formas de la selección con prefijo numerado.
- Z-order: traer al frente, enviar al fondo, adelante, atrás.
- Agrupar / desagrupar.
- Inspector: muestra left/top/width/height editables para la selección.

### Fase 5: Check + Presentation

- Eliminar notas del orador de todas las diapositivas.
- Eliminar formas vacías (textbox sin texto) de la diapositiva o de toda la presentación.
- Listar fuentes usadas en la presentación.
- Nueva presentación desde plantilla en blanco.
- Exportar diapositivas seleccionadas a un PPTX nuevo (`Slide.exportAsBase64` + insertar).

### Fase 6: Wizards

- Agenda wizard: el usuario escribe N secciones; el add-in genera una diapositiva de agenda por
  sección con la sección actual destacada, usando formas y texto nativos.

### Fuera de alcance (límites del API Office.js)

Imprimir, enviar por correo, corrector ortográfico e idioma, animaciones, edición del patrón
de diapositivas, atajos de teclado globales, galería de elementos basada en PPTX externos.

## 6. Manejo de errores

- Cada comando declara `minShapes` y, si aplica, `maxShapes`. El panel deshabilita el botón
  cuando la selección actual no cumple y muestra el motivo en tooltip.
- El panel escucha `Office.context.document.addHandlerAsync(DocumentSelectionChanged)` para
  refrescar el estado de los botones.
- `office/runtime.ts` envuelve `PowerPoint.run`; captura `OfficeExtension.Error`, registra
  `{commandId, code, debugInfo}` en consola y muestra un aviso de 3 s en el panel.
- Si `isSetSupported('PowerPointApi','1.9')` o `1.10` es falso, los comandos que lo requieran
  se ocultan con un aviso en el registro.

## 7. Pruebas

- TDD en `core/`: cada función con casos de 2 formas, N formas, formas rotadas (se usa el
  bounding box sin rotar, igual que PowerPoint), grupos (se tratan como una forma).
- `office/`: mock mínimo de `PowerPoint.run` que registra escrituras; se comprueba que cada
  comando produce un único `sync` y las escrituras esperadas.
- Manual: `npm run dev` (Vite con certificados de `office-addin-dev-certs`) + `manifest.dev.xml`
  cargado en PowerPoint Mac; checklist por fase en el README.

## 8. Distribución

- Build estático en `dist/` publicado por GitHub Actions a GitHub Pages en cada push a `main`.
  URL base: `https://xictorlr.github.io/skillsEEpowerPoint/`. El repo debe ser público.
- Mac: `scripts/sideload-mac.sh` copia `manifest.xml` a la carpeta `wef` de PowerPoint; el
  add-in aparece en Insertar > Mis complementos.
- Windows: instrucciones para añadir una carpeta compartida de confianza con el manifest
  (Archivo > Opciones > Centro de confianza > Catálogos de complementos de confianza).
- Versionado: `Version` en manifest y `package.json` sincronizados por script.

## 9. Criterios de éxito

- Fase 1 instalada en PowerPoint Mac: seleccionar 3 formas, pulsar "Alinear izquierda" y
  "Distribuir H", y obtener el resultado esperado; un solo Cmd+Z lo deshace.
- Cobertura de tests en `core/` ≥ 90 %.
- Despliegue automático funcionando: un push a `main` actualiza el add-in sin reinstalar.
