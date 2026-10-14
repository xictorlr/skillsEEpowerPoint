# Slide Tools

Panel lateral para PowerPoint (Mac y Windows) con herramientas de alineación, tamaño, formato y texto,
inspirado en el flujo de trabajo de los paneles de productividad para presentaciones.

Autor: Victor Lopez · Add-in web (Office.js) · Hospedado en GitHub Pages.

## Instalación

### Mac

1. Descarga `manifest.xml` desde https://xictorlr.github.io/skillsEEpowerPoint/manifest.xml
2. Cópialo a `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` (crea la carpeta si no existe).
   Con el repo clonado: `bash scripts/sideload-mac.sh manifest.xml`
3. Reinicia PowerPoint. Insertar > Complementos > Mis complementos > Complementos para desarrolladores > Slide Tools.
4. En la pestaña Inicio aparece el botón "Abrir panel".

### Windows

1. Descarga `manifest.xml` a una carpeta y compártela (clic derecho > Propiedades > Compartir > compartir contigo mismo).
   Copia la ruta de red (`\\TU-PC\carpeta`).
2. PowerPoint > Archivo > Opciones > Centro de confianza > Configuración > Catálogos de complementos de confianza.
   Añade la ruta de red, marca "Mostrar en el menú" y acepta.
3. Reinicia PowerPoint. Insertar > Mis complementos > Carpeta compartida > Slide Tools.

## Comandos (Fase 1)

**Alignment:** alinear izquierda / centro / derecha / arriba / medio / abajo (a la primera forma seleccionada),
distribuir horizontal y vertical, intercambiar posiciones, alinear en matriz (columnas y separación configurables),
centrar en la diapositiva (horizontal, vertical, ambos).

**Size:** igualar ancho / alto / ambos al primero, igualar al más ancho / alto / estrecho / bajo,
estirar hasta el borde izquierdo / derecho / superior / inferior de la referencia.

La **referencia** es siempre la primera forma que seleccionaste. Cada comando es un único paso de deshacer.

El panel también incluye una sección **Ajustes** con las columnas y el espaciado (gap) de la matriz y un
selector de **Diapositiva**. En modo Auto se detecta la proporción de la diapositiva activa y se asume el alto estándar de 540 pt;
si tu presentación usa un alto distinto, como el 16:9 clásico de 720×405, elígelo explícitamente para que
"centrar en la diapositiva" quede exacto.

## Comandos (Fase 2)

**Format:** copiar formato (de la primera forma seleccionada) y pegar formato en el resto, ya sea todo junto
(según las casillas elegidas) o solo relleno / línea / fuente / tamaño. Las casillas "Relleno", "Línea",
"Fuente" y "Tamaño" controlan qué partes copia "Pegar formato" y se recuerdan entre sesiones.

**Color:** destino del color (relleno, línea o texto), paleta neutra de 16 colores, hasta 8 colores recientes
y un selector de color personalizado; además "Sin relleno", "Sin línea" y transparencia ±10 %. El destino
elegido y los colores recientes también se recuerdan entre sesiones.

Nota: la paleta es fija porque Office.js no expone los colores del tema de la presentación.

## Limitaciones conocidas

- Office.js no expone el tamaño de la diapositiva: en Auto se deduce de la proporción asumiendo 540 pt de alto,
  y el valor se guarda mientras el panel esté abierto. Si cambias Diseño > Tamaño de diapositiva, elige el
  tamaño en Ajustes o vuelve a abrir el panel.
- Las formas rotadas se alinean por su marco sin rotar (igual que expone el API), no por su caja visual.
- Los comandos actúan sobre las formas devueltas por la selección; dentro de un grupo se trata el grupo como una forma.
- El formato copiado ("Copiar formato") vive en memoria mientras el panel está abierto; se pierde al cerrarlo.
- La transparencia (+10 % / −10 %) no aplica cuando el destino de color es "Texto".

## Desarrollo

```bash
npm install
npm run certs          # una vez: certificados HTTPS de desarrollo
npm run manifest:dev   # genera manifest.dev.xml (localhost:3000)
npm run sideload:mac   # copia el manifest dev a PowerPoint
npm run dev            # servidor Vite
npm test               # Vitest
npm run typecheck      # TypeScript sin emitir
npm run build          # dist/
```

Despliegue: GitHub Actions publica `dist/` en GitHub Pages en cada push a `main`. La primera vez hay que
poner Settings > Pages > Source = "GitHub Actions" en el repositorio.

`npm run sideload:mac` ya genera el manifest de desarrollo (ejecuta `manifest:dev` internamente) antes de
copiarlo a PowerPoint, así que no hace falta llamarlo aparte en ese flujo.

Estructura: `src/core` (lógica pura con tests), `src/office` (adaptador Office.js), `src/commands` (comandos
declarativos), `src/taskpane` (UI). Diseño completo en `docs/superpowers/specs/`.

## Hoja de ruta

Fase 2 Format + Color (hecho) · Fase 3 Text · Fase 4 Navigation & Expert Tools · Fase 5 Check + Presentation ·
Fase 6 Wizards.
