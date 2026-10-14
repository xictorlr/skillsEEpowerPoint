# Guía de instalación de Slide Tools

Esta guía está pensada para cualquier persona, aunque no haya instalado nunca un complemento de
PowerPoint. Tardarás unos 5 minutos.

## ¿Qué es esto?

Slide Tools es un panel lateral para PowerPoint con herramientas para alinear, dimensionar,
formatear y organizar formas y texto en tus diapositivas. Se instala como "complemento" (add-in):
no hace falta descargar ni instalar un programa, solo un pequeño archivo llamado `manifest.xml`
que le dice a PowerPoint dónde encontrar el panel.

Necesitas:

- PowerPoint de Microsoft 365 (la versión de suscripción; no funciona con PowerPoint 2016/2019 "de
  toda la vida" sin suscripción).
- Conexión a internet: el panel se carga desde internet (GitHub Pages) cada vez que lo abres.

Elige la sección de tu sistema operativo.

## Mac (5 minutos)

### Opción con Terminal (más rápida)

1. Abre la app **Terminal** (Lanzador > Terminal, o Spotlight con Cmd+Espacio y escribe "Terminal").
2. Copia y pega este comando completo, y pulsa Intro:

   ```
   mkdir -p "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef" && \
   curl -L -o "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/slide-tools-manifest.xml" \
   https://xictorlr.github.io/skillsEEpowerPoint/manifest.xml
   ```

   Qué verás: ninguna ventana especial, solo el cursor del Terminal esperando. Si no aparece ningún
   error en rojo, ha funcionado.
3. El comando ya ha descargado el archivo y lo ha dejado en su sitio. Continúa en el paso 4 de
   "Opción manual" más abajo (cerrar y volver a abrir PowerPoint).

   Esta ruta es la misma aunque hayas instalado PowerPoint desde la Mac App Store.

### Opción manual (con el Finder)

1. Descarga el archivo del complemento abriendo este enlace en tu navegador:
   https://xictorlr.github.io/skillsEEpowerPoint/manifest.xml
   Qué verás: el navegador descarga un archivo llamado `manifest.xml` (o se abre como texto; en
   ese caso, usa "Archivo > Guardar como" para guardarlo).
2. Abre el Finder y pulsa Cmd+Mayús+G ("Ir a la carpeta"). Pega esta ruta:

   ```
   ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
   ```

   Qué verás: si la carpeta no existe, el Finder te avisa. Créala manualmente (puedes crear las
   carpetas que falten paso a paso, por ejemplo abriendo primero `~/Library/Containers` y creando
   `com.microsoft.Powerpoint`, `Data`, `Documents` y `wef` a mano si hiciera falta) o usa la
   opción con Terminal de arriba, que crea la carpeta sola.
3. Copia el `manifest.xml` descargado dentro de esa carpeta `wef`.
4. Cierra PowerPoint por completo (Cmd+Q) y vuelve a abrirlo.
5. Abre o crea una presentación. Ve a **Insertar > Complementos > Mis complementos >
   Complementos para desarrolladores > Slide Tools**.
   Qué verás: "Slide Tools" aparece listado como complemento para desarrolladores; selecciónalo
   e insértalo.
6. En la pestaña **Inicio** de la cinta de opciones aparece un nuevo botón **"Abrir panel"**.
   Púlsalo: el panel de Slide Tools se abre a la derecha de la pantalla.

Si no ves el complemento, repasa "Problemas frecuentes" más abajo.

## Windows (5 minutos)

1. Descarga el archivo del complemento abriendo este enlace en tu navegador:
   https://xictorlr.github.io/skillsEEpowerPoint/manifest.xml
   Guárdalo en una carpeta fácil de encontrar, por ejemplo `C:\SlideTools`. Evita espacios en el
   nombre de la carpeta (usa `SlideTools`, no `Slide Tools`).
   Qué verás: un archivo `manifest.xml` dentro de esa carpeta.
2. Comparte esa carpeta contigo mismo para poder apuntar PowerPoint a ella:
   - Clic derecho sobre la carpeta `C:\SlideTools` > **Propiedades** > pestaña **Compartir** >
     **Compartir...** > elige tu propio usuario y pulsa **Compartir**.
   - Qué verás: un mensaje confirmando que la carpeta está compartida, con una ruta de red del
     tipo `\\NOMBRE-PC\SlideTools`. Apunta esa ruta, la necesitas en el paso siguiente.
   - Si Windows muestra un aviso de seguridad al compartir la carpeta, permite el acceso para
     continuar.
3. Abre PowerPoint y ve a **Archivo > Opciones > Centro de confianza > Configuración del Centro
   de confianza... > Catálogos de complementos de confianza**.
4. Pega la ruta de red (`\\NOMBRE-PC\SlideTools`) en el campo de catálogo y pulsa **"Agregar
   catálogo"**. Marca la casilla **"Mostrar en el menú"** y pulsa **Aceptar** en ambas ventanas.
5. Cierra PowerPoint por completo y vuelve a abrirlo.
6. Abre o crea una presentación. Ve a **Insertar > Mis complementos**, abre la pestaña **CARPETA
   COMPARTIDA** y selecciona **Slide Tools**.
   Qué verás: "Slide Tools" listado ahí; insértalo.
7. En la pestaña **Inicio** aparece el botón **"Abrir panel"**. Púlsalo para abrir el panel a la
   derecha.

Nota: Slide Tools requiere PowerPoint de Microsoft 365 (versiones de 2024 en adelante funcionan
mejor); las versiones más antiguas pueden ocultar algún botón, ver más abajo.

## ¿Funciona?

Para comprobar que todo está bien instalado:

1. Abre cualquier presentación y dibuja un par de formas (rectángulos, por ejemplo).
2. Pulsa el botón **"Abrir panel"** en la pestaña Inicio.
3. Deberías ver el panel de Slide Tools a la derecha, con secciones como "Alignment" y "Size".
4. Selecciona las dos formas y pulsa cualquier botón de alineación: las formas deberían moverse.

Si todo esto ocurre, el complemento está instalado y funcionando correctamente.

## Problemas frecuentes

**No aparece el complemento en el menú.**
Reinicia PowerPoint por completo (ciérralo del todo, no solo la ventana) y vuelve a abrirlo.
Repasa que el archivo esté en la carpeta correcta (Mac) o que la ruta de red esté bien escrita y
el catálogo activo (Windows). Comprueba también que el archivo se llame tal cual termina en
`.xml` y no, por ejemplo, `manifest.xml.txt` (algunos navegadores añaden esa extensión al
descargar).

**El panel aparece en blanco o no carga.**
El panel se carga desde internet cada vez que lo abres, así que revisa tu conexión. Si estás en
una red de oficina con proxy o restricciones de seguridad, puede que bloquee la dirección
`xictorlr.github.io`; prueba desde otra red o pide a tu equipo de IT que la permita.

**Faltan algunos botones del panel (por ejemplo, "Exportar diapositivas seleccionadas").**
Algunos botones requieren una versión reciente de PowerPoint (la que soporta "PowerPointApi
1.10"). Si tu PowerPoint es más antiguo, esos botones concretos no aparecen, pero el resto del
panel funciona con normalidad. Actualizar PowerPoint a la última versión suele solucionarlo.

**En Mac sigue apareciendo una versión antigua del panel después de actualizar.**
PowerPoint guarda una copia en caché. Cierra PowerPoint y borra el contenido de esta carpeta:

```
~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/Wef
```

Vuelve a abrir PowerPoint y el panel se recargará desde cero.

## Desinstalar

**Mac:** borra el archivo `manifest.xml` (o `slide-tools-manifest.xml`) de la carpeta
`~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef` y reinicia PowerPoint.

**Windows:** ve a **Archivo > Opciones > Centro de confianza > Configuración del Centro de
confianza... > Catálogos de complementos de confianza**, selecciona el catálogo que añadiste y
quítalo con el botón correspondiente. Reinicia PowerPoint.

En ambos casos, el complemento desaparece del menú "Mis complementos" la próxima vez que abras
PowerPoint.
