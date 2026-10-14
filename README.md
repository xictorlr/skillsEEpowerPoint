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
selector de **Diapositiva**. En modo Auto, con PowerPoint reciente (PowerPointApi 1.10) el tamaño se lee exacto; en versiones anteriores se
detecta la proporción de la diapositiva activa y se asume el alto estándar de 540 pt;
si tu presentación usa un alto distinto, como el 16:9 clásico de 720×405, elígelo explícitamente para que
"centrar en la diapositiva" quede exacto.

## Comandos (Fase 2)

**Format:** copiar formato (de la primera forma seleccionada) y pegarlo en las formas seleccionadas (incluso en otra diapositiva), ya sea todo junto
(según las casillas elegidas) o solo relleno / línea / fuente / tamaño. Las casillas "Relleno", "Línea",
"Fuente" y "Tamaño" controlan qué partes pega "Pegar formato" y se recuerdan entre sesiones.

**Color:** destino del color (relleno, línea o texto), paleta neutra de 16 colores, hasta 8 colores recientes
y un selector de color personalizado; además "Sin relleno", "Sin línea" y transparencia ±10 %. El destino
elegido y los colores recientes también se recuerdan entre sesiones.

Nota: la paleta es fija porque Office.js no expone los colores del tema de la presentación.

## Comandos (Fase 3)

**Text:** MAYÚSCULAS / minúsculas / Tipo oración / Tipo Título; tamaño de fuente ± (paso configurable en
Ajustes); alineación de párrafo izquierda / centro / derecha / justificado; texto arriba / medio / abajo;
márgenes internos a cero / estándar; autoajuste: sin autoajuste / reducir texto al desbordar / ajustar
forma al texto; ajustar texto a la forma sí/no; limpiar texto.

## Comandos (Fase 4)

**Navigation:** primera / anterior / siguiente / última diapositiva, e "Ir a diapositiva" (campo
numérico + botón "Ir"). El indicador "N / M" y el límite del campo "Ir a" se actualizan solos con
la diapositiva activa.

**Expert:** seleccionar por tipo (Todas, Texto, Formas, Imágenes, Líneas, Tablas, Grupos, Otras,
sobre la diapositiva activa), seleccionar por nombre (subcadena, sin distinguir mayúsculas ni
acentos), renombrar numerado (prefijo configurable y persistido, orden de lectura; si lo dejas
vacío, se usa "Forma"), traer al frente / adelante
/ enviar atrás / al fondo, agrupar / desagrupar, e inspector de posición y tamaño (X, Y, Ancho,
Alto en puntos) que lee y escribe la caja envolvente de la selección en vivo.

Con PowerPoint reciente (PowerPointApi 1.10) el tamaño de diapositiva es exacto (ver Ajustes) y
en **Color** aparece una fila "Tema" con los 12 colores del tema de la diapositiva activa, que se
pueden usar igual que la paleta o los recientes.

Limitaciones de Fase 4:

- El inspector mueve/escala la selección de forma proporcional respecto a su caja envolvente
  actual: cambiar solo un eje (p. ej. Ancho) no deja el resto del eje intacto si la caja de
  partida no estaba ya alineada a ese borde.
- "Seleccionar por tipo" y "Seleccionar por nombre" actúan sobre la diapositiva activa, no sobre
  toda la presentación.
- La fila "Tema" solo aparece en PowerPoint con PowerPointApi 1.10; en el resto, Color se queda
  sin esa fila (la paleta fija y los recientes siguen disponibles).

## Comandos (Fases 5 y 6)

**Check:** eliminar los cuadros de texto vacíos de la diapositiva activa o de toda la presentación
(este segundo, por ser destructivo sobre varias diapositivas a la vez, pide confirmación: pulsa el
botón una vez para armarlo y una segunda vez en los 5 segundos siguientes para ejecutarlo; cualquier
otro clic o dejar pasar ese tiempo lo cancela sin cambiar nada); seleccionar las formas que quedan
fuera de la diapositiva activa; listar las fuentes usadas en todo el mazo.

**Presentation:** abrir una presentación nueva en vacío; exportar las diapositivas seleccionadas a
una presentación nueva (requiere PowerPoint reciente, PowerPointApi 1.10; en el resto avisa sin
exportar).

**Wizards:** asistente "Crear agenda" con un título (por defecto "Agenda") y un cuadro de texto con
una sección por línea; "Crear agenda" inserta una diapositiva por sección justo después de la
diapositiva activa, cada una con el título y la lista completa de secciones, resaltando la propia
con el color de énfasis del tema (o un azul por defecto sin PowerPointApi 1.10). El título y las
secciones se recuerdan entre sesiones.

## Limitaciones conocidas

- Al pegar formato no se copian rellenos degradados, de trama ni de imagen; los valores mixtos (por ejemplo, varias
  fuentes dentro de un mismo texto) se omiten campo a campo; la transparencia solo se ajusta en rellenos sólidos.

- Sin PowerPointApi 1.10, Office.js no expone el tamaño de la diapositiva: en Auto se deduce de la proporción asumiendo 540 pt de alto,
  y el valor se guarda mientras el panel esté abierto. Si cambias Diseño > Tamaño de diapositiva, elige el
  tamaño en Ajustes o vuelve a abrir el panel.
- Las formas rotadas se alinean por su marco sin rotar (igual que expone el API), no por su caja visual.
- Los comandos actúan sobre las formas devueltas por la selección; dentro de un grupo se trata el grupo como una forma.
- El formato copiado ("Copiar formato") vive en memoria mientras el panel está abierto; se pierde al cerrarlo.
- La transparencia (+10 % / −10 %) no aplica cuando el destino de color es "Texto".
- Cambiar mayúsculas y limpiar texto reescriben el texto por sub-rangos (palabra a palabra cuando la
  longitud no cambia) para conservar negritas y viñetas; cuando la longitud de una línea cambia o se
  eliminan párrafos vacíos, esa línea o el cuadro completo pierden formatos mixtos.
- "Tamaño ±" omite cuadros con tamaños de fuente mixtos.
- Las siglas pierden las mayúsculas con Tipo oración / Tipo Título.
- Office.js no expone las notas del orador ni las animaciones: Check no puede detectarlas ni
  limpiarlas, solo actúa sobre las formas de la diapositiva.
- Las comprobaciones de Check (cuadros de texto vacíos, formas fuera de la diapositiva, fuentes)
  no entran dentro de grupos: un cuadro vacío o una forma descolocada dentro de un grupo no se
  detecta.
- "Exportar diapositivas seleccionadas" requiere PowerPoint reciente (PowerPointApi 1.10); en el
  resto, el botón avisa sin exportar.
- El asistente "Crear agenda" usa cuadros de texto nativos con una maquetación fija, no el diseño
  del patrón de diapositivas (layout) de la presentación.

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

Fase 2 Format + Color (hecho) · Fase 3 Text (hecho) · Fase 4 Navigation & Expert Tools (hecho) ·
Fase 5 Check + Presentation (hecho) · Fase 6 Wizards (hecho). Todas las fases planeadas están completas.

## Prueba manual

Tras instalar el add-in, en una presentación de prueba:

1. **Alignment/Size:** selecciona 3 rectángulos de distinto tamaño; "Alinear a la izquierda" los alinea al primero
   seleccionado y "Distribuir horizontalmente" los reparte con separación uniforme. Cmd+Z deshace cada acción de una vez.
2. **Diapositiva:** con una forma, "Centrar en la diapositiva" la centra tanto en 16:9 como en 4:3.
3. **Format:** copia el formato de un cuadro de texto con relleno y pégalo sobre otra forma; desmarca "Tamaño"
   y comprueba que el tamaño no cambia. Con una imagen seleccionada junto a un texto, pegar fuente no da error.
4. **Color:** cambia el destino (relleno, línea, texto), pulsa una muestra y un color personalizado; "Sin relleno"
   y "Sin línea" funcionan; la transparencia se deshabilita con destino texto.
5. **Formas especiales:** selecciona a la vez un grupo, una tabla, una línea y un marcador de imagen, y ejecuta
   "Copiar formato" y una muestra de color: el panel debe mostrar un aviso o aplicar el cambio, nunca quedarse colgado.
6. **Text:** en un cuadro con viñetas y una primera línea en negrita, aplica "MAYÚSCULAS" y comprueba que las
   viñetas y la negrita se conservan; pon los márgenes internos a cero y verifica que el texto queda pegado
   al borde; aplica "ajustar forma al texto" y comprueba que la forma crece o se encoge según el contenido.

7. **Text, casos límite:** aplica MAYÚSCULAS a un cuadro con un emoji y a otro con un campo de número de
   diapositiva (deben cambiar solo las letras, sin desplazar caracteres); "Limpiar texto" en un cuadro con
   párrafos vacíos en medio y al final (deben desaparecer conservando negritas); "Tipo oración" en un texto con
   un salto de línea suave (Mayús+Intro) dentro de una frase (no debe capitalizar tras el salto); "Texto en el
   medio" sobre una forma ya centrada horizontalmente (debe conservar el centrado). Si algo falla, Cmd+Z lo deshace.

8. **Navigation/Expert:** en una presentación de 4+ diapositivas, usa primera/anterior/siguiente/última y
   comprueba que el indicador "N / M" sigue a la diapositiva activa; escribe un número en "Ir a" y pulsa "Ir"
   (y prueba uno fuera de rango: debe avisar sin mover la selección). En una diapositiva con formas de varios
   tipos, "Seleccionar por tipo" (p. ej. "Imágenes") selecciona solo esas formas de la diapositiva activa, y
   "Seleccionar por nombre" selecciona las que contienen el texto buscado. Cambia el "Prefijo" y pulsa
   "Renombrar numerado" sobre 3 formas: deben quedar nombradas en orden de lectura. Con 3 formas superpuestas,
   "Traer al frente" / "Enviar al fondo" cambian su apilado; selecciona 2 y "Agrupar", luego "Desagrupar".
   Por último, con una forma seleccionada, cambia X/Y/Ancho/Alto en el inspector y confirma que la forma se
   mueve/escala en consecuencia, y que los campos se actualizan solos al cambiar de selección.

9. **Check/Presentation/Wizards:** en una diapositiva con un cuadro de texto vacío y otro con texto,
   "Eliminar cuadros de texto vacíos (diapositiva)" borra solo el vacío; añade cuadros vacíos en
   varias diapositivas y pulsa "Eliminar cuadros de texto vacíos (mazo)": el botón pide confirmación
   (cambia a "Pulsa de nuevo para confirmar") y solo borra al pulsarlo una segunda vez dentro de los
   5 segundos; pulsa una vez y espera (o haz clic en otro botón) para comprobar que se cancela sin
   borrar nada. Con una forma arrastrada fuera del lienzo, "Seleccionar formas fuera de la
   diapositiva" la selecciona; "Fuentes del mazo" lista las fuentes usadas. "Nueva presentación"
   abre una presentación en blanco en otra ventana; selecciona 2 diapositivas y "Exportar
   diapositivas seleccionadas" crea una presentación nueva solo con esas 2 (en PowerPoint sin API
   1.10 debe avisar en vez de fallar). En Wizards, escribe un título y 5 secciones (una por línea)
   y pulsa "Crear agenda": deben insertarse 5 diapositivas justo después de la activa, cada una con
   el título, la lista completa de secciones y su propia sección resaltada con el color de énfasis
   del tema; cierra y vuelve a abrir el panel y comprueba que el título y las secciones siguen ahí.