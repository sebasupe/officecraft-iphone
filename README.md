# OfficeCraft para iPhone (WordCraft · DeckCraft · GridCraft)

Instala [**WordCraft**](https://github.com/storytold/wordcraft), [**DeckCraft**](https://github.com/storytold/deckcraft) y [**GridCraft**](https://github.com/storytold/gridcraft) (de [ArtCraft](https://getartcraft.com)) en la pantalla de inicio de tu **iPhone** como apps web progresivas (PWA) gratuitas en **GitHub Pages**, con interfaces adaptadas a celular fieles a **Microsoft Word, PowerPoint y Excel para iOS**.

- **Gratis, sin App Store, sin Apple ID y sin suscripciones:** nunca caducan.
- **100 % sin conexión:** después de la primera instalación, todo corre localmente en el iPhone (WebAssembly + IndexedDB). Tus archivos nunca salen del teléfono.
- **Actualización automática:** GitHub Actions revisa cada día si salió una nueva versión oficial de WordCraft, DeckCraft o GridCraft. Si hay novedades, publica la actualización y tu iPhone descarga únicamente los archivos que cambiaron (sin volver a bajar el motor WASM si solo cambió la interfaz).
- **120 fps sin tirones (ProMotion iPhone 13 Pro):**
  - **Pellizco para zoom acelerado en GPU** (`translate3d` + `scale` a 120 Hz durante el gesto) con renderizado nítido al soltar los dedos, evitando re-renderizar WASM en cada cuadro.
  - **Desplazamiento con inercia nativa de iOS** agrupado a un solo evento por `requestAnimationFrame` y pausa automática del sondeo mientras desplazas o haces zoom.

## Cómo instalar en tu iPhone

Publica este repositorio en GitHub Pages (Configuración → Pages → Source: **GitHub Actions**). Una vez desplegado en `https://<tu-usuario>.github.io/<tu-repo>/`, puedes instalar las apps **por separado** (como Word, PowerPoint y Excel) o todas juntas desde el portal:

| App | Ruta para abrir en Safari | Qué instala en la pantalla de inicio |
|---|---|---|
| **WordCraft** (Word) | `https://<tu-usuario>.github.io/<tu-repo>/word/` | Ícono azul de WordCraft · PWA independiente sin conexión |
| **DeckCraft** (PowerPoint) | `https://<tu-usuario>.github.io/<tu-repo>/deck/` | Ícono naranja de DeckCraft · PWA independiente sin conexión |
| **GridCraft** (Excel) | `https://<tu-usuario>.github.io/<tu-repo>/grid/` | Ícono verde de GridCraft · PWA independiente sin conexión |
| **OfficeCraft** (Suite) | `https://<tu-usuario>.github.io/<tu-repo>/` | Portal unificado con recientes de las 3 apps |

En **Safari** en tu iPhone:
1. Abre la URL de la app que deseas instalar.
2. Toca **Compartir** → **Agregar a pantalla de inicio** → **Agregar**.
3. Abre el ícono desde la pantalla de inicio y espera a que termine la primera instalación offline.

## Qué incluye cada app

### WordCraft para iPhone (`/word/`) — Estilo Microsoft Word iOS
- **Inicio con plantillas:** Documento en blanco, Currículum, Informe, Carta formal, Apuntes, Recientes, Destacados y carpetas locales.
- **Botón Vista móvil / Diseño de impresión:** adapta el texto al ancho de la pantalla del iPhone (`view.webLayout`) o muestra páginas completas (`view.printLayout`).
- **Barra rápida inferior + Cinta móvil desplegable (`▲`):** pestañas **Inicio · Insertar · Disposición · Revisar · Vista** con Negrita, Cursiva, Subrayado, Resaltado, Color de fuente, Estilos (Título 1/2/3), Alineación, Viñetas, Numeración, Tablas, Imágenes desde la galería/cámara, Comentarios, Control de cambios, Contar palabras, Página oscura y Exportar a PDF / `.docx`.

### DeckCraft para iPhone (`/deck/`) — Estilo Microsoft PowerPoint iOS
- **Inicio con plantillas panorámicas 16:9** y biblioteca local de `.pptx` / `.deckcraft`.
- **Tira horizontal de diapositivas (filmstrip)** con miniaturas reales generadas por el motor y botón `+` para añadir diapositivas con selector de diseño.
- **Modo Presentación a pantalla completa (`▶`)** con gestos de deslizar para avanzar/retroceder diapositivas y notas del orador editables.
- **Cinta móvil desplegable (`▲`):** **Inicio · Insertar · Dibujar · Diseño · Transiciones · Presentación** (pluma, resaltador, formas, tablas, gráficos, fondos, Morph/transiciones y exportar a `.pptx` / PDF / PNG).

### GridCraft para iPhone (`/grid/`) — Estilo Microsoft Excel iOS
- **Inicio con plantillas con fórmulas reales (`=SUM`, `=AVERAGE`, `=COUNTIF`):** Libro en blanco, Presupuesto mensual, Registro de gastos, Lista de tareas, Horario semanal.
- **Barra de fórmulas nativa (`fx`):** indicador de celda/rango activo (tocar para Ir a celda), catálogo rápido de funciones `fx` y edición nativa con teclado iOS (`✓` confirmar / `✕` cancelar).
- **Cuadrícula Excel limpia a pantalla completa** (con encabezados `A, B, C` y `1, 2, 3`), estadísticas automáticas al seleccionar rangos (`Suma`, `Prom.`, `Recuento`) y barra inferior de pestañas de hojas (`Hoja1`, `+`).
- **Cinta móvil desplegable (`▲`):** **Inicio · Insertar · Fórmulas · Datos · Vista** (fuente, relleno, bordes, combinar celdas, formatos `$`, `%`, decimales, gráficos, Autosuma, ordenar A→Z/Z→A, filtros, inmovilizar fila superior/primera columna y exportar a `.xlsx` / PDF).

## Probar localmente en tu PC

```bash
python scripts/descargar_web.py web
python scripts/preparar.py web assets _site
python scripts/servidor.py _site 8765
```

Luego abre `http://127.0.0.1:8765/` en tu navegador.

## Licencias y créditos

Los binarios web de **WordCraft**, **DeckCraft** y **GridCraft** se publican directamente desde los releases oficiales de [Storytold / ArtCraft](https://github.com/storytold) con verificación SHA-256 y sin modificar. Se distribuyen bajo licencia MIT o Apache-2.0 (`assets/LICENSE.txt`). Este proyecto no está afiliado a ArtCraft ni a Microsoft.
