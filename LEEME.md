# Avakian · Grupo Inmobiliario — sitio web

Sitio estático (no necesita base de datos ni programación para actualizarse).
Todo el contenido sale del archivo **`datos/propiedades.xlsx`**.

## Carpetas

| Carpeta | Qué contiene |
|---|---|
| `index.html` | La página |
| `datos/propiedades.xlsx` | **La base de datos**: propiedades, desarrollos y datos de contacto |
| `imagenes/` | Fotos de las propiedades (las que se nombran en el Excel) |
| `assets/` | Estilos, lógica, tipografías, logo y favicon |
| `vendor/` | Librerías del mapa (MapLibre) y del lector de Excel (SheetJS) |

## Actualizar propiedades

1. Abre `datos/propiedades.xlsx` (la hoja **Instrucciones** explica cada paso).
2. Agrega una fila en **Propiedades** o **Desarrollos**.
3. Ubicación: clic derecho en Google Maps sobre el punto → copiar coordenadas → pegar (ej. `32.5149, -117.0382`).
4. Sube las fotos a `imagenes/` y escribe sus nombres separados por coma.
5. Guarda y sube el Excel al hosting reemplazando el anterior.

Vista previa sin publicar: abre el sitio y arrastra el Excel sobre la página.

## Publicar

Sube la carpeta completa a cualquier hosting estático (Netlify, Vercel, GitHub Pages, cPanel…).
Para verlo en tu computadora hace falta un servidor local; con doble clic no carga el Excel:

```bash
python3 -m http.server 8000
```

## Marca

- **Tipografía:** Archivo (marca, texto e interfaz; en ancho extendido para el nombre) + Instrument Serif (títulos editoriales). Ambas de licencia libre (OFL).
- **Color:** carbón `#222027` (el del logo), hueso `#F5F1EA`, arena `#E7DFD2`, cobre `#C2572B`, pacífico `#16343D`.
- **Logo:** el sitio lo compone con texto (Archivo ancho, peso 600) para que se vea nítido a cualquier tamaño. `assets/logo.svg` y `assets/logo-claro.svg` son la misma recreación; `assets/logo-original.jpg` es el archivo recibido.

## Portada

1. **Intro:** el logo AVAKIAN aparece sobre carbón, sus letras se vuelven una ventana hacia la ciudad y la cámara atraviesa la «V».
2. **Escena 1 · La ciudad:** vuelo sobre imágenes satelitales reales de Tijuana.
3. **Escenas 2 y 3 · Torres e inversión:** fotos reales de Tijuana a pantalla completa con zoom lento
   (`imagenes/tijuana-skyline.jpg` y `imagenes/tijuana-agua-caliente.jpg`, de Urbaner44 en Wikimedia Commons,
   licencia CC BY-SA 4.0: exige mantener el crédito visible, por eso aparece en la esquina de la portada).
   Para usar fotos o renders propios, escríbelos en **Configuración → Imágenes de portada** del Excel; el crédito desaparece solo.

Los titulares de cada escena están en `assets/app.js` (lista `SCENES`).
Si tienes una toma de dron, súbela a `assets/` y escribe su nombre en **Configuración → Video de portada**.

## Créditos y licencias

- Mapa: MapLibre GL, OpenFreeMap / OpenStreetMap.
- Imagen satelital: Esri World Imagery. Para uso comercial en producción revisa sus términos
  (puede requerir una cuenta gratuita de ArcGIS o cambiar a otro proveedor como MapTiler o Mapbox).
- Las fotos de `imagenes/` son de muestra (Unsplash) y las propiedades con clave `EJ-`/`DES-` son ficticias.
