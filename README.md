# elpx-docx

Aplicaci\u00f3n web para convertir un `.elpx` a `.docx` reutilizando el exportador `HTML5 Single Page` de eXeLearning y una conversi\u00f3n HTML a DOCX en JavaScript puro.

## Requisitos

- `bun`
- Un checkout local del repo de eXeLearning

Por defecto se busca en `../exelearning`, pero tambi\u00e9n puede indicarse desde la interfaz web o con `EXELEARNING_DIR`.

## Arranque

```bash
bun install
bun run start
```

La aplicaci\u00f3n queda disponible en `http://localhost:3007`.

## Uso

1. Abre `http://localhost:3007`.
2. Sube un archivo `.elpx` o `.elp`.
3. Ajusta la ruta del repo de eXeLearning si hace falta.
4. Descarga el `.docx` generado.

## Estructura

- `src/converter.ts`: l\u00f3gica de conversi\u00f3n reutilizable.
- `src/index.ts`: interfaz web y endpoint `POST /convert`.
La idea es que `src/converter.ts` pueda moverse luego a eXeLearning y llamarse desde una ruta propia del servidor.

## Flujo t\u00e9cnico

1. Ejecuta `export-html5-sp` del CLI de eXeLearning.
2. Lee el ZIP en memoria e inlina CSS e im\u00e1genes.
3. Convierte el HTML resultante a `.docx` con `@turbodocx/html-to-docx`.

## Limitaciones

- Recursos remotos referenciados desde el HTML exportado no se incrustan; solo se embeben los recursos locales que vienen dentro del ZIP single-page.
- El resultado final depende de c\u00f3mo exporte eXeLearning el contenido single-page y de c\u00f3mo la librer\u00eda HTML a DOCX interprete ese HTML.
