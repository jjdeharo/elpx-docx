import path from "node:path";
import { convertElpxToDocx } from "./converter";

const PORT = Number(process.env.PORT || 3007);
const DEFAULT_EXELEARNING_DIR = path.resolve(process.env.EXELEARNING_DIR || path.resolve(process.cwd(), "..", "exelearning"));

const server = Bun.serve({
  port: PORT,
  routes: {
    "/": {
      GET: () => new Response(renderHome(), { headers: { "content-type": "text/html; charset=utf-8" } }),
    },
    "/convert": {
      POST: async req => handleConvert(req),
    },
    "/health": {
      GET: () => Response.json({ ok: true }),
    },
  },
  fetch: () => new Response("Not found", { status: 404 }),
});

console.log(`Servidor en http://localhost:${server.port}`);

async function handleConvert(request: Request): Promise<Response> {
  try {
    const formData = await request.formData();
    const file = formData.get("file");
    const exelearningDirValue = formData.get("exelearningDir");

    if (!(file instanceof File)) {
      return htmlError("Debes subir un archivo .elpx o .elp.", 400);
    }

    if (!(file.name.endsWith(".elpx") || file.name.endsWith(".elp"))) {
      return htmlError("El archivo debe tener extensi\u00f3n .elpx o .elp.", 400);
    }

    const result = await convertElpxToDocx({
      inputBuffer: new Uint8Array(await file.arrayBuffer()),
      inputFilename: file.name,
      exelearningDir:
        typeof exelearningDirValue === "string" && exelearningDirValue.trim()
          ? exelearningDirValue.trim()
          : DEFAULT_EXELEARNING_DIR,
    });

    const headers = new Headers({
      "content-type":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "content-disposition": `attachment; filename="${result.outputFilename}"`,
    });

    return new Response(new Uint8Array(result.outputBuffer), { headers });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return htmlError(message, 500);
  }
}

function htmlError(message: string, status: number): Response {
  return new Response(renderHome(message), {
    status,
    headers: { "content-type": "text/html; charset=utf-8" },
  });
}

function renderHome(errorMessage?: string): string {
  const safeDir = escapeHtml(DEFAULT_EXELEARNING_DIR);
  const errorBlock = errorMessage
    ? `<p class="error">${escapeHtml(errorMessage)}</p>`
    : "";

  return `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>ELPX a DOCX</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f3efe5;
      --panel: #fffaf0;
      --ink: #1f2a1f;
      --accent: #1d6b57;
      --accent-2: #d6efe5;
      --border: #c8bea8;
      --danger: #7b1e1e;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Georgia, "Times New Roman", serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top right, rgba(29,107,87,0.18), transparent 32%),
        linear-gradient(180deg, #f6f0e0 0%, var(--bg) 100%);
      min-height: 100vh;
      display: grid;
      place-items: center;
      padding: 24px;
    }
    main {
      width: min(760px, 100%);
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 18px;
      padding: 28px;
      box-shadow: 0 18px 50px rgba(47, 39, 23, 0.12);
    }
    h1 {
      margin: 0 0 8px;
      font-size: clamp(2rem, 4vw, 3.2rem);
      line-height: 1;
    }
    p {
      margin: 0 0 16px;
      line-height: 1.45;
    }
    form {
      display: grid;
      gap: 14px;
      margin-top: 18px;
    }
    label {
      display: grid;
      gap: 6px;
      font-weight: 700;
    }
    input[type="file"],
    input[type="text"] {
      width: 100%;
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 12px;
      background: #fff;
      font: inherit;
    }
    button {
      border: 0;
      border-radius: 999px;
      padding: 14px 18px;
      background: linear-gradient(135deg, var(--accent) 0%, #2e8b72 100%);
      color: #fff;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
    }
    .hint {
      font-size: 0.95rem;
      color: #4d5746;
    }
    .error {
      border: 1px solid rgba(123, 30, 30, 0.2);
      background: #fff1ef;
      color: var(--danger);
      border-radius: 10px;
      padding: 12px;
      margin-bottom: 12px;
    }
    code {
      background: var(--accent-2);
      padding: 0.1em 0.35em;
      border-radius: 0.35em;
    }
  </style>
</head>
<body>
  <main>
    <h1>ELPX a DOCX</h1>
    <p>Sube un proyecto de eXeLearning y la aplicaci\u00f3n generar\u00e1 un <code>.docx</code> reutilizando el exportador <code>HTML5 Single Page</code>.</p>
    ${errorBlock}
    <form action="/convert" method="post" enctype="multipart/form-data">
      <label>
        Archivo .elpx
        <input type="file" name="file" accept=".elpx,.elp" required>
      </label>
      <label>
        Ruta de eXeLearning
        <input type="text" name="exelearningDir" value="${safeDir}">
      </label>
      <button type="submit">Convertir y descargar</button>
    </form>
    <p class="hint">Pensado para evolucionar luego a una ruta interna de eXeLearning: la l\u00f3gica de conversi\u00f3n est\u00e1 aislada en <code>src/converter.ts</code>.</p>
  </main>
</body>
</html>`;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
