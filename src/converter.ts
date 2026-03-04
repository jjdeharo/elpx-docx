import HtmlToDocx from "@turbodocx/html-to-docx";
import { unzipSync } from "fflate";
import { spawn } from "node:child_process";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

const DEFAULT_EXELEARNING_DIR = path.resolve(process.cwd(), "..", "exelearning");

export interface ConvertOptions {
  inputBuffer: Uint8Array;
  inputFilename: string;
  outputFilename?: string;
  exelearningDir?: string;
}

export interface ConvertResult {
  outputBuffer: Buffer;
  outputFilename: string;
}

export async function convertElpxToDocx(options: ConvertOptions): Promise<ConvertResult> {
  const exelearningDir = path.resolve(options.exelearningDir || DEFAULT_EXELEARNING_DIR);
  await assertDirectory(exelearningDir, `No existe el repo de eXeLearning: ${exelearningDir}`);

  const tempDir = await mkdtemp(path.join(tmpdir(), "elpx-to-docx-"));
  const inputName = sanitizeInputFilename(options.inputFilename);
  const inputPath = path.join(tempDir, inputName);
  const zipPath = path.join(tempDir, "single-page.zip");
  const outputFilename = sanitizeOutputFilename(options.outputFilename || inputName);
  await writeFile(inputPath, options.inputBuffer);

  try {
    const cliArgs = await resolveExeLearningCliArgs(exelearningDir);

    await runCommand(cliArgs[0], [...cliArgs.slice(1), "export-html5-sp", inputPath, zipPath], {
      cwd: exelearningDir,
    });

    await assertFile(zipPath, `No se ha generado ${zipPath}`);
    const zipBuffer = await readFile(zipPath);
    const zipEntries = unzipSync(new Uint8Array(zipBuffer));
    const html = buildInlineHtml(zipEntries);
    const outputBuffer = await asBuffer(
      HtmlToDocx(html, undefined, {
        title: outputFilename.replace(/\.docx$/i, ""),
        creator: "elpx-docx",
        lang: "es-ES",
        imageProcessing: {
          svgHandling: "native",
          suppressSharpWarning: true,
        },
      }),
    );

    return { outputBuffer, outputFilename };
  } finally {
    await rm(tempDir, { recursive: true, force: true });
  }
}

async function resolveExeLearningCliArgs(exelearningDir: string): Promise<string[]> {
  const distCli = path.join(exelearningDir, "dist", "cli.js");
  const srcCli = path.join(exelearningDir, "src", "cli", "index.ts");

  if (await exists(distCli)) {
    return ["bun", "run", "dist/cli.js"];
  }

  if (await exists(srcCli)) {
    return ["bun", "src/cli/index.ts"];
  }

  throw new Error(`No encuentro el CLI de eXeLearning en ${exelearningDir}`);
}

async function exists(filePath: string): Promise<boolean> {
  try {
    await stat(filePath);
    return true;
  } catch {
    return false;
  }
}

async function assertDirectory(dirPath: string, errorMessage: string): Promise<void> {
  try {
    const info = await stat(dirPath);
    if (!info.isDirectory()) {
      throw new Error(errorMessage);
    }
  } catch {
    throw new Error(errorMessage);
  }
}

async function assertFile(filePath: string, errorMessage: string): Promise<void> {
  try {
    const info = await stat(filePath);
    if (!info.isFile()) {
      throw new Error(errorMessage);
    }
  } catch {
    throw new Error(errorMessage);
  }
}

async function asBuffer(value: Promise<ArrayBuffer | Blob | Buffer>): Promise<Buffer> {
  const resolved = await value;

  if (Buffer.isBuffer(resolved)) {
    return resolved;
  }

  if (resolved instanceof Blob) {
    return Buffer.from(await resolved.arrayBuffer());
  }

  return Buffer.from(resolved);
}

function buildInlineHtml(entries: Record<string, Uint8Array>): string {
  const indexEntry = entries["index.html"];
  if (!indexEntry) {
    throw new Error("El ZIP exportado no contiene index.html");
  }

  const assets = new Map<string, Uint8Array>();
  for (const [entryPath, content] of Object.entries(entries)) {
    assets.set(normalizeAssetPath(entryPath), content);
  }

  let html = decodeText(indexEntry);
  html = stripScripts(html);
  html = neutralizeExternalImages(html);
  html = inlineStylesheets(html, assets);
  html = inlineImages(html, assets);

  return html;
}

function stripScripts(html: string): string {
  return html.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, "");
}

function inlineStylesheets(html: string, assets: Map<string, Uint8Array>): string {
  return html.replace(/<link\b([^>]*?)href=(["'])([^"']+)\2([^>]*?)>/gi, (full, before, _quote, href, after) => {
    const rel = `${before} ${after}`.toLowerCase();
    if (!rel.includes("stylesheet")) {
      return full;
    }

    const asset = assets.get(normalizeAssetPath(href));
    if (!asset) {
      return "";
    }

    const css = rewriteCssUrls(decodeText(asset), assets);
    return `<style data-source="${escapeHtmlAttribute(href)}">\n${css}\n</style>`;
  });
}

function inlineImages(html: string, assets: Map<string, Uint8Array>): string {
  return html.replace(/\b(src|poster)=(["'])([^"']+)\2/gi, (full, attr, _quote, value) => {
    if (isExternalUrl(value)) {
      return `${attr}=""`;
    }

    if (value.startsWith("data:") || value.startsWith("#")) {
      return full;
    }

    const asset = assets.get(normalizeAssetPath(value));
    if (!asset) {
      return full;
    }

    return `${attr}="data:${getMimeType(value)};base64,${toBase64(asset)}"`;
  });
}

function neutralizeExternalImages(html: string): string {
  return html.replace(/<img\b[^>]*>/gi, tag => {
    const src = getAttributeValue(tag, "src");
    if (!src || !isExternalUrl(src)) {
      return tag;
    }

    const alt = getAttributeValue(tag, "alt")?.trim();
    const content = alt || "Imagen externa no incrustada";
    return `<span>${escapeHtmlText(content)}</span>`;
  });
}

function rewriteCssUrls(css: string, assets: Map<string, Uint8Array>): string {
  return css.replace(/url\(([^)]+)\)/gi, (full, rawValue) => {
    const value = rawValue.trim().replace(/^['"]|['"]$/g, "");
    if (!value || isExternalUrl(value) || value.startsWith("data:") || value.startsWith("#")) {
      return full;
    }

    const asset = assets.get(normalizeAssetPath(value));
    if (!asset) {
      return full;
    }

    return `url("data:${getMimeType(value)};base64,${toBase64(asset)}")`;
  });
}

function normalizeAssetPath(assetPath: string): string {
  return assetPath
    .trim()
    .replace(/\\/g, "/")
    .replace(/^[.][/]+/, "")
    .replace(/^\//, "")
    .replace(/[?#].*$/, "");
}

function decodeText(content: Uint8Array): string {
  return new TextDecoder().decode(content);
}

function toBase64(content: Uint8Array): string {
  return Buffer.from(content).toString("base64");
}

function isExternalUrl(value: string): boolean {
  return /^(?:[a-z]+:)?\/\//i.test(value) || value.startsWith("mailto:");
}

function escapeHtmlAttribute(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll('"', "&quot;");
}

function escapeHtmlText(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function getAttributeValue(tag: string, attributeName: string): string | null {
  const match = tag.match(new RegExp(`\\b${attributeName}=(["'])(.*?)\\1`, "i"));
  return match?.[2] ?? null;
}

function getMimeType(filePath: string): string {
  const extension = path.extname(filePath).toLowerCase();

  switch (extension) {
    case ".css":
      return "text/css";
    case ".gif":
      return "image/gif";
    case ".ico":
      return "image/x-icon";
    case ".jpg":
    case ".jpeg":
      return "image/jpeg";
    case ".otf":
      return "font/otf";
    case ".png":
      return "image/png";
    case ".svg":
      return "image/svg+xml";
    case ".ttf":
      return "font/ttf";
    case ".webp":
      return "image/webp";
    case ".woff":
      return "font/woff";
    case ".woff2":
      return "font/woff2";
    default:
      return "application/octet-stream";
  }
}

function sanitizeInputFilename(filename: string): string {
  const baseName = path.basename(filename || "document.elpx");
  if (baseName.endsWith(".elpx") || baseName.endsWith(".elp")) {
    return baseName;
  }
  return `${baseName}.elpx`;
}

function sanitizeOutputFilename(filename: string): string {
  const baseName = path.basename(filename);
  const stem = baseName.replace(/\.[^.]+$/, "");
  const safeStem = stem || "documento";
  return `${safeStem}.docx`;
}

async function runCommand(
  command: string,
  args: string[],
  options?: { cwd?: string },
): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: options?.cwd,
      stdio: ["ignore", "pipe", "pipe"],
    });

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", chunk => {
      stdout += chunk.toString();
    });

    child.stderr.on("data", chunk => {
      stderr += chunk.toString();
    });

    child.on("error", error => {
      reject(new Error(`No se pudo ejecutar ${command}: ${error.message}`));
    });

    child.on("close", code => {
      if (code === 0) {
        resolve();
        return;
      }

      const details = [stdout.trim(), stderr.trim()].filter(Boolean).join("\n");
      reject(
        new Error(
          details
            ? `Fallo al ejecutar ${command} ${args.join(" ")}:\n${details}`
            : `Fallo al ejecutar ${command} ${args.join(" ")} (exit code ${code ?? "desconocido"})`,
        ),
      );
    });
  });
}
