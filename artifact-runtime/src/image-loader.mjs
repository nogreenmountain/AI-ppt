import { readFile } from "node:fs/promises";
import { basename, extname, isAbsolute, resolve } from "node:path";

export function resolveAssetPath(assetPath, baseDir = process.cwd()) {
  if (!assetPath) return null;
  return isAbsolute(assetPath) ? assetPath : resolve(baseDir, assetPath);
}

export function imageContentType(assetPath = "") {
  const ext = extname(assetPath).toLowerCase();
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg";
  if (ext === ".webp") return "image/webp";
  if (ext === ".gif") return "image/gif";
  return "image/png";
}

export async function loadImageBytes(assetPath, baseDir = process.cwd()) {
  const path = resolveAssetPath(assetPath, baseDir);
  const bytes = new Uint8Array(await readFile(path));
  return {
    bytes,
    path,
    name: basename(path),
    contentType: imageContentType(path),
  };
}

export function imageDataUri(bytes, contentType = "image/png") {
  return `data:${contentType};base64,${Buffer.from(bytes).toString("base64")}`;
}
