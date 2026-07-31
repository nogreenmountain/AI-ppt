import { readFile } from "node:fs/promises";

export class SchemaError extends Error {
  constructor(message) {
    super(message);
    this.name = "SchemaError";
  }
}

function stripJsonComments(text) {
  return text
    .split(/\r?\n/)
    .filter((line) => !line.trimStart().startsWith("//"))
    .join("\n");
}

function requireObject(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new SchemaError(`${label} must be an object`);
  }
}

function requireNumber(value, label) {
  if (!Number.isFinite(Number(value))) {
    throw new SchemaError(`${label} must be a number`);
  }
}

function validateElement(el, index) {
  requireObject(el, `elements[${index}]`);
  if (!["text", "shape", "image"].includes(el.kind)) {
    throw new SchemaError(`elements[${index}].kind is invalid`);
  }
  if (el.z === undefined) {
    throw new SchemaError(`elements[${index}].z is required`);
  }
  requireObject(el.bbox, `elements[${index}].bbox`);
  for (const key of ["left", "top", "width", "height"]) {
    requireNumber(el.bbox[key], `elements[${index}].bbox.${key}`);
  }
  if (el.kind === "text" && typeof el.text !== "string") {
    throw new SchemaError(`elements[${index}].text is required`);
  }
  if (el.kind === "shape" && typeof el.geometry !== "string") {
    throw new SchemaError(`elements[${index}].geometry is required`);
  }
  if (el.kind === "image" && typeof el.image_path !== "string") {
    throw new SchemaError(`elements[${index}].image_path is required`);
  }
}

export function loadSchemaFromObject(schema) {
  requireObject(schema, "schema");
  if (schema.version !== "1.0") {
    throw new SchemaError("version must be 1.0");
  }
  requireObject(schema.slide, "slide");
  requireNumber(schema.slide.width, "slide.width");
  requireNumber(schema.slide.height, "slide.height");
  requireObject(schema.background, "background");
  if (!Array.isArray(schema.elements)) {
    throw new SchemaError("elements must be an array");
  }
  schema.elements.forEach(validateElement);
  return schema;
}

export async function loadSchemaFromPath(filePath) {
  const raw = await readFile(filePath, "utf8");
  return loadSchemaFromObject(JSON.parse(stripJsonComments(raw)));
}
