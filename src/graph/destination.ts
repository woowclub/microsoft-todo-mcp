import { GRAPH_BASE } from "../config";
import { AppError } from "./errors";

const GRAPH_ORIGIN = new URL(GRAPH_BASE).origin;
const MAX_URL_LENGTH = 8192;
const id = "(?:[A-Za-z0-9_+=-]|%2B|%3D){1,1024}";
const route = new RegExp(`^/v1\\.0/me/todo/lists(?:/${id}(?:/tasks(?:/${id}(?:/checklistItems(?:/${id})?)?)?)?)?$`);

function invalid(): never {
  throw new AppError("validation_error", "Invalid Microsoft To Do request destination.");
}

/** Validate the exact URL handed to fetch. Never use URL resolution with attacker-controlled paths. */
export function graphUrl(input: string): URL {
  if (typeof input !== "string" || input.length > MAX_URL_LENGTH || /[^!-~]|[\\#]|%(?![\da-f]{2})/i.test(input)) invalid();
  let url: URL;
  try {
    url = new URL(input);
  } catch {
    return invalid();
  }
  // Comparing serialization also rejects paths the URL parser repaired or normalized (e.g. dot segments).
  if (url.origin !== GRAPH_ORIGIN || url.username || url.password || url.href !== input || !route.test(url.pathname)) invalid();
  const seen = new Set<string>();
  for (const [name, value] of url.searchParams) {
    if (seen.has(name.toLowerCase()) || !validQuery(name, value)) invalid();
    seen.add(name.toLowerCase());
  }
  return url;
}

function validQuery(name: string, value: string): boolean {
  switch (name.toLowerCase()) {
    case "$select": return value === "id,isOwner";
    case "$skiptoken": return /^[!-~]+$/.test(value);
    case "$skip": return /^\d{1,10}$/.test(value);
    case "$top": return /^(?:[1-9]\d?|100)$/.test(value);
    case "$filter": return /^status eq '(?:completed|notStarted|inProgress)'$/.test(value);
    case "$expand": return /^(?:checklistItems(?:,linkedResources)?|linkedResources)$/.test(value);
    default: return false;
  }
}

export function pageUrl(input: string, collection: string): string {
  const url = graphUrl(input);
  // IDs may return with literal padding (+ or =) where the original adapter encoded it.
  const expected = graphUrl(GRAPH_BASE + collection);
  if (decodeURIComponent(url.pathname) !== decodeURIComponent(expected.pathname)) invalid();
  return url.href;
}
