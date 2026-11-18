/**
 * Tiny structured logger. Writes JSON lines to **stderr** — stdout is reserved for the
 * MCP protocol stream on a stdio transport, so nothing chatty may ever touch it.
 */

export type LogLevel = "silent" | "error" | "warn" | "info" | "debug";

const ORDER: Record<LogLevel, number> = { silent: 0, error: 1, warn: 2, info: 3, debug: 4 };
const LEVELS = new Set<string>(["silent", "error", "warn", "info", "debug"]);

// Emit only operational metadata. Arbitrary provider messages, task content, paths,
// nested objects and HTTP headers have no place in a local security log.
function safeFields(fields: LogFields): Record<string, unknown> {
  const safe: Record<string, unknown> = {};
  for (const key of ["durationMs", "status", "attempt", "waitMs", "registeredTools", "skippedTools"]) {
    if (typeof fields[key] === "number" && Number.isFinite(fields[key])) safe[key] = fields[key];
  }
  for (const key of ["clientIdConfigured", "readonly", "noDestructive", "scopeReadonly"]) {
    if (typeof fields[key] === "boolean") safe[key] = fields[key];
  }
  if (fields.outcome === "ok" || fields.outcome === "error") safe.outcome = fields.outcome;
  return safe;
}

export interface LogFields {
  traceId?: string;
  clientRequestId?: string;
  requestId?: string;
  tool?: string;
  op?: string;
  durationMs?: number;
  status?: number;
  attempt?: number;
  waitMs?: number;
  outcome?: "ok" | "error";
  errorCode?: string;
  [key: string]: unknown;
}

export class Logger {
  constructor(private readonly level: LogLevel, private readonly base: LogFields = {}) {}

  child(fields: LogFields): Logger {
    return new Logger(this.level, { ...this.base, ...fields });
  }

  error(msg: string, fields?: LogFields): void { this.write("error", msg, fields); }
  warn(msg: string, fields?: LogFields): void { this.write("warn", msg, fields); }
  info(msg: string, fields?: LogFields): void { this.write("info", msg, fields); }
  debug(msg: string, fields?: LogFields): void { this.write("debug", msg, fields); }

  private write(level: Exclude<LogLevel, "silent">, msg: string, fields: LogFields = {}): void {
    if (ORDER[level] > ORDER[this.level]) return;
    const line = { level, time: new Date().toISOString(), msg, ...safeFields({ ...this.base, ...fields }) };
    process.stderr.write(JSON.stringify(line) + "\n");
  }
}

export function createLogger(level: string | undefined): Logger {
  const lvl = (level && LEVELS.has(level) ? level : "info") as LogLevel;
  return new Logger(lvl);
}
