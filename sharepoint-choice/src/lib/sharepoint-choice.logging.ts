import { ErrorHandler } from "@angular/core";
import { App } from "./sharepoint-choice.settings";

type GrafanaLogEntry = [timestamp: string, line: string];

interface SharePointPageContext {
  userLoginName?: string;
  webTitle?: string;
  webAbsoluteUrl?: string;
}

interface GrafanaLoggerState {
  stream: Record<string, string>;
  values: GrafanaLogEntry[];
  pendingMessages: Set<string>;
  sending: boolean;
  retryCount: number;
  flushTimer: number | undefined;
  lastTimestampNs: bigint;
  configurationWarningShown: boolean;
}

type LoggerWindow = Window & {
  GrafanaLogging?: unknown;
  _spPageContextInfo?: SharePointPageContext;
};

const FLUSH_DELAY_MS = 50;
const REQUEST_TIMEOUT_MS = 15_000;
const RETRY_BASE_DELAY_MS = 1_000;
const MAX_RETRY_DELAY_MS = 30_000;
const MAX_RETRY_COUNT = 7;
const STACK_TRACE_LIMIT = 2_048;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isGrafanaLogEntry(value: unknown): value is GrafanaLogEntry {
  return Array.isArray(value) && typeof value[0] === "string" && typeof value[1] === "string";
}

function isGrafanaLoggerState(value: unknown): value is GrafanaLoggerState {
  if (!isRecord(value) || !isRecord(value["stream"])) return false;

  return (
    Object.values(value["stream"]).every((label) => typeof label === "string") &&
    Array.isArray(value["values"]) &&
    value["values"].every(isGrafanaLogEntry) &&
    value["pendingMessages"] instanceof Set &&
    typeof value["sending"] === "boolean" &&
    typeof value["retryCount"] === "number" &&
    (value["flushTimer"] === undefined || typeof value["flushTimer"] === "number") &&
    typeof value["lastTimestampNs"] === "bigint" &&
    typeof value["configurationWarningShown"] === "boolean"
  );
}

function normalizeStream(value: unknown): Record<string, string> {
  if (!isRecord(value)) return {};

  const stream: Record<string, string> = {};
  for (const [label, rawValue] of Object.entries(value)) {
    stream[label] = rawValue == null ? "" : String(rawValue);
  }
  return stream;
}

function createDefaultStream(): Record<string, string> {
  return {
    Environment: App.Release,
    System: "SharePoint-Choice",
    Hostname: window.location.origin,
    Username: "",
    UserAgent: navigator.userAgent,
    Language: navigator.language,
  };
}

function initialTimestamp(values: GrafanaLogEntry[]): bigint {
  let latest = 0n;
  for (const [timestamp] of values) {
    if (/^\d+$/.test(timestamp)) {
      const parsed = BigInt(timestamp);
      if (parsed > latest) latest = parsed;
    }
  }
  return latest;
}

function getSharedLoggerState(): GrafanaLoggerState {
  const browserWindow = window as LoggerWindow;
  const existing = browserWindow.GrafanaLogging;
  if (isGrafanaLoggerState(existing)) return existing;

  const existingObject = isRecord(existing) ? existing : undefined;
  const values = Array.isArray(existingObject?.["values"])
    ? existingObject["values"].filter(isGrafanaLogEntry)
    : [];
  const configuredStream = normalizeStream(existingObject?.["stream"] ?? App.Grafana);
  const stream = Object.keys(configuredStream).length > 0 ? configuredStream : createDefaultStream();
  const state: GrafanaLoggerState = {
    stream,
    values,
    pendingMessages: new Set(values.map(([, line]) => line)),
    sending: false,
    retryCount: 0,
    flushTimer: undefined,
    lastTimestampNs: initialTimestamp(values),
    configurationWarningShown: false,
  };

  browserWindow.GrafanaLogging = state;
  return state;
}

function errorMessage(value: unknown): string {
  if (value instanceof Error) return value.message;
  if (typeof value === "string") return value;

  try {
    return JSON.stringify(value) ?? String(value);
  } catch (serializationError) {
    console.warn("Unable to serialize a SharePoint Choice error message.", serializationError);
    return String(value);
  }
}

export class SharepointChoiceLogging implements ErrorHandler {
  private readonly _grafana: GrafanaLoggerState;

  constructor() {
    this._grafana = getSharedLoggerState();
  }

  public handleError(error: unknown): void {
    void this.processError(error).catch((loggingError: unknown) => {
      console.error("SharepointChoiceLogging failed to process an error.", loggingError);
    });
  }

  private async processError(error: unknown): Promise<void> {
    const source = isRecord(error) ? (error["error"] ?? error) : error;
    const sourceObject = isRecord(source) ? source : undefined;
    const outerObject = isRecord(error) ? error : undefined;
    const levelValue = sourceObject?.["level"] ?? outerObject?.["level"];
    const level = typeof levelValue === "string" ? levelValue : "Unknown";
    const messageValue = outerObject?.["message"] ?? sourceObject?.["message"] ?? source;
    const message = errorMessage(messageValue);

    const stackValue = sourceObject?.["stack"] ?? outerObject?.["stack"];
    const rawStackTrace = typeof stackValue === "string" ? stackValue : "";
    const stackTrace =
      rawStackTrace.length >= STACK_TRACE_LIMIT
        ? `${rawStackTrace.substring(0, STACK_TRACE_LIMIT - 3)}...`
        : rawStackTrace;

    this.updateSharePointContext();

    const pathSegments = window.location.pathname.split("/");
    const pathMatch = window.location.pathname.match(/\/[^\/]*\/[^\/]*\/([^?]*)/);
    // Accepted privacy and payload-size risk: page query/hash values, user identifiers, errors, and stack traces
    // are sent without redaction or a total payload-size limit; callers should not put secrets in URLs or errors.
    const Params = {
      prefix: pathSegments[1] ?? "",
      site: pathSegments[2] ?? "",
      path: pathMatch?.[1] ?? "",
      search: window.location.search,
      hash: window.location.hash,
    };

    switch (level.substring(0, 1).toUpperCase()) {
      case "E":
        console.error(error);
        break;
      case "W":
        console.warn(error);
        break;
      case "I":
        console.info(error);
        break;
      case "D":
        console.log(error);
        break;
      default:
        console.trace(error);
        break;
    }

    const line = JSON.stringify({
      Params,
      Level: level,
      Message: message,
      StackTrace: stackTrace,
    });

    if (this._grafana.pendingMessages.has(line)) return;

    const wallClockTimestamp = BigInt(Date.now()) * 1_000_000n;
    const timestamp = wallClockTimestamp > this._grafana.lastTimestampNs
      ? wallClockTimestamp
      : this._grafana.lastTimestampNs + 1n;
    this._grafana.lastTimestampNs = timestamp;

    this._grafana.values.push([timestamp.toString(), line]);
    this._grafana.pendingMessages.add(line);
    this.scheduleFlush(FLUSH_DELAY_MS);
  }

  private updateSharePointContext(): void {
    const pageContext = (window as LoggerWindow)._spPageContextInfo;
    if (!pageContext) return;

    const replacements: Record<string, string | undefined> = {
      Username: pageContext.userLoginName,
      System: pageContext.webTitle,
      Hostname: pageContext.webAbsoluteUrl,
    };

    for (const label of Object.keys(this._grafana.stream)) {
      for (const [name, value] of Object.entries(replacements)) {
        if (label.includes(name) && value) this._grafana.stream[label] = value;
      }
    }
  }

  private scheduleFlush(delay: number): void {
    if (this._grafana.sending || this._grafana.flushTimer !== undefined || this._grafana.values.length === 0) {
      return;
    }

    this._grafana.flushTimer = window.setTimeout(() => {
      this._grafana.flushTimer = undefined;
      void this.sendErrors();
    }, delay);
  }

  private async sendErrors(): Promise<void> {
    if (this._grafana.sending || this._grafana.values.length === 0) return;

    // Remote logging is skipped on localhost. Without a release endpoint, pending records are warned about and
    // discarded; there is no durable offline store.
    if (window.location.hostname === "localhost") {
      this.discardPendingLogs();
      return;
    }

    const configuredEndpoints: unknown = App.ApiServers?.["grafana"];
    const endpoint = isRecord(configuredEndpoints) ? configuredEndpoints[App.Release] : undefined;
    if (typeof endpoint !== "string" || endpoint.length === 0) {
      if (!this._grafana.configurationWarningShown) {
        console.warn(`Grafana logging is not configured for release ${App.Release}; pending logs will be discarded.`);
        this._grafana.configurationWarningShown = true;
      }
      this.discardPendingLogs();
      return;
    }

    const batch = this._grafana.values.slice();
    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
    this._grafana.sending = true;

    let retryDelay: number | undefined;
    try {
      const body = JSON.stringify({
        streams: [{ stream: this._grafana.stream, values: batch }],
      });
      const response = await fetch(endpoint, {
        method: "POST",
        body,
        headers: {
          "Content-Type": "application/json",
        },
        signal: controller.signal,
      });

      if (!response.ok) {
        throw new Error(`Grafana returned HTTP ${response.status} ${response.statusText}.`);
      }

      this._grafana.values.splice(0, batch.length);
      for (const [, line] of batch) this._grafana.pendingMessages.delete(line);
      this._grafana.retryCount = 0;
      console.info(`Submitted ${batch.length} logs to Grafana.`);
    } catch (error) {
      this._grafana.retryCount = Math.min(this._grafana.retryCount + 1, MAX_RETRY_COUNT);
      retryDelay = Math.min(
        RETRY_BASE_DELAY_MS * 2 ** Math.min(this._grafana.retryCount - 1, MAX_RETRY_COUNT),
        MAX_RETRY_DELAY_MS,
      );
      console.error(`Error logging ${batch.length} records to Grafana; retrying in ${retryDelay}ms.`, error);
    } finally {
      window.clearTimeout(timeout);
      this._grafana.sending = false;
    }

    if (retryDelay !== undefined) this.scheduleFlush(retryDelay);
    else if (this._grafana.values.length > 0) this.scheduleFlush(FLUSH_DELAY_MS);
  }

  private discardPendingLogs(): void {
    this._grafana.values.length = 0;
    this._grafana.pendingMessages.clear();
  }
}
